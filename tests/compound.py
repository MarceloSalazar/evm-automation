"""Archivos compuestos de OLE (el formato del .mpp) para las pruebas de evm.mpp_macros.

compound_file({"VBA": {"dir": b"..."}, "PROJECT": b"Module=Modulo1"}) crea un archivo de versión 3 (sectores de 512
bytes) con esos almacenamientos (dict) y flujos (bytes, de menos de 4096 bytes: van al mini flujo, como los flujos
pequeños de un .mpp).
"""
from __future__ import annotations

import struct

SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
FREE, END, FATSECT, NONE = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
SECTOR, MINI = 512, 64


def _sectors(n_bytes: int, size: int = SECTOR) -> int:
    return -(-n_bytes // size)


def compound_file(tree: dict) -> bytes:
    entries = [{"name": "Root Entry", "type": 5, "children": [], "start": END, "size": 0}]
    mini = bytearray()
    minifat: list[int] = []

    def add(node: dict, parent: int) -> None:
        for name, value in node.items():
            index = len(entries)
            entries[parent]["children"].append(index)
            if isinstance(value, dict):
                entries.append({"name": name, "type": 1, "children": [], "start": 0, "size": 0})
                add(value, index)
                continue
            assert len(value) < 4096, "solo flujos pequeños (mini flujo)"
            n = _sectors(len(value), MINI)
            start = len(minifat) if n else END
            minifat.extend([len(minifat) + k + 1 for k in range(n - 1)] + ([END] if n else []))
            mini.extend(value + b"\0" * (n * MINI - len(value)))
            entries.append({"name": name, "type": 2, "children": [], "start": start, "size": len(value)})

    add(tree, 0)
    n_dir = _sectors(len(entries) * 128)
    n_minifat = _sectors(len(minifat) * 4) if minifat else 0
    n_mini = _sectors(len(mini))
    first_dir, first_minifat = 1, 1 + n_dir
    first_mini = first_minifat + n_minifat
    total = first_mini + n_mini
    assert total <= SECTOR // 4, "archivo de prueba demasiado grande para una sola tabla FAT"
    fat = [FATSECT]
    for first, n in ((first_dir, n_dir), (first_minifat, n_minifat), (first_mini, n_mini)):
        fat += [first + k + 1 for k in range(n - 1)] + ([END] if n else [])
    fat += [FREE] * (SECTOR // 4 - len(fat))
    if n_mini:
        entries[0]["start"], entries[0]["size"] = first_mini, len(mini)

    for e in entries:                               # hermanos encadenados por la derecha
        for a, b in zip(e["children"], e["children"][1:]):
            entries[a]["right"] = b
    directory = bytearray()
    for e in entries:
        kids = e["children"]
        name = e["name"].encode("utf-16-le") + b"\0\0"
        record = bytearray(128)
        record[:len(name)] = name
        struct.pack_into("<HBB", record, 0x40, len(name), e["type"], 1)
        struct.pack_into("<III", record, 0x44, NONE, e.get("right", NONE), kids[0] if kids else NONE)
        struct.pack_into("<IQ", record, 0x74, e["start"], e["size"])
        directory += record
    directory += b"\0" * (n_dir * SECTOR - len(directory))

    header = bytearray(SECTOR)
    header[:8] = SIGNATURE
    struct.pack_into("<HHHHH", header, 0x18, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIIII", header, 0x28, 0, 1, first_dir, 0, 4096, first_minifat if n_minifat else END,
                     n_minifat, END, 0, 0)
    struct.pack_into("<108I", header, 0x50, *([FREE] * 108))
    minifat_bytes = struct.pack(f"<{len(minifat)}I", *minifat) if minifat else b""
    minifat_bytes += struct.pack("<I", FREE) * ((n_minifat * SECTOR - len(minifat_bytes)) // 4)
    body = struct.pack(f"<{len(fat)}I", *fat) + bytes(directory) + minifat_bytes
    body += bytes(mini) + b"\0" * (n_mini * SECTOR - len(mini))
    return bytes(header) + body


def project_stream(modules: list[str], classes: list[str] = ()) -> bytes:
    """Flujo PROJECT de un proyecto VBA de MS Project (texto en la página de códigos de Windows)."""
    lines = ['ID="{41C5E297-0000-0000-0000-000000000000}"', "Document=ThisProject/&H00000000"]
    lines += [f"Module={m}" for m in modules] + [f"Class={c}" for c in classes]
    lines += ['Name="VBAProject"', 'HelpContextID="0"', "", "[Workspace]", "ThisProject=38, 38, 1817, 1017, C"]
    return "\r\n".join(lines).encode("cp1252")


def mpp_with_macros(modules: list[str], classes: list[str] = ()) -> bytes:
    """Archivo con la estructura de un .mpp con macros: el proyecto VBA en un almacenamiento de MS Project."""
    vba = {"dir": b"\x01" + b"x" * 40, "ThisProject": b"\x00" * 80, "_VBA_PROJECT": b"\xcc\x61" + b"\0" * 6}
    vba.update({m: b"\x00" * 100 for m in [*modules, *classes]})
    return compound_file({
        "   114": {"Props": b"\0" * 100, "TBkndTask": {"VarMeta": b"\0" * 64}},
        "   214": {"CVba": {"Props": b"\0" * 10,
                            "2b400008_ffffffff": {"VBA": vba, "PROJECT": project_stream(modules, classes),
                                                  "PROJECTwm": b"\0" * 8}}},
        "\x05SummaryInformation": b"\0" * 48,
    })
