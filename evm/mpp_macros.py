"""Macros (módulos VBA) guardadas en un archivo .mpp, leídas sin MS Project.

Un .mpp es un archivo compuesto de OLE (Compound File Binary, [MS-CFB]): un pequeño sistema de archivos con
almacenamientos y flujos. Si el archivo tiene macros, uno de sus almacenamientos tiene el proyecto VBA: un
almacenamiento VBA con el código y un flujo PROJECT de texto que enumera los módulos ([MS-OVBA] 2.3.1):

    Document=ThisProject/&H00000000
    Module=Módulo1
    Class=Clase1
    BaseClass=UserForm1

generar-cliente lo usa para saber qué módulos tiene el archivo de origen y para comprobar que el archivo del cliente,
ya guardado, no tiene ninguno. ThisProject (el módulo del documento) siempre existe y no se cuenta.
"""
from __future__ import annotations

import struct
from pathlib import Path

_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
_END = 0xFFFFFFFA               # a partir de aquí, valores especiales de la cadena de sectores (fin, libre...)
_NONE = 0xFFFFFFFF              # sin entrada de directorio
_STORAGE, _STREAM, _ROOT = 1, 2, 5
_MODULE_KEYS = ("Module", "Class", "BaseClass")


class _Unreadable(Exception):
    """El archivo no es un archivo compuesto o está dañado (interno: macro_modules devuelve None)."""


class _Compound:
    """Lector mínimo de archivos compuestos: directorio y contenido de los flujos."""

    def __init__(self, data: bytes):
        if data[:8] != _SIGNATURE or len(data) < 512:
            raise _Unreadable("no es un archivo compuesto")
        self.data = data
        major = struct.unpack_from("<H", data, 0x1A)[0]
        self.sector = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector = 1 << struct.unpack_from("<H", data, 0x20)[0]
        n_fat, dir_start = struct.unpack_from("<II", data, 0x2C)
        self.cutoff, minifat_start = struct.unpack_from("<II", data, 0x38)
        difat_start = struct.unpack_from("<I", data, 0x44)[0]
        per_sector = self.sector // 4
        difat = list(struct.unpack_from("<109I", data, 0x4C))
        seen: set[int] = set()
        while difat_start < _END and difat_start not in seen:
            seen.add(difat_start)
            values = struct.unpack_from(f"<{per_sector}I", data, self._offset(difat_start))
            difat += values[:-1]
            difat_start = values[-1]
        self.fat: list[int] = []
        for sid in difat[:n_fat]:
            if sid < _END:
                self.fat += struct.unpack_from(f"<{per_sector}I", data, self._offset(sid))
        directory = self._chain_bytes(dir_start)
        self.entries = []
        for i in range(len(directory) // 128):
            e = directory[i * 128:(i + 1) * 128]
            size_name = struct.unpack_from("<H", e, 0x40)[0]
            name = e[:max(min(size_name, 64) - 2, 0)].decode("utf-16-le", "replace")
            left, right, child = struct.unpack_from("<III", e, 0x44)
            start, size = struct.unpack_from("<IQ", e, 0x74)
            if major == 3:              # versión 3: solo valen los 32 bits bajos del tamaño
                size &= 0xFFFFFFFF
            self.entries.append({"name": name, "type": e[0x42], "left": left, "right": right, "child": child,
                                 "start": start, "size": size})
        root = self.entries[0] if self.entries else None
        if root is None or root["type"] != _ROOT:
            raise _Unreadable("archivo compuesto sin directorio")
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]]
        self.minifat: list[int] = []
        for sid in self._chain(minifat_start):
            self.minifat += struct.unpack_from(f"<{per_sector}I", data, self._offset(sid))

    def _offset(self, sid: int) -> int:
        offset = (sid + 1) * self.sector
        if offset + self.sector > len(self.data):
            raise _Unreadable("sector fuera del archivo")
        return offset

    def _chain(self, start: int, table: list[int] | None = None) -> list[int]:
        table = self.fat if table is None else table
        out: list[int] = []
        seen: set[int] = set()
        while start < _END:
            if start in seen or start >= len(table):
                raise _Unreadable("cadena de sectores dañada")
            seen.add(start)
            out.append(start)
            start = table[start]
        return out

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self.data[self._offset(s):self._offset(s) + self.sector] for s in self._chain(start))

    def children(self, index: int) -> list[int]:
        """Entradas de un almacenamiento (su árbol de hermanos)."""
        out: list[int] = []
        stack = [self.entries[index]["child"]]
        seen: set[int] = set()
        while stack:
            i = stack.pop()
            if i == _NONE or i in seen or i >= len(self.entries):
                continue
            seen.add(i)
            e = self.entries[i]
            out.append(i)
            stack += [e["left"], e["right"]]
        return out

    def read(self, index: int) -> bytes:
        e = self.entries[index]
        if e["size"] < self.cutoff:
            parts = [self.mini_stream[s * self.mini_sector:(s + 1) * self.mini_sector]
                     for s in self._chain(e["start"], self.minifat)]
            return b"".join(parts)[:e["size"]]
        return self._chain_bytes(e["start"])[:e["size"]]


def _project_streams(cf: _Compound) -> list[bytes]:
    """Flujos PROJECT de los proyectos VBA del archivo: los de un almacenamiento que también tiene uno llamado VBA."""
    out: list[bytes] = []
    stack, seen = [0], set()
    while stack:
        index = stack.pop()
        if index in seen:
            continue
        seen.add(index)
        kids = cf.children(index)
        names = {cf.entries[i]["name"].upper(): i for i in kids}
        vba, project = names.get("VBA"), names.get("PROJECT")
        if vba is not None and project is not None and cf.entries[vba]["type"] == _STORAGE \
                and cf.entries[project]["type"] == _STREAM:
            out.append(cf.read(project))
        stack += [i for i in kids if cf.entries[i]["type"] == _STORAGE]
    return out


def modules_in_project_stream(text: bytes) -> list[str]:
    """Módulos (Module, Class y BaseClass; no el del documento) de un flujo PROJECT."""
    try:
        decoded = text.decode("cp1252")
    except UnicodeDecodeError:
        decoded = text.decode("latin-1")
    out = []
    for line in decoded.splitlines():
        key, sep, value = line.partition("=")
        if sep and key.strip() in _MODULE_KEYS and value.strip():
            out.append(value.strip())
    return out


def macro_modules(path: str | Path) -> list[str] | None:
    """Módulos de macros del archivo (sin ThisProject): [] si no tiene; None si no se pudo leer (no es un archivo
    compuesto, está dañado o no se puede abrir)."""
    try:
        cf = _Compound(Path(path).read_bytes())
        modules: list[str] = []
        for stream in _project_streams(cf):
            modules += [m for m in modules_in_project_stream(stream) if m not in modules]
        return modules
    except (OSError, _Unreadable, struct.error, IndexError, ValueError):
        return None
