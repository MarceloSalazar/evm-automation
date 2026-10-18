# -*- mode: python ; coding: utf-8 -*-
"""Especificación de PyInstaller: evm_export.exe y generar_ejemplo.exe en la carpeta ExportadorEVM.

Se ejecuta con instalador\\construir.py, que prepara el entorno y luego compila el instalador.
"""
import re
from pathlib import Path

from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct, StringTable,
                                                 VarFileInfo, VarStruct, VSVersionInfo)

ROOT = Path(SPECPATH).parent
VERSION = re.search(r'__version__ = "([^"]+)"',
                    (ROOT / "evm" / "__init__.py").read_text(encoding="utf-8")).group(1)
NUMBERS = tuple(int(x) for x in VERSION.split(".")) + (0,)
EXCLUDES = ["tkinter", "unittest", "pydoc_data", "pytest"]


def version_file(name: str, description: str) -> str:
    """Recurso de versión de Windows (Propiedades > Detalles del .exe), en español."""
    info = VSVersionInfo(
        ffi=FixedFileInfo(filevers=NUMBERS, prodvers=NUMBERS),
        kids=[StringFileInfo([StringTable("0C0A04B0", [
                  StringStruct("CompanyName", "Marcelo Salazar"),
                  StringStruct("FileDescription", description),
                  StringStruct("FileVersion", VERSION),
                  StringStruct("InternalName", name),
                  StringStruct("LegalCopyright", "Copyright (c) 2026 Marcelo Salazar. Licencia MIT."),
                  StringStruct("OriginalFilename", f"{name}.exe"),
                  StringStruct("ProductName", "Exportador EVM"),
                  StringStruct("ProductVersion", VERSION)])]),
              VarFileInfo([VarStruct("Translation", [0x0C0A, 1200])])])
    path = Path(workpath) / f"version_{name}.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(str(info), encoding="utf-8")
    return str(path)


a_evm = Analysis([str(ROOT / "evm_export.py")], pathex=[str(ROOT)], excludes=EXCLUDES)
a_ejemplo = Analysis([str(ROOT / "ejemplos" / "generar_ejemplo.py")], pathex=[str(ROOT)], excludes=EXCLUDES)

exe_evm = EXE(PYZ(a_evm.pure), a_evm.scripts, [], exclude_binaries=True, name="evm_export", console=True,
              upx=False, version=version_file("evm_export", "Exportador EVM: de MS Project a Excel"))
exe_ejemplo = EXE(PYZ(a_ejemplo.pure), a_ejemplo.scripts, [], exclude_binaries=True, name="generar_ejemplo",
                  console=True, upx=False,
                  version=version_file("generar_ejemplo", "Exportador EVM: proyecto de ejemplo con avance simulado"))

COLLECT(exe_evm, a_evm.binaries, a_evm.datas,
        exe_ejemplo, a_ejemplo.binaries, a_ejemplo.datas,
        name="ExportadorEVM", upx=False)
