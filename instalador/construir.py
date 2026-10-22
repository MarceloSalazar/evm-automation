"""Genera el instalador de Windows del Exportador EVM (PyInstaller + Inno Setup).

Pasos:
1. Crea (la primera vez) un entorno virtual en build\\instalador\\venv con las dependencias del programa
   y PyInstaller.
2. Genera con PyInstaller la carpeta ExportadorEVM con evm_export.exe y generar_ejemplo.exe
   (instalador\\ExportadorEVM.spec).
3. Prueba los ejecutables con el proyecto de ejemplo: generar, avance simulado, actualizar y analizar.
4. Compila con Inno Setup dist\\ExportadorEVM-X.Y.Z-instalador.exe (instalador\\ExportadorEVM.iss).

Requisitos: Windows con Python 3.11 o posterior (el instalador incluye esa versión de Python), conexión a
internet la primera vez (pip) e Inno Setup 6 (winget install JRSoftware.InnoSetup), o la ruta de
ISCC.exe en la variable de entorno ISCC.

Uso (desde la carpeta raíz del repositorio, en el commit etiquetado de la versión):
    python instalador\\construir.py
    gh release upload vX.Y.Z dist\\ExportadorEVM-X.Y.Z-instalador.exe
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HERE = ROOT / "instalador"
BUILD = ROOT / "build" / "instalador"
VENV = BUILD / "venv"
OUTPUT = ROOT / "dist"
PYINSTALLER = "pyinstaller>=6.10,<7"


def version() -> str:
    text = (ROOT / "evm" / "__init__.py").read_text(encoding="utf-8")
    return re.search(r'__version__ = "([^"]+)"', text).group(1)


def run(args: list, **kw) -> subprocess.CompletedProcess:
    print(">", " ".join(str(a) for a in args), flush=True)
    return subprocess.run([str(a) for a in args], check=True, **kw)


def git(*args: str) -> str:
    result = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True)
    return result.stdout.strip() if result.returncode == 0 else ""


def check_git(ver: str) -> None:
    """Avisa si el instalador no se genera desde la versión publicada (no lo impide)."""
    if git("status", "--porcelain", "--untracked-files=no"):
        print(f"[!] Hay cambios sin confirmar: el instalador no corresponderá exactamente a v{ver}.")
    if git("describe", "--tags", "--exact-match") != f"v{ver}":
        print(f"[!] El commit actual no tiene la etiqueta v{ver}: el instalador se debe generar desde la "
              "versión publicada.")


def venv_python() -> Path:
    python = VENV / "Scripts" / "python.exe"
    if not python.exists():
        run([sys.executable, "-m", "venv", VENV])
    run([python, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
         "-r", ROOT / "requirements.txt", PYINSTALLER])
    return python


def smoke_test(folder: Path, ver: str) -> None:
    """Ciclo del proyecto de ejemplo con los ejecutables generados, en una carpeta temporal."""
    evm, ejemplo = folder / "evm_export.exe", folder / "generar_ejemplo.exe"
    out = run([evm, "--version"], capture_output=True, text=True).stdout
    if ver not in out:
        raise SystemExit(f"evm_export --version no muestra {ver}: {out!r}")
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        config = tmp / "config.toml"
        text = (ROOT / "config.ejemplo.toml").read_text(encoding="utf-8")
        text = text.replace('metodo = ""', 'metodo = "Método EVM"').replace('peso_hito = ""', 'peso_hito = "Peso hito"')
        config.write_text(text, encoding="utf-8")
        libro, avance = tmp / "ejemplo.xlsx", tmp / "avance.xml"
        quiet = {"stdout": subprocess.DEVNULL}
        run([evm, "generar", "--origen", ROOT / "ejemplos" / "programa_ejemplo.xml", "--config", config,
             "--salida", libro], **quiet)
        run([ejemplo, "--libro", libro], **quiet)                     # BAC del ejemplo (modo porcentaje)
        run([ejemplo, "--fecha-estado", "2027-06-30", "--salida", avance, "--libro", libro], **quiet)
        run([evm, "actualizar", "--origen", avance, "--config", config, "--libro", libro, "--sin-respaldo"], **quiet)
        run([evm, "analizar", "--libro", libro, "--sin-respaldo"], **quiet)
    print("Prueba de los ejecutables: correcta.")


def find_iscc() -> Path:
    candidates = [os.environ.get("ISCC"), shutil.which("iscc")]
    for base in ("LOCALAPPDATA", "ProgramFiles(x86)", "ProgramFiles"):
        if os.environ.get(base):
            sub = "Programs/Inno Setup 6" if base == "LOCALAPPDATA" else "Inno Setup 6"
            candidates.append(str(Path(os.environ[base]) / sub / "ISCC.exe"))
    for c in candidates:
        if c and Path(c).exists():
            return Path(c)
    raise SystemExit("No se encontró Inno Setup 6 (ISCC.exe). Se instala con: winget install JRSoftware.InnoSetup "
                     "(o se indica su ruta en la variable de entorno ISCC).")


def main() -> None:
    if sys.platform != "win32":
        raise SystemExit("El instalador se genera en Windows.")
    ver = version()
    print(f"Exportador EVM {ver}: instalador de Windows")
    check_git(ver)
    iscc = find_iscc()
    python = venv_python()
    run([python, "-m", "PyInstaller", "--noconfirm", "--clean", "--log-level", "WARN",
         "--distpath", BUILD / "dist", "--workpath", BUILD / "work", HERE / "ExportadorEVM.spec"])
    folder = BUILD / "dist" / "ExportadorEVM"
    smoke_test(folder, ver)
    OUTPUT.mkdir(exist_ok=True)
    run([iscc, "/Qp", f"/DAppVersion={ver}", f"/DOrigen={folder}", f"/O{OUTPUT}", HERE / "ExportadorEVM.iss"])
    setup = OUTPUT / f"ExportadorEVM-{ver}-instalador.exe"
    digest = hashlib.sha256(setup.read_bytes()).hexdigest()
    print(f"\nInstalador: {setup} ({setup.stat().st_size / 1e6:.1f} MB)\nSHA-256: {digest}")


if __name__ == "__main__":
    main()
