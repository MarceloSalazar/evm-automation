"""Versión del programa que se muestra en la consola, el registro de diagnóstico, el libro y el informe de mejoras.

Una versión publicada es la de evm.__version__ (p. ej. 2.6.0). Al ejecutar desde una copia del repositorio que no
está en una versión publicada, la versión indica la última versión publicada y el commit (hash abreviado, como lo
muestra GitHub): «2.5.0 + desarrollo (b513fbb)». Las versiones publicadas son las etiquetas vX.Y.Z de git.
"""
from __future__ import annotations

import functools
import re
import subprocess
import sys
from pathlib import Path

from evm import __version__

ROOT = Path(__file__).resolve().parents[1]
# Salida de git describe --long: v2.5.0-19-gb513fbb (versión, commits desde ella y commit)
_DESCRIBE = re.compile(r"v(?P<tag>\d+(?:\.\d+)*)-(?P<ahead>\d+)-g(?P<commit>[0-9a-f]+)")


def describe(text: str) -> str | None:
    """Versión a partir de la salida de git describe --long (None si no la reconoce)."""
    m = _DESCRIBE.fullmatch(text.strip())
    if m is None:
        return None
    if m["ahead"] == "0":
        return m["tag"]
    return f"{m['tag']} + desarrollo ({m['commit']})"


@functools.cache
def version_label() -> str:
    """La versión que se muestra: la publicada (__version__) o, desde una copia del repositorio que no está en una
    versión publicada, «X.Y.Z + desarrollo (commit)». Sin git, en el ejecutable del instalador o si git no responde,
    __version__."""
    if getattr(sys, "frozen", False) or not (ROOT / ".git").exists():
        return __version__
    try:
        out = subprocess.run(["git", "-C", str(ROOT), "describe", "--tags", "--long", "--abbrev=7",
                              "--match", "v[0-9]*"], capture_output=True, text=True, timeout=10,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError):
        return __version__
    return (describe(out.stdout) if out.returncode == 0 else None) or __version__
