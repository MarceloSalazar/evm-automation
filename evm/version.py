"""Versión del programa que se muestra en la consola, el registro de diagnóstico, el libro y el informe de mejoras.

Una versión publicada es la de evm.__version__ (p. ej. 2.6.0). Al ejecutar desde una copia del repositorio que no
está en una versión publicada, la versión indica la última versión publicada y el commit (hash abreviado, como lo
muestra GitHub): «2.5.0 + desarrollo (b513fbb)». Las versiones publicadas son las etiquetas vX.Y.Z de git.

La etiqueta de una versión queda en el commit de fusión de su publicación en main, que puede no estar en el historial
de la rama de trabajo (desarrollo o una rama que sale de ella): por eso la versión de desarrollo toma la etiqueta más
alta que no supera __version__, aunque git describe encuentre una anterior.
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


def _key(version: str) -> tuple[int, ...]:
    return tuple(int(x) for x in version.split("."))


def latest_release(described: str, tags: list[str], current: str = __version__) -> str:
    """Versión publicada más alta (etiquetas vX.Y.Z) que no supera current, y al menos la de git describe."""
    best = described
    for tag in tags:
        m = re.fullmatch(r"v(\d+(?:\.\d+)*)", tag.strip())
        if m and _key(best) < _key(m[1]) <= _key(current):
            best = m[1]
    return best


def describe(text: str, tags: list[str] | None = None, current: str = __version__) -> str | None:
    """Versión a partir de la salida de git describe --long (None si no la reconoce). tags: todas las etiquetas del
    repositorio (git tag), para nombrar la última versión publicada aunque no esté en el historial de la rama."""
    m = _DESCRIBE.fullmatch(text.strip())
    if m is None:
        return None
    if m["ahead"] == "0":
        return m["tag"]
    return f"{latest_release(m['tag'], tags or [], current)} + desarrollo ({m['commit']})"


@functools.cache
def version_label() -> str:
    """La versión que se muestra: la publicada (__version__) o, desde una copia del repositorio que no está en una
    versión publicada, «X.Y.Z + desarrollo (commit)». Sin git, en el ejecutable del instalador o si git no responde,
    __version__."""
    if getattr(sys, "frozen", False) or not (ROOT / ".git").exists():
        return __version__
    def git(*args: str) -> subprocess.CompletedProcess:
        return subprocess.run(["git", "-C", str(ROOT), *args], capture_output=True, text=True, timeout=10,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        out = git("describe", "--tags", "--long", "--abbrev=7", "--match", "v[0-9]*")
        tags = git("tag", "--list", "v[0-9]*")
    except (OSError, subprocess.SubprocessError):
        return __version__
    if out.returncode != 0:
        return __version__
    return describe(out.stdout, tags.stdout.split() if tags.returncode == 0 else []) or __version__
