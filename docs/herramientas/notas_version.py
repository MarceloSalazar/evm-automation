"""Título y notas de una versión del CHANGELOG, para la publicación (release) en GitHub.

Uso (desde la carpeta raíz del repositorio):
    python docs\\herramientas\\notas_version.py 1.5.0 --titulo     # "v1.5.0 · Análisis de desviaciones..."
    python docs\\herramientas\\notas_version.py 1.5.0 > notas.md    # secciones de la versión

Publicación de una versión (después de fusionar en main el PR que la cambia):
    git tag -a v1.5.0 -m "Versión 1.5.0" <commit de main>
    git push origin v1.5.0
    gh release create v1.5.0 --title "<título>" --notes-file notas.md
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

CHANGELOG = Path(__file__).resolve().parents[2] / "CHANGELOG.md"


def version_section(version: str, text: str) -> tuple[str, str]:
    """(título, notas) de la versión: el título es el texto después de la fecha en el encabezado."""
    head = re.compile(rf"^## \[{re.escape(version)}\] - (\S+) · (.+)$", re.M)
    m = head.search(text)
    if not m:
        raise SystemExit(f"La versión {version} no está en {CHANGELOG.name} (encabezado '## [{version}] - fecha · "
                         "título').")
    rest = text[m.end():]
    end = re.search(r"^## \[|^\[\d+\.\d+\.\d+\]: ", rest, re.M)
    notes = rest[:end.start()] if end else rest
    for section in ("Nuevas funciones", "Correcciones", "Otros"):
        if f"### {section}" not in notes:
            raise SystemExit(f"A la versión {version} le falta la sección '### {section}'.")
    return f"v{version} · {m.group(2).strip()}", notes.strip() + "\n"


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("version", help="versión, por ejemplo 1.5.0")
    p.add_argument("--titulo", action="store_true", help="muestra solo el título de la publicación")
    a = p.parse_args()
    title, notes = version_section(a.version, CHANGELOG.read_text(encoding="utf-8"))
    sys.stdout.reconfigure(encoding="utf-8")
    print(title if a.titulo else notes, end="\n" if a.titulo else "")


if __name__ == "__main__":
    main()
