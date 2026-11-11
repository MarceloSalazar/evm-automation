"""Título y notas de una versión del CHANGELOG, para la publicación (release) en GitHub.

Las notas de la publicación son el resumen de la versión (sección `### Resumen` de su entrada en el
CHANGELOG), con los enlaces completos a la documentación de la etiqueta y, al final, los enlaces a las
notas detalladas del CHANGELOG y a los cambios en el código. El detalle (Nuevas funciones, Correcciones y
Otros) queda en el CHANGELOG.

Uso (desde la carpeta raíz del repositorio):
    python docs\\herramientas\\notas_version.py 1.5.0 --titulo     # "v1.5.0 · Análisis de desviaciones..."
    python docs\\herramientas\\notas_version.py 1.5.0 > notas.md    # resumen para la publicación
    python docs\\herramientas\\notas_version.py 1.5.0 --detalle     # notas detalladas de la versión

Publicación de una versión (después de fusionar en main el PR que la cambia):
    git tag -a v1.5.0 -m "Versión 1.5.0" <commit de main>
    git push origin v1.5.0
    gh release create v1.5.0 --title "<título>" --notes-file notas.md

Notas de una publicación ya creada: gh release edit v1.5.0 --notes-file notas.md
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

CHANGELOG = Path(__file__).resolve().parents[2] / "CHANGELOG.md"
REPO = "https://github.com/MarceloSalazar/evm-automation"
SECTIONS = ("Nuevas funciones", "Correcciones", "Otros")
# Enlace relativo de Markdown (docs/..., README.md...): ni dirección completa ni ancla de la misma página.
RELATIVE_LINK = re.compile(r"\]\((?!https?://|mailto:|#)([^)\s]+)\)")
SUMMARY_LABEL = re.compile(r"^\*\*(.+)\*\*$")
LIST_ITEM = re.compile(r"\s*(?:[-*+]|\d+\.)\s")


def absolute_links(text: str, version: str) -> str:
    """Enlaces relativos del CHANGELOG convertidos en enlaces completos a los archivos de la etiqueta."""
    return RELATIVE_LINK.sub(lambda m: f"]({REPO}/blob/v{version}/{m.group(1)})", text)


def unwrap(text: str) -> str:
    """Cada viñeta en una sola línea: la publicación de GitHub muestra cada salto de línea del texto, y las
    líneas del CHANGELOG se cortan a unos 110 caracteres (una línea cortada que empieza con `>` se vería,
    además, como una cita)."""
    out: list[str] = []
    fence = False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            fence = not fence
        elif (not fence and out and out[-1].strip() and line[:1] in (" ", "\t") and line.strip()
              and not LIST_ITEM.match(line)):
            out[-1] = out[-1].rstrip() + " " + line.strip()
            continue
        out.append(line)
    return "\n".join(out)


def version_entry(version: str, text: str) -> tuple[str, str]:
    """(título, texto) de la entrada de la versión: el título es el texto después de la fecha en el
    encabezado; el texto va hasta la entrada siguiente o los enlaces de comparación del final."""
    head = re.compile(rf"^## \[{re.escape(version)}\] - (\S+) · (.+)$", re.M)
    m = head.search(text)
    if not m:
        raise SystemExit(f"La versión {version} no está en {CHANGELOG.name} (encabezado '## [{version}] - fecha · "
                         "título').")
    if f'<a id="v{version}"></a>' not in text[:m.start()].rsplit("\n## ", 1)[-1]:
        raise SystemExit(f'A la versión {version} le falta el ancla <a id="v{version}"></a> antes de su encabezado '
                         "(el enlace a las notas detalladas de la publicación la usa).")
    rest = text[m.end():]
    end = re.search(r'^<a id="|^## \[|^\[[^\]]+\]: ', rest, re.M)
    return m.group(2).strip(), (rest[:end.start()] if end else rest).strip()


def split_entry(version: str, entry: str) -> tuple[str, str]:
    """(resumen, detalle) de la entrada: el resumen es la sección `### Resumen` y el detalle, el resto."""
    m = re.search(r"^### Resumen[ \t]*\n(.*?)(?=^### |\Z)", entry, re.M | re.S)
    if not m:
        raise SystemExit(f"A la versión {version} le falta la sección '### Resumen' (al principio de la entrada).")
    detail = (entry[:m.start()] + entry[m.end():]).strip()
    for section in SECTIONS:
        if f"### {section}" not in detail:
            raise SystemExit(f"A la versión {version} le falta la sección '### {section}'.")
    return m.group(1).strip(), detail


def summary_sections(version: str, summary: str) -> str:
    """Resumen con sus partes (**Nuevas funciones**, **Correcciones**, **Otros**) como encabezados."""
    lines, labels = [], []
    for line in summary.splitlines():
        label = SUMMARY_LABEL.match(line.strip())
        if label:
            if label.group(1) not in SECTIONS:
                raise SystemExit(f"Parte del resumen de la versión {version} no reconocida: '{line.strip()}' "
                                 f"(se admiten {', '.join(f'**{s}**' for s in SECTIONS)}).")
            labels.append(label.group(1))
            line = f"### {label.group(1)}"
        lines.append(line)
    if not labels or not any(line.startswith("- ") for line in lines):
        raise SystemExit(f"El resumen de la versión {version} debe tener al menos una parte "
                         f"({', '.join(f'**{s}**' for s in SECTIONS)}) con viñetas.")
    return "\n".join(lines)


def footer(version: str, text: str) -> str:
    """Enlaces a las notas detalladas del CHANGELOG y a los cambios en el código desde la versión anterior."""
    links = [f"[notas detalladas de la versión]({REPO}/blob/main/CHANGELOG.md#v{version})"]
    compare = re.search(rf"^\[{re.escape(version)}\]: \S+/compare/v(\S+)\.\.\.v{re.escape(version)}$", text, re.M)
    if compare:
        links.append(f"[cambios en el código desde la versión {compare.group(1)}]"
                     f"({REPO}/compare/v{compare.group(1)}...v{version})")
    else:
        links.append(f"[código de la versión]({REPO}/tree/v{version})")
    return "---\n\n**Más información:** " + " · ".join(links) + "\n"


def release_notes(version: str, text: str) -> str:
    """Notas de la publicación: el resumen de la versión con los enlaces completos y el pie con los enlaces."""
    _, entry = version_entry(version, text)
    summary, _ = split_entry(version, entry)
    return unwrap(absolute_links(summary_sections(version, summary), version)) + "\n\n" + footer(version, text)


def detailed_notes(version: str, text: str) -> str:
    """Notas detalladas de la versión (sin el resumen), con los enlaces completos y una línea por viñeta."""
    _, entry = version_entry(version, text)
    _, detail = split_entry(version, entry)
    return unwrap(absolute_links(detail, version)) + "\n"


def release_title(version: str, text: str) -> str:
    title, _ = version_entry(version, text)
    return f"v{version} · {title}"


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("version", help="versión, por ejemplo 1.5.0")
    group = p.add_mutually_exclusive_group()
    group.add_argument("--titulo", action="store_true", help="muestra solo el título de la publicación")
    group.add_argument("--detalle", action="store_true", help="muestra las notas detalladas, sin el resumen")
    a = p.parse_args()
    text = CHANGELOG.read_text(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    if a.titulo:
        print(release_title(a.version, text))
    elif a.detalle:
        print(detailed_notes(a.version, text), end="")
    else:
        print(release_notes(a.version, text), end="")


if __name__ == "__main__":
    main()
