"""Genera docs/codigos.md (códigos de errores y advertencias) a partir del catálogo del programa (evm/messages.py).

Uso (desde la carpeta raíz del repositorio):
    python docs\\herramientas\\codigos.py              # escribe docs/codigos.md
    python docs\\herramientas\\codigos.py --comprobar  # solo comprueba que esté al día (código de salida 1 si no)

Cada vez que se agrega, cambia o retira un código del catálogo se vuelve a generar el documento (las pruebas
comprueban que está al día). Metodología: docs/desarrollo.md, sección 5.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from evm.messages import AREAS, CATALOG, SEV_ERROR, SEV_FATAL, SEV_INFO, SEV_WARN  # noqa: E402

OUT = ROOT / "docs" / "codigos.md"

SEVERITY_LABEL = {SEV_FATAL: "Error (detiene el programa)", SEV_ERROR: "Error", SEV_WARN: "Advertencia",
                  SEV_INFO: "Info"}

SEVERITIES = [
    ("Severidad", "En la consola", "Qué indica"),
    ("Error (detiene el programa)", "`ERROR CFG-003: ...` (salida de error)",
     "El comando no puede seguir: termina sin crear ni modificar archivos, con código de salida 2 (3 si es un "
     "error inesperado, `SIS-001`)."),
    ("Error", "`[X] EDT-009 ...`", "Dato incorrecto que se debe corregir, o dato que no se pudo trasladar o calcular. "
                                   "`validar` termina con código de salida 1."),
    ("Advertencia", "`[!] FEC-007 ...`", "Se debe revisar: puede afectar el resultado."),
    ("Info", "`[i] EDT-016 ...`", "Información sobre lo que hizo el programa."),
]
FORMS = [
    ("Forma", "Qué silencia"),
    ("`LOG-004`", "Todas las observaciones de ese código."),
    ("`MET-004:1.2.1`", "Solo las de la fila o tarea con ese código EDT (los ceros finales no cuentan: `1.2.1` es "
                        "`1.2.1.0`). Las demás filas siguen mostrando el aviso."),
]

# Formato Markdown de los textos del catálogo: opciones, claves y archivos como código; códigos como enlaces
CODE_PATTERNS = [r"(?<![\w`-])--[a-z][a-z-]*", r"\[\[proyectos\]\]",
                 r"\[(?:campos|exportacion|horizonte|semaforo|validacion|multiempresa|metodos)\](?: [a-z_0-9]+)?",
                 r"\b(?:config\.ejemplo\.toml|evm/messages\.py|_meta|_estructura)\b"]

INTRO = """\
# Códigos de errores y advertencias

[Inicio](../README.md) › [Documentación](README.md) › Códigos de errores y advertencias

Cada mensaje que muestra el programa, en la consola o en la hoja `Validación` del libro, lleva un
**código** como `FEC-003`. Este documento explica cada código: qué lo causa y qué hacer. También explica
cómo **silenciar** las observaciones que ya se revisaron y no se quieren volver a ver.

> Este documento se genera a partir del catálogo del programa (`evm/messages.py`) con
> `python docs\\herramientas\\codigos.py`: no se edita a mano. Para agregar un código, ver
> [Desarrollo](desarrollo.md#5-códigos-de-errores-y-advertencias).

---

## Contenido

1. [Cómo se lee un código](#1-cómo-se-lee-un-código)
2. [Silenciar observaciones](#2-silenciar-observaciones)
3. [Áreas](#3-áreas)
4. [Códigos](#4-códigos)
{areas_toc}

---

## 1. Cómo se lee un código

El código tiene la forma **ÁREA-NNN**: tres letras del área (por ejemplo `FEC`, *Fechas y duración*) y un
número. El código de un mensaje no cambia aunque cambien su texto o su severidad, y un código que deja de
usarse no se reutiliza: queda en este documento como *retirado*.

{severities}

Algunos códigos tienen una severidad u otra según el caso (por ejemplo, *Advertencia o Info*); el
documento lo indica en cada uno.

En la **consola**, cada observación muestra su código, el código EDT y el nombre de la fila o tarea (si
corresponde) y el mensaje:

```text
  [!] FEC-007 1.2.3 Montaje: Duración estimada («?»).
  [i] EDT-016 Se excluyeron 2 tareas inactivas: no forman parte del % planificado.
```

En la **hoja `Validación`** del libro, el código está en la columna `Código`, junto a la severidad; con el
filtro de la columna se pueden ver las observaciones de un código.

---

## 2. Silenciar observaciones

Las observaciones de un código se pueden **silenciar**: no se muestran en la consola ni en la hoja
`Validación`. Sirve para no volver a ver lo que ya se revisó y se acepta (por ejemplo, un control DCMA que
no aplica al proyecto). Se indica en el archivo de configuración, en la sección `[validacion]`:

```toml
[validacion]
silenciar = ["LOG-004", "MET-004:1.2.1", "FEC-007"]
```

o con la opción `--silenciar` de cualquier comando (se puede repetir, y los códigos se separan con comas);
se suma a los de la configuración:

```powershell
python evm_export.py validar --origen "C:\\Programa\\Programa.mpp" --silenciar LOG-004,FEC-007
```

{forms}

- Siempre se informa cuántas observaciones se silenciaron y de qué códigos (`CFG-049`), para que no
  desaparezcan sin dejar rastro.
- Los errores silenciados no cuentan para el código de salida de `validar`.
- Los errores que detienen el programa no se pueden silenciar (`CFG-048`).
- Un código que no existe en el catálogo se informa con `CFG-047` y se ignora.
- Un código retirado se acepta y no silencia nada: las configuraciones que lo usan siguen siendo válidas.

---

## 3. Áreas

| Área | Nombre | Tema en la hoja Validación | Dónde aparecen | Códigos |
|---|---|---|---|---|
{areas_table}

---

## 4. Códigos"""


def _anchor(title: str) -> str:
    """Ancla de un encabezado en GitHub: minúsculas, sin signos de puntuación y con guiones."""
    text = "".join(ch for ch in title.lower() if ch.isalnum() or ch in " -_")
    return text.replace(" ", "-")


def _table(rows: list[tuple]) -> str:
    head, *body = rows
    lines = ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)]
    return "\n".join(lines + ["| " + " | ".join(r) + " |" for r in body])


def _severity(code) -> str:
    if code.severities == (SEV_FATAL, SEV_ERROR):
        return "Error (detiene el programa, salvo en validar)"
    return " o ".join(SEVERITY_LABEL[s] for s in code.severities)


def _md(text: str) -> str:
    """Texto del catálogo en Markdown: opciones y claves como código, y códigos del catálogo como enlaces."""
    for pattern in CODE_PATTERNS:
        text = re.sub(pattern, lambda m: f"`{m.group(0)}`", text)
    return re.sub(r"\b[A-Z]{3}-\d{3}\b",
                  lambda m: f"[{m.group(0)}](#{m.group(0).lower()})" if m.group(0) in CATALOG else m.group(0), text)


def render() -> str:
    """Contenido de docs/codigos.md."""
    by_area: dict[str, list] = {p: [] for p in AREAS}
    for code in CATALOG.values():
        by_area[code.id[:3]].append(code)
    titles = {p: f"4.{i} {p} · {a.title}" for i, (p, a) in enumerate(AREAS.items(), start=1)}
    toc = "\n".join(f"   {i}. [{a.title}](#{_anchor(titles[p])})"
                    for i, (p, a) in enumerate(AREAS.items(), start=1))
    table = "\n".join(f"| [{p}](#{_anchor(titles[p])}) | {a.title} | {a.topic} | {_md(a.where)} | "
                      f"{len(by_area[p])} |" for p, a in AREAS.items())
    out = [INTRO.format(areas_toc=toc, areas_table=table, severities=_table(SEVERITIES), forms=_table(FORMS))]
    for p, area in AREAS.items():
        out.append(f"\n### {titles[p]}\n\n{_md(area.summary)}\n")
        rows = [("Código", "Severidad", "Descripción")]
        rows += [(f"[{c.id}](#{c.id.lower()})", _severity(c), _md(c.title) + (" *(retirado)*" if c.retired else ""))
                 for c in by_area[p]]
        out.append(_table(rows))
        for code in by_area[p]:
            out.append(f"\n#### {code.id}\n")
            facts = [f"**{_md(code.title)}**", _severity(code), _md(code.location),
                     "se puede silenciar" if code.silenceable else "no se puede silenciar"]
            if code.retired:
                facts.append(f"retirado en la versión {code.retired}")
            out.append(" · ".join(facts) + "\n")
            out.append(f"{_md(code.cause)}\n")
            out.append(f"**Qué hacer:** {_md(code.action)}")
        out.append("\n---")
    return "\n".join(out[:-1]) + "\n"


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--comprobar", action="store_true", help="solo comprueba que docs/codigos.md esté al día")
    args = p.parse_args()
    text = render()
    if args.comprobar:
        current = OUT.read_text(encoding="utf-8") if OUT.exists() else ""
        if current != text:
            raise SystemExit(f"{OUT.relative_to(ROOT)} no está al día: se debe ejecutar "
                             "python docs\\herramientas\\codigos.py")
        print(f"{OUT.relative_to(ROOT)} está al día ({len(CATALOG)} códigos).")
        return
    OUT.write_text(text, encoding="utf-8", newline="\n")
    print(f"{OUT.relative_to(ROOT)}: {len(CATALOG)} códigos.")


if __name__ == "__main__":
    main()
