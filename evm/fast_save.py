"""Guardado rápido del libro ([exportacion] guardado_rapido, activado por defecto).

Al guardar, openpyxl arma cada celda como un elemento XML y lo escribe con un serializador genérico: en los libros
grandes (cientos de miles de celdas) es la mayor parte del tiempo de 'generar', 'actualizar' y 'analizar'. Mientras se
guarda el libro, este módulo reemplaza la función de openpyxl que escribe cada celda por otra que arma el texto XML de
la celda directamente. El archivo que resulta es el mismo, byte a byte (lo comprueban las pruebas de
tests/test_rendimiento.py). Las celdas poco frecuentes (fechas, fórmulas matriciales, hipervínculos, texto
enriquecido, errores) se escriben con la función de openpyxl.

Se aplica solo con la versión de openpyxl con que se probó (3.1.x) y con su escritor estándar (sin lxml); si no, el
libro se guarda en el modo normal, sin aviso. Con guardado_rapido = false se usa siempre el modo normal.
"""
from __future__ import annotations

import logging
from contextlib import contextmanager

log = logging.getLogger(__name__)

TESTED_VERSIONS = ("3.1.",)         # versiones de openpyxl con las que el resultado es idéntico al modo normal

try:                                # partes internas de openpyxl: si no existen, se guarda en el modo normal
    import openpyxl
    from openpyxl.cell._writer import etree_write_cell
    from openpyxl.compat import safe_string
    from openpyxl.worksheet import _writer as sheet_writer
except ImportError:                 # pragma: no cover
    sheet_writer = None


def available() -> bool:
    """El guardado rápido se puede usar: openpyxl de una versión probada, con su escritor estándar de celdas."""
    return (sheet_writer is not None and openpyxl.__version__.startswith(TESTED_VERSIONS)
            and getattr(sheet_writer, "write_cell", None) is etree_write_cell)


@contextmanager
def fast_cells(active: bool = True):
    """Mientras dura el bloque, openpyxl escribe las celdas con _write_cell (si active y si se puede usar)."""
    if not active or not available():
        if active:
            log.debug("Guardado rápido no disponible (openpyxl %s): se guarda en el modo normal",
                      getattr(openpyxl, "__version__", "?"))
        yield
        return
    original = sheet_writer.write_cell
    sheet_writer.write_cell = _write_cell
    try:
        yield
    finally:
        sheet_writer.write_cell = original


def _escape(text: str) -> str:
    """Como xml.etree.ElementTree._escape_cdata, que usa openpyxl para el texto de los elementos."""
    if "&" in text:
        text = text.replace("&", "&amp;")
    if "<" in text:
        text = text.replace("<", "&lt;")
    if ">" in text:
        text = text.replace(">", "&gt;")
    return text


def _write_cell(xf, worksheet, cell, styled=None) -> None:
    """Escribe la celda <c> como lo hace openpyxl.cell._writer.etree_write_cell, armando el texto directamente."""
    write = getattr(xf, "_file", None)      # función de escritura del escritor estándar (et_xmlfile)
    value = cell._value
    kind = cell.data_type
    simple = (kind in ("n", "b") or (kind == "s" and isinstance(value, str))
              or (kind == "f" and isinstance(value, str)))
    if write is None or not simple or cell.hyperlink:
        etree_write_cell(xf, worksheet, cell, styled)
        return
    head = f'<c r="{cell.coordinate}"'
    if styled:
        head += f' s="{cell.style_id}"'
    if kind == "s":
        head += ' t="inlineStr"'
    elif kind != "f":
        head += f' t="{kind}"'
    if value is None or value == "":
        write(head + " />")
    elif kind == "f":
        formula = value[1:]
        write(f"{head}><f>{_escape(formula)}</f><v /></c>" if formula else f"{head}><f /><v /></c>")
    elif kind == "s":
        stripped = value.strip()
        space = ' xml:space="preserve"' if stripped and value != stripped else ""
        write(f"{head}><is><t{space}>{_escape(value)}</t></is></c>")
    else:
        text = safe_string(value)
        write(f"{head}><v>{_escape(text)}</v></c>" if text else f"{head}><v /></c>")
