"""Genera el libro de referencia del formato actual del libro (formato_libro de la hoja _meta).

El libro de cada formato se genera una sola vez, con la versión que introduce el formato, y se conserva en esta
carpeta (formato_N.xlsx): la prueba tests/test_libros_referencia.py comprueba que cada versión del programa sigue
actualizando, analizando y migrando los libros de todos los formatos. Por eso no se sobrescribe un libro existente.

El libro es el del proyecto de ejemplo después del cierre de junio de 2027: BAC y AC de junio cargados como a mano,
'actualizar' de junio y texto en el Análisis de desviaciones.

Uso (desde la carpeta raíz del repositorio):
    python tests\\libros\\generar_libro_referencia.py
"""
from __future__ import annotations

import importlib.util
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from openpyxl import load_workbook  # noqa: E402

from evm.cli import main  # noqa: E402
from evm.workbook import BOOK_FORMAT  # noqa: E402

FOLDER = Path(__file__).resolve().parent
# Configuración del proyecto de ejemplo, la misma que usan las pruebas (tests/conftest.py)
CONFIG = ('[campos]\nmetodo = "Método EVM"\npeso_hito = "Peso hito"\n\n'
          '[exportacion]\nplanificacion = "INTERNO"\n')
JUN, JUL = date(2027, 6, 30), date(2027, 7, 31)


def generator():
    spec = importlib.util.spec_from_file_location("generar_ejemplo", ROOT / "ejemplos" / "generar_ejemplo.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["generar_ejemplo"] = mod
    spec.loader.exec_module(mod)
    return mod


def example_files(folder: Path) -> dict[str, Path]:
    """XML del ejemplo (línea base, junio y julio de 2027) y su configuración."""
    gen = generator()
    files = {"base": folder / "base.xml", "jun": folder / "jun.xml", "jul": folder / "jul.xml",
             "config": folder / "config.toml"}
    for key, status in (("base", None), ("jun", JUN), ("jul", JUL)):
        gen.write(gen.build_plans(), files[key], status)
    files["config"].write_text(CONFIG, encoding="utf-8")
    return files


def build(out: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        files = example_files(Path(tmp))
        gen = generator()
        book = Path(tmp) / "libro.xlsx"
        assert main(["generar", "--origen", str(files["base"]), "--config", str(files["config"]),
                     "--salida", str(book)]) == 0
        gen.fill_book(book)
        gen.fill_book(book, JUN)
        assert main(["actualizar", "--origen", str(files["jun"]), "--config", str(files["config"]),
                     "--libro", str(book), "--sin-respaldo"]) == 0
        # Texto del análisis de junio, como se completa a mano
        wb = load_workbook(book)
        from evm.analysis import (A_ACTION, A_ACTION_ID, A_CAUSE, A_CLOSE, A_CODE, A_DATE, A_OWNER, A_STATUS, FIRST,
                                  SHEET_ANALYSIS)
        ws = wb[SHEET_ANALYSIS]
        row = next(r for r in range(FIRST, ws.max_row + 1) if ws.cell(r, A_CODE).value == "2.2.1")
        for column, value in ((A_CAUSE, "Planos del proveedor atrasados"), (A_ACTION, "Reforzar el equipo"),
                              (A_ACTION_ID, "ACC-001"), (A_OWNER, "Ingeniería"),
                              (A_DATE, datetime(2027, 7, 15)), (A_CLOSE, "SPI del paquete en verde"),
                              (A_STATUS, "En curso")):
            ws.cell(row, column, value)
        wb.save(book)
        out.write_bytes(book.read_bytes())


if __name__ == "__main__":
    target = FOLDER / f"formato_{BOOK_FORMAT}.xlsx"
    if target.exists():
        sys.exit(f"Ya existe {target.name}: el libro de referencia de un formato no se vuelve a generar.")
    build(target)
    print(f"Libro de referencia: {target}")
