"""Libros de referencia: cada versión del programa debe seguir actualizando, analizando y migrando los libros
generados con las versiones anteriores desde la 2.7.1 (uno por formato, en tests/libros/formato_N.xlsx).

Además, la variable de entorno EVM_LIBROS (rutas separadas por ';') indica libros reales para revisar sin MS Project:
se leen y se ejecuta 'analizar' en una copia.
"""
from __future__ import annotations

import os
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from evm.analysis import A_ACTION_ID, A_CAUSE, A_CODE, A_STATUS, SHEET_ANALYSIS, read_entries
from evm.cli import main
from evm.update import read_book
from evm.workbook import BOOK_FORMAT

LIBROS = Path(__file__).resolve().parent / "libros"
REFERENCES = sorted(LIBROS.glob("formato_*.xlsx"))
REAL = [Path(p) for p in os.environ.get("EVM_LIBROS", "").split(";") if p.strip()]


def test_there_is_a_reference_book_of_the_current_format():
    """Al aumentar formato_libro se debe generar el libro de referencia del formato nuevo (y migrar los anteriores):
    python tests\\libros\\generar_libro_referencia.py"""
    assert LIBROS / f"formato_{BOOK_FORMAT}.xlsx" in REFERENCES


@pytest.mark.parametrize("reference", REFERENCES, ids=[p.stem for p in REFERENCES])
def test_reference_book_updates_analyses_and_migrates(reference, tmp_path, example_files, generator):
    book = tmp_path / "libro.xlsx"
    shutil.copy(reference, book)
    before = read_book(book)
    assert before.status_date == date(2027, 6, 30)
    # Cierre de julio: AC cargado a mano, actualizar y analizar
    generator.fill_book(book, date(2027, 7, 31))
    assert main(["actualizar", "--origen", str(example_files["jul"]), "--config", str(example_files["config"]),
                 "--libro", str(book), "--sin-respaldo"]) == 0
    assert main(["analizar", "--libro", str(book), "--sin-respaldo"]) == 0
    after = read_book(book)
    assert after.status_date == date(2027, 7, 31)
    june = after.horizon.index(date(2027, 6, 1))
    for s in before.input_rows:                    # el historial de junio no cambia
        assert after.pct_ev[""][s.row][: june + 1] == pytest.approx(before.pct_ev[""][s.row][: june + 1])
    # El texto del análisis de junio se conserva
    texts = {e.code: e for e in read_entries(load_workbook(book)[SHEET_ANALYSIS]) if e.month == date(2027, 6, 1)}
    from evm.analysis import TEXT_COLS
    assert texts["2.2.1"].texts[TEXT_COLS.index(A_CAUSE)] == "Planos del proveedor atrasados"
    assert texts["2.2.1"].texts[TEXT_COLS.index(A_ACTION_ID)] == "ACC-001"
    assert texts["2.2.1"].texts[TEXT_COLS.index(A_STATUS)] == "En curso"
    # Libro nuevo con --migrar-desde: conserva el BAC y el historial
    new = tmp_path / "nuevo.xlsx"
    assert main(["generar", "--origen", str(example_files["jul"]), "--config", str(example_files["config"]),
                 "--salida", str(new), "--migrar-desde", str(book)]) == 0
    migrated = read_book(new)
    assert sum(migrated.bac.values()) == pytest.approx(sum(after.bac.values()))
    ws = load_workbook(new)[SHEET_ANALYSIS]
    assert any(ws.cell(r, A_CODE).value == "2.2.1" for r in range(1, ws.max_row + 1))


@pytest.mark.skipif(not REAL, reason="sin EVM_LIBROS (libros reales, separados por ';')")
@pytest.mark.parametrize("path", REAL, ids=[p.name for p in REAL])
def test_real_book_is_read_and_analysed(path, tmp_path):
    """Libro real (EVM_LIBROS): se lee y 'analizar' funciona en una copia, sin MS Project."""
    assert path.is_file(), f"No existe {path}"
    book = tmp_path / path.name
    shutil.copy(path, book)
    data = read_book(book)
    assert data.input_rows
    if data.status_date is not None:
        assert main(["analizar", "--libro", str(book), "--sin-respaldo"]) == 0
