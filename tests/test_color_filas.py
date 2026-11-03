"""Color de fondo de las filas del libro: depende del nivel del código EDT, no del tipo de fila (resumen o de ingreso)
ni de lo que cuelga por debajo. Nivel 1, azul oscuro; nivel 2, azul claro; los niveles siguientes y las filas -D, sin
fondo. La negrita, solo en las filas resumen."""
import pytest
from openpyxl import load_workbook

from evm import dashboard
from evm.config import Config
from evm.wbs import KIND_RESIDUAL, code_key
from evm.workbook import (COL_BAC, FILL_INPUT, FILL_L1, FILL_L2, FIRST_MONTH_COL, SHEET_AC, SHEET_EV, SHEET_IND,
                          SHEET_PV, pct_sheet, write_workbook)

from builders import report, rows_by_code, tree

# 2.7.1 tiene actividades con fila propia (fila resumen); 2.8.1 y 2.8.2 no (filas de ingreso). 2.9 solo tiene tareas
# sin código EDT (fila de ingreso) y 2.8 tiene una tarea sin código (fila 2.8-D)
TASKS = '''
    2 Fase 2
      2.7 Entregable siete
        2.7.1 Paquete con actividades
          2.7.1.1 Actividad a
          2.7.1.2 Actividad b
        2.7.2 Paquete simple
      2.8 Entregable ocho
        2.8.1 Paquete uno
        2.8.2 Paquete dos
        - Tarea suelta
      2.9 Entregable nueve
        - Tarea sin código
'''
L1, L2 = FILL_L1.fgColor.rgb[-6:], FILL_L2.fgColor.rgb[-6:]
EXPECTED = {"2": L1, "2.7": L2, "2.8": L2, "2.9": L2, "2.7.1": None, "2.7.2": None, "2.7.1.1": None,
            "2.8.1": None, "2.8.2": None, "2.8-D": None}
SUMMARIES = {"2", "2.7", "2.8", "2.7.1"}


def fill_of(cell) -> str | None:
    return cell.fill.fgColor.rgb[-6:] if cell.fill.fill_type == "solid" else None


def rows_of(ws) -> dict[str, int]:
    """Fila de cada código EDT de la hoja (columna A), comparado sin los ceros finales."""
    return {code_key(str(ws.cell(r, 1).value)): r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}


@pytest.fixture(scope="module")
def book(tmp_path_factory):
    path = tmp_path_factory.mktemp("color") / "libro.xlsx"
    cfg = Config(max_level=4)
    write_workbook(report(tree(TASKS), cfg), cfg, path)
    return load_workbook(path)


def test_wbs_level_is_the_level_of_the_code():
    rows = {code_key(k): r for k, r in rows_by_code(report(tree(TASKS), Config(max_level=4))).items()}
    assert {k: r.wbs_level for k, r in rows.items()} == {
        "2": 1, "2.7": 2, "2.7.1": 3, "2.7.1.1": 4, "2.7.1.2": 4, "2.7.2": 3, "2.8": 2, "2.8.1": 3, "2.8.2": 3,
        "2.8-D": 0, "2.9": 2}
    assert rows["2.8-D"].kind == KIND_RESIDUAL


@pytest.mark.parametrize("sheet", [SHEET_PV, SHEET_EV, SHEET_AC, pct_sheet(SHEET_PV), pct_sheet(SHEET_EV), SHEET_IND])
def test_fill_and_bold_by_wbs_level(book, sheet):
    ws = book[sheet]
    rows = rows_of(book[SHEET_PV])      # las hojas usan las filas de la hoja PV (Indicadores EDT, con fórmulas)
    assert {k: fill_of(ws.cell(rows[k], 1)) for k in EXPECTED} == EXPECTED
    assert {k for k in EXPECTED if ws.cell(rows[k], 2).font.b} == SUMMARIES
    if sheet != SHEET_IND:              # el horizonte mensual lleva el mismo fondo (salvo las celdas de ingreso)
        assert ws.cell(rows["2.8-D"], 2).font.i
        assert fill_of(ws.cell(rows["2.8"], FIRST_MONTH_COL)) == L2
    if sheet in (pct_sheet(SHEET_PV), pct_sheet(SHEET_EV)):  # fila de ingreso de nivel 2: también los meses vacíos
        assert all(fill_of(ws.cell(rows["2.9"], c)) == L2 for c in range(FIRST_MONTH_COL, FIRST_MONTH_COL + 12))


def test_input_cells_stay_yellow_on_colored_rows(book):
    yellow = FILL_INPUT.fgColor.rgb[-6:]
    pv, ac = book[SHEET_PV], book[SHEET_AC]
    rows = rows_of(pv)
    # 2.9: fila de ingreso de nivel 2 (azul claro), con el BAC y el AC del mes en amarillo
    assert fill_of(pv.cell(rows["2.9"], COL_BAC)) == yellow and fill_of(pv.cell(rows["2.9"], 1)) == L2
    assert fill_of(ac.cell(rows["2.9"], FIRST_MONTH_COL)) == yellow
    assert fill_of(pv.cell(rows["2.8"], COL_BAC)) == L2         # subtotal: fórmula, no se ingresa


def test_fill_uses_code_level_when_the_parent_has_no_row(tmp_path):
    """Con la lista EDT sin 2 ni 2.7, 2.8 queda en el primer nivel del libro pero conserva el color del nivel 2."""
    cfg = Config(wbs_list=["2.7.1", "2.8", "2.8.1", "2.8.2", "2.9"], wbs_list_file="edt.txt")
    path = tmp_path / "libro.xlsx"
    rep = report(tree(TASKS), cfg)
    assert {code_key(r.code): r.level for r in rep.rows if r.kind != KIND_RESIDUAL}["2.8"] == 1
    write_workbook(rep, cfg, path)
    wb = load_workbook(path)
    ws = wb[SHEET_PV]
    rows = rows_of(ws)
    assert {k: fill_of(ws.cell(rows[k], 1)) for k in ("2.7.1", "2.8", "2.8.1", "2.9")} == {
        "2.7.1": None, "2.8": L2, "2.8.1": None, "2.9": L2}
    # Cuadro de Mando: ningún elemento es de nivel 1, aunque 2.7.1, 2.8 y 2.9 estén en el primer nivel del libro
    assert [fill_of(c) for c, _ in dashboard_table(wb)] == [None] * 5


def dashboard_table(wb) -> list:
    """Celda del código y de la actividad de cada fila de la tabla de elementos EDT del Cuadro de Mando."""
    ws = wb[dashboard.SHEET_DASHBOARD]
    first = dashboard.chart_rows()[2] + 1
    out = []
    while ws.cell(first + len(out), 12).value:
        out.append((ws.cell(first + len(out), 12), ws.cell(first + len(out), 14)))
    return out


def test_dashboard_table_highlights_level_1_and_bolds_summaries(book):
    """Tabla del Cuadro de Mando, en el orden del libro (2, 2.7, 2.8, 2.9): fondo solo en el nivel 1 del código EDT y
    negrita en las filas resumen (2.9 es una fila de ingreso)."""
    table = dashboard_table(book)
    level1 = dashboard.FILL_LEVEL1.fgColor.rgb[-6:]
    assert [fill_of(c) for c, _ in table] == [level1, None, None, None]
    assert [n.font.b for _, n in table] == [True, True, True, False]


def test_children_of_a_planificacion_row_are_one_level_below():
    """Sin elegir planificación, INTERNO es una fila de nivel 1 y sus hijos, con códigos desde 1, son de nivel 2."""
    rep = report(tree('''
        INTERNO Planificación interna
          1 Fase
            1.1 Entregable
              1.1.1 Paquete
    '''), Config(max_level=4))
    assert [(r.code, r.wbs_level) for r in rep.rows] == [
        ("INTERNO", 1), ("1.0.0.0", 2), ("1.1.0.0", 3), ("1.1.1.0", 4)]


@pytest.mark.parametrize("ignore, level", [(False, 3), (True, 2)])
def test_interior_zero_follows_ignorar_ceros_intermedios(ignore, level):
    rep = report(tree('''
        1 Fase
          1.0.1 Actividad de la fase
          1.1 Entregable
    '''), Config(ignore_interior_zeros=ignore))
    assert {code_key(r.code): r.wbs_level for r in rep.rows}["1.0.1"] == level
