"""Depuración de costes: hoja PV-depuración, filas «(coste)» y comando procesar-costes."""
import zipfile
from datetime import date, datetime

import pytest
from openpyxl import load_workbook

from evm import cost_review as cr
from evm.cli import main
from evm.config import Config, ConfigError, load_config
from evm.cost_plan import run_cost_plan
from evm.messages import InputError
from evm.migrate import plan_migration, prior_from_book
from evm.update import indicators_for, read_book, run_analysis, run_update
from evm.wbs import SEV_INFO, SEV_WARN, build_report
from evm.workbook import (FILL_COST, FILL_INPUT, KIND_COST, SHEET_STRUCT, read_structure,
                          write_workbook)

from builders import P, T, issues_with, rows_by_code


JAN = (datetime(2027, 1, 4, 8), datetime(2027, 1, 29, 17))
FEB = (datetime(2027, 2, 1, 8), datetime(2027, 2, 26, 17))
MAR = (datetime(2027, 3, 1, 8), datetime(2027, 3, 26, 17))
START = date(2027, 1, 1)


def tasks() -> list:
    """EDT con un resumen de nivel 2 (1.2), una tarea suelta bajo 1 (fila «Otras tareas» 1.0.0-D) y tareas de las
    dos empresas."""
    return [
        T(1, 1, "1", "Gestión", bs=JAN[0], bf=MAR[1]),
        T(2, 2, "1.1", "Dirección", bs=JAN[0], bf=FEB[1], company="Ambas", method="% Completado"),
        T(3, 2, "1.2", "Contratos", bs=JAN[0], bf=MAR[1]),
        T(4, 3, "1.2.1", "Licitación", bs=JAN[0], bf=JAN[1], company="EmpresaA", method="0/100"),
        T(5, 3, "1.2.2", "Obras", bs=FEB[0], bf=MAR[1], company="Ambas", method="% Completado"),
        T(6, 2, "", "Asesoría", bs=MAR[0], bf=MAR[1], company="Ambas"),
        T(7, 1, "2", "Ingeniería", bs=JAN[0], bf=MAR[1]),
        T(8, 2, "2.1", "Básica", bs=JAN[0], bf=MAR[1], company="Ambas", method="% Completado"),
    ]


def config(review: bool = True, multi: bool = True, **kw) -> Config:
    kw.setdefault("start_month", START)
    cfg = Config(months=24, multi_enabled=multi, cost_review=review, **kw)
    cfg.validate()
    return cfg


def make_book(path, cfg, prior=None, task_list=None):
    rep = build_report(P(task_list or tasks()), cfg)
    if prior is not None:
        rep.issues.extend(plan_migration(rep, prior, cfg.max_level))
    write_workbook(rep, cfg, path, prior)
    return rep


def rows(wb) -> dict[str, int]:
    return {s.code: s.row for s in read_structure(wb)}


def set_bac(path, sheet: str, values: dict[str, float]) -> None:
    wb = load_workbook(path)
    r = rows(wb)
    for code, v in values.items():
        wb[sheet].cell(row=r[code], column=4, value=v)
    wb.save(path)


def paste(path, lines: list[list]) -> None:
    """Pega líneas en la hoja PV-depuración: [EDT cliente, actividad, EDT interna, actividad interna o componente,
    BAC, EAC, mes 1, mes 2...]."""
    wb = load_workbook(path)
    ws = wb[cr.SHEET_PASTE]
    for k, line in enumerate(lines):
        for c, v in enumerate(line, start=1):
            if v is not None:
                ws.cell(row=cr.P_FIRST + k, column=c, value=v)
    wb.save(path)




def python_status(lines: list[list], wb) -> list:
    """Columna 'Asignado a' según la lectura en Python de procesar-costes (cost_review.read_lines)."""
    structure = read_structure(wb)
    code = {s.row: s.code for s in structure}
    out = []
    for x in cr.read_lines(lines, cr.structure_table(structure)):
        if x.state == cr.STATE_OTHER:
            out.append(None)
        else:
            out.append(code[x.match] if x.state == cr.STATE_READ else cr.STATUS_OF[x.state])
    return out


# Plan de costes en tres niveles, como el de la empresa: EDT cliente, EDT interna (GPR00) y componentes de coste
# (solo la columna D). PLAN_DATA tiene solo los datos (celdas amarillas); calc_plan hace los cálculos (naranjas).
PLAN_DATA = [   # EDT cliente, actividad EDT, EDT interna, actividad interna o componente, BAC, EAC, meses
    ["1.0.0", "Gestión"],                                                        # resumen del libro
    ["1.1.0", "Dirección"],                                                      # -> 1.1.0
    [None, None, "GPR00", "Lanzamiento"],
    [None, None, None, "Costes indirectos X", 30, None, 10, 9, 8, 7, 6],
    [None, None, None, "Costes de gestión Y", 60, None, 11, 11, 22, 22, 11],
    ["1.2.0", "Contratos"],                                                      # resumen del libro
    ["1.2.1", "Licitación"],                                                     # -> 1.2.1
    [None, None, "GPR01", "Pliegos"],
    [None, None, None, "Personal", 30, None, 10, 10, 10],
    [None, None, None, "Viajes", 10, None, None, 10],
    [None, None, "GPR02", "EDT interna sin componentes", 20, None, None, None, 20],
    ["1.2.1.5", "Licitación, detalle", None, None, None, None, 5],               # incluida en 1.2.1
    ["1.2.2", "Obras", None, None, None, None, 56, 40],                          # -> 1.2.2, sin EDT interna
    ["1.2.7", "Código de un resumen", None, None, None, None, 5],                # sin fila del libro
    [None, "Nota sin código"],
    ["2.0.0", "Ingeniería", None, None, None, None, None, 25, 25],               # resumen sin detalle
    ["9.9", "Código inexistente", None, None, None, None, 7],                    # sin fila del libro
]
PLAN_STATUS = [cr.STATUS_SUMMARY, "1.1.0", None, None, None, cr.STATUS_SUMMARY, "1.2.1", None, None, None, None,
               cr.STATUS_INCLUDED, "1.2.2", cr.STATUS_NONE, None, cr.STATUS_SUMMARY_EMPTY, cr.STATUS_NONE]
PLAN_COST = {"1.1.0": [21, 20, 30, 29, 17], "1.2.1": [15, 20, 30], "1.2.2": [56, 40]}
PLAN_EAC = [283, 117, 117, 40, 77, 166, 65, 40, 30, 10, 20, 5, 96, 5, None, 50, 7]
PLAN = cr.calc_plan(PLAN_DATA)[0]           # con los cálculos hechos
BAC_A = {"1.1.0": 117.0, "1.2.1": 65.0, "1.2.2": 100.0, "2.1.0": 50.0, "1.0.0-D": 10.0}


def months(*pairs) -> list:
    """Importes por mes: months((1, 100), (3, 50)) -> [100, None, 50]."""
    out = [None] * max(m for m, _v in pairs)
    for m, v in pairs:
        out[m - 1] = v
    return out


# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
def test_config_keys_and_defaults(tmp_path):
    assert Config().cost_review is False and Config().cost_review_threshold == 10
    path = tmp_path / "c.toml"
    path.write_text('[exportacion]\ndepuracion_costes = true\ndepuracion_empresa = "empresab"\n'
                    'depuracion_alineamiento = true\ndepuracion_umbral = 15\ndepuracion_calculos = true\n\n'
                    '[multiempresa]\nactivar = true\n',
                    encoding="utf-8")
    cfg = load_config(path)
    assert cfg.cost_review and cfg.cost_review_company == "EmpresaB" and cfg.cost_review_role == "B"
    assert cfg.cost_review_alignment and cfg.cost_review_threshold == 15 and cfg.cost_review_calcs
    assert Config().cost_review_calcs is False
    assert config().cost_review_role == "A"                      # por defecto, la empresa A
    assert config(multi=False).cost_review_role == ""            # una sola empresa: la hoja PV
    assert config(review=False).cost_review_role is None
    with pytest.raises(ConfigError, match="depuracion_empresa"):
        config(cost_review_company="Otra")
    with pytest.raises(ConfigError, match="depuracion_umbral"):
        config(cost_review_threshold=0)


def test_codes_without_trailing_zeros():
    assert cr.norm_code("1.2.0") == "1.2" and cr.norm_code("1.0.0") == "1" and cr.norm_code(" 10.0 ") == "10"
    assert cr.norm_code("1.10") == "1.10" and cr.norm_code("1,2") == "1.2" and cr.norm_code(3) == "3"
    for text in ("XRF01", "", None, "1..2", ".1", "1.", "1.2-D", "1 2", "1e5"):
        assert cr.norm_code(text) is None, text


# ----------------------------------------------------------------------------
# Libro
# ----------------------------------------------------------------------------
def test_cost_rows_below_every_row_and_hidden_rows_elsewhere(tmp_path):
    path = tmp_path / "libro.xlsx"
    rep = make_book(path, config())
    wb = load_workbook(path)
    names = wb.sheetnames
    assert names.index(cr.SHEET_CHART) == names.index("% PV-EmpresaB") + 1
    assert names.index(cr.SHEET_PASTE) == names.index(cr.SHEET_CHART) + 1
    assert wb[cr.SHEET_HELPER].sheet_state == "hidden"
    r = rows(wb)
    assert sorted(r.values()) == list(range(5, 5 + 2 * len(rep.rows), 2))      # una fila «(coste)» por medio
    pv_a = wb["PV-EmpresaA"]
    total = max(r.values()) + 2
    for code, row in list(r.items()) + [("TOTAL", total)]:
        cost = pv_a.cell(row=row + 1, column=2)
        assert pv_a.cell(row=row + 1, column=1).value == code and cost.value.endswith(" (coste)"), code
        assert cost.fill.fgColor.rgb.endswith(FILL_COST.fgColor.rgb[-6:])
    # Filas de ingreso: los meses y el EAC de su EDT cliente; resúmenes: la suma de sus filas
    assert pv_a.cell(row=r["1.2.1"] + 1, column=7).value == cr.paste_sum(0, r["1.2.1"])
    assert pv_a.cell(row=r["1.2.1"] + 1, column=4).value == cr.paste_eac(r["1.2.1"])
    assert pv_a.cell(row=r["1.2.0"] + 1, column=4).value == f"=SUM(D{r['1.2.1'] + 1},D{r['1.2.2'] + 1})"
    assert pv_a.cell(row=r["1.2.0"] + 1, column=7).value == f"=SUM(G{r['1.2.1'] + 1},G{r['1.2.2'] + 1})"
    assert pv_a.row_dimensions[r["1.2.1"] + 1].outline_level == 3            # un nivel más que la EDT
    for sheet in ("PV", "PV-EmpresaB", "% PV-EmpresaA", "EV-EmpresaA", "AC", "Indicadores EDT"):
        ws = wb[sheet]
        assert ws.row_dimensions[r["1.2.1"] + 1].hidden, sheet
        assert all(ws.cell(row=r["1.2.1"] + 1, column=c).value is None for c in range(1, 12)), sheet
    struct = wb[SHEET_STRUCT]
    assert struct.cell(row=r["1.2.1"] - 3, column=3).value == "1.2.1"
    assert struct.cell(row=r["1.2.1"] - 2, column=5).value == KIND_COST
    meta = {k: v for k, v in wb["_meta"].iter_rows(min_row=2, values_only=True)}
    assert meta["depuracion_costes"] is True and meta["depuracion_hoja"] == "PV-EmpresaA"
    assert issues_with(rep, "Depuración de costes de EmpresaA", SEV_INFO)


def test_single_company_book_reviews_the_pv_sheet(tmp_path):
    path = tmp_path / "libro.xlsx"
    make_book(path, config(multi=False))
    wb = load_workbook(path)
    r = rows(wb)
    pv = wb["PV"]
    assert pv.cell(row=r["2.1.0"] + 1, column=2).value == "Básica (coste)"
    validations = " ".join(str(dv.sqref) for dv in pv.data_validations.dataValidation)
    assert f"C{r['2.1.0']}" in validations and f"C{r['2.1.0'] + 1}" not in validations
    assert wb["% PV"].row_dimensions[r["2.1.0"] + 1].hidden


def test_update_and_analysis_match_a_book_without_review(tmp_path):
    """La depuración no cambia el PV, el EV ni los indicadores: las filas «(coste)» no forman parte del libro."""
    results = []
    for review in (True, False):
        path = tmp_path / f"libro_{review}.xlsx"
        make_book(path, config(review=review))
        for sheet, bac in (("PV-EmpresaA", 1000.0), ("PV-EmpresaB", 500.0)):
            set_bac(path, sheet, {c: bac for c in ("1.1.0", "1.2.1", "1.2.2", "1.0.0-D", "2.1.0")})
        progress = tasks()
        progress[1].percent_complete = 50
        progress[3].percent_complete = 100
        run_update(P(progress), config(review=review), path, date(2027, 2, 28), backup=False)
        run_analysis(config(review=review), path, backup=False)
        book = read_book(path)
        code = {s.row: s.code for s in book.structure}
        results.append(({code[r]: v for r, v in book.pv.items()}, {code[r]: v for r, v in book.ev.items()},
                        indicators_for(book, date(2027, 2, 28))))
        with zipfile.ZipFile(path) as z:            # los gráficos se vuelven a crear al guardar
            charts = [n for n in z.namelist() if n.startswith("xl/charts/chart")]
        assert len(charts) == (4 if review else 3)
    assert results[0][0] == results[1][0] and results[0][1] == results[1][1]
    assert results[0][2] == results[1][2] and results[0][2].ev > 0


def test_lookup_of_book_codes(tmp_path):
    rep = build_report(P(tasks()), config())
    entries, nonnumeric, repeated = cr.lookup_rows(rep)
    by = rows_by_code(rep)
    assert entries["1.2.1"] == (by["1.2.1"], by["1.2.1"])
    assert entries["1"] == (cr.NO_ROW, by["1.0.0"]) and entries["1.2"] == (cr.NO_ROW, by["1.2.0"])     # resúmenes
    assert "1.0.0-D" not in entries and not nonnumeric and not repeated
    path = tmp_path / "libro.xlsx"                              # la misma tabla, con las filas de la hoja
    make_book(path, config())
    r = rows(load_workbook(path))
    table = cr.structure_table(read_structure(load_workbook(path)))
    assert table["1.2.1"] == (r["1.2.1"], r["1.2.1"]) and table["1"] == (cr.NO_ROW, r["1.0.0"])
    assert len(table) == len(entries)


def test_plan_calculations_and_reading_in_python(tmp_path):
    """Los componentes se suman en su EDT interna, las EDT interna en su EDT cliente y cada EDT cliente en la de su
    código superior; solo se leen las EDT cliente con fila de ingreso del libro."""
    plan = cr.plan_lines(PLAN_DATA)
    assert [x.kind for x in plan[:5]] == [cr.KIND_CLIENT, cr.KIND_CLIENT, cr.KIND_INTERNAL, cr.KIND_COMPONENT,
                                          cr.KIND_COMPONENT]
    assert plan[0].children == [1, 5] and plan[1].children == [2] and plan[2].children == [3, 4]
    assert plan[6].children == [7, 10, 11] and plan[5].children == [6, 12, 13] and plan[14].kind == ""
    assert plan[15].parent is None and plan[16].parent is None
    done = PLAN
    assert [line[cr.P_EAC - 1] for line in done] == PLAN_EAC
    assert done[2][cr.P_BAC - 1] == 90 and done[7][cr.P_BAC - 1] == 40      # BAC de la EDT interna: sus componentes
    assert done[0][cr.P_FIRST_MONTH - 1:cr.P_FIRST_MONTH + 4] == [97, 80, 60, 29, 17]
    assert done[3] == PLAN_DATA[3][:cr.P_EAC - 1] + [40] + PLAN_DATA[3][cr.P_EAC:]     # los datos no cambian
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    wb = load_workbook(path)
    assert python_status(PLAN, wb) == PLAN_STATUS
    lines = cr.read_lines(PLAN, cr.structure_table(read_structure(wb)))
    r = rows(wb)
    for code, expected in PLAN_COST.items():
        line = next(x for x in lines if x.state == cr.STATE_READ and x.match == r[code])
        assert [v or 0 for v in line.months[:len(expected)]] == expected, code
    assert [x.key for x in lines if x.state in cr.UNREAD] == ["1.2.7", "2", "9.9"]


def test_plan_lines_of_internal_codes_and_repeated_clients():
    plan = cr.plan_lines([["1.1", "A"], ["1.1", "A", "GPR01", "Con la EDT cliente repetida"], [None, None, None, "X"],
                          [None, None, "XRF01"], ["XRF02", "Código no numérico en A"], [None, None, None, "Y"]])
    assert [x.kind for x in plan] == [cr.KIND_CLIENT, cr.KIND_INTERNAL, cr.KIND_COMPONENT, cr.KIND_INTERNAL,
                                      cr.KIND_INTERNAL, cr.KIND_COMPONENT]
    assert plan[0].children == [1, 3, 4] and plan[1].children == [2] and plan[4].children == [5]
    lines = cr.read_lines([["1.1", "A", None, None, None, None, 5], ["1.1.0", "A", None, None, None, None, 7]],
                          {"1.1": (9, 9)})
    assert [(x.state, x.match, x.repeated) for x in lines] == [(cr.STATE_READ, 9, True)] * 2


# ----------------------------------------------------------------------------
# --migrar-desde
# ----------------------------------------------------------------------------
LINES = [["1.1", "Dirección", None, None, 200, 200, 100, 100],
         ["1.2.1", "Licitación", "ABC01", "Pliegos", None, 50, 50]]


def test_migration_keeps_the_pasted_forecast_and_the_bac(tmp_path):
    old = tmp_path / "old.xlsx"
    make_book(old, config())
    set_bac(old, "PV-EmpresaA", {"1.1.0": 1234.0})
    paste(old, LINES)
    new = tmp_path / "new.xlsx"
    rep = make_book(new, config(), prior_from_book(old, "Programa"))
    ws = load_workbook(new)[cr.SHEET_PASTE]
    assert [ws.cell(row=cr.P_FIRST, column=c).value for c in range(1, 9)] == LINES[0]
    assert [ws.cell(row=cr.P_FIRST + 1, column=c).value for c in range(1, 5)] == LINES[1][:4]
    assert issues_with(rep, f"Se conservó el contenido de la hoja {cr.SHEET_PASTE} (2 líneas).")
    assert read_book(new).bac_role["A"][rows(load_workbook(new))["1.1.0"]] == 1234.0

    # Libro de entrega: sin depuración, conserva el BAC y no traslada la previsión
    final = tmp_path / "final.xlsx"
    rep = make_book(final, config(review=False), prior_from_book(new, "Programa"))
    assert cr.SHEET_PASTE not in load_workbook(final, read_only=True).sheetnames
    assert issues_with(rep, "no se trasladó: el libro nuevo no tiene depuración", SEV_INFO)
    book = read_book(final)
    assert book.bac_role["A"][rows(load_workbook(final))["1.1.0"]] == 1234.0
    assert sorted(s.row for s in book.structure) == list(range(5, 5 + len(book.structure)))


def test_migration_rewrites_the_plan_formulas(tmp_path):
    """La hoja con los cálculos del plan (fórmulas) pasa con sus datos, y las fórmulas se vuelven a escribir, también
    si cambia el mes 1 del horizonte."""
    old = tmp_path / "old.xlsx"
    make_book(old, config())
    set_bac(old, "PV-EmpresaA", BAC_A)
    paste(old, PLAN_DATA)
    run_cost_plan(config(cost_review_calcs=True), old, backup=False)
    later = tmp_path / "later.xlsx"
    rep = make_book(later, config(start_month=date(2027, 2, 1)), prior_from_book(old, "Programa"))
    assert issues_with(rep, "se volvieron a escribir las fórmulas de los cálculos del plan")
    ws = load_workbook(later)[cr.SHEET_PASTE]
    r1 = cr.P_FIRST + 1                             # 1.1.0: suma de su EDT interna (fila siguiente)
    assert ws.cell(row=r1, column=cr.P_FIRST_MONTH).value == f"=SUM(G{r1 + 1})"
    assert ws.cell(row=r1 + 1, column=cr.P_FIRST_MONTH).value == f"=SUM(G{r1 + 2}:G{r1 + 3})"
    assert ws.cell(row=r1, column=cr.P_BAC).value.startswith("=IF(")
    # Los datos (amarillo) se conservan, desplazados un mes: el mes 1 anterior queda fuera
    assert [ws.cell(row=r1 + 2, column=c).value for c in range(cr.P_FIRST_MONTH, cr.P_FIRST_MONTH + 4)] == [9, 8, 7, 6]
    assert ws.cell(row=r1 + 2, column=cr.P_BAC).fill.fgColor.rgb.endswith(FILL_INPUT.fgColor.rgb[-6:])
    assert ws.cell(row=r1, column=cr.P_EAC).fill.fgColor.rgb.endswith(cr.FILL_CALC.fgColor.rgb[-6:])


def test_migration_shifts_months_and_checks_the_company(tmp_path):
    old = tmp_path / "old.xlsx"
    make_book(old, config())
    paste(old, [["1.1", "Dirección", None, None, 300, 300, 100, 100, 100]])
    later = tmp_path / "later.xlsx"                 # el horizonte empieza un mes después
    rep = make_book(later, config(start_month=date(2027, 2, 1)), prior_from_book(old, "Programa"))
    ws = load_workbook(later)[cr.SHEET_PASTE]
    assert [ws.cell(row=cr.P_FIRST, column=c).value for c in range(5, 10)] == [300, 300, 100, 100, None]
    assert issues_with(rep, "se desplazaron 1 columnas", SEV_WARN) and issues_with(rep, "quedaron fuera 100,00")
    other = tmp_path / "other.xlsx"
    rep = make_book(other, config(cost_review_company="EmpresaB"), prior_from_book(old, "Programa"))
    assert issues_with(rep, "tiene la previsión de EmpresaA y el libro nuevo depura EmpresaB", SEV_WARN)
    assert load_workbook(other)[cr.SHEET_PASTE].cell(row=cr.P_FIRST, column=1).value is None


def test_migration_converts_a_paste_without_internal_wbs_columns(tmp_path):
    """Libro anterior a las columnas de EDT interna: código, descripción, BAC y meses desde la columna D."""
    old = tmp_path / "old.xlsx"
    make_book(old, config())
    wb = load_workbook(old)
    ws = wb[cr.SHEET_PASTE]
    ws.cell(row=cr.P_HEADER_ROW, column=3, value=cr.LEGACY_BAC_HEADER)
    for k, line in enumerate([[1.2, "Contratos", None], ["XRF01", "Subtarea", 50, 20, None, 30],
                              [None, "Nota", None], ["2.1", "Básica", 10, 10]]):
        for c, v in enumerate(line, start=1):
            ws.cell(row=cr.P_FIRST + k, column=c, value=v)
    wb.save(old)
    assert cr.paste_format(load_workbook(old)[cr.SHEET_PASTE]) == cr.FORMAT_LEGACY
    assert cr.last_paste_month(load_workbook(old)[cr.SHEET_PASTE], 24) == 3
    new = tmp_path / "new.xlsx"
    rep = make_book(new, config(), prior_from_book(old, "Programa"))
    assert issues_with(rep, "tenía el formato sin EDT interna", SEV_INFO)
    ws = load_workbook(new)[cr.SHEET_PASTE]
    assert cr.paste_format(ws) == cr.FORMAT_CURRENT
    got = [[ws.cell(row=cr.P_FIRST + k, column=c).value for c in range(1, 10)] for k in range(4)]
    assert got == [[1.2, "Contratos", None, None, None, None, None, None, None],
                   [None, None, "XRF01", "Subtarea", 50, 50, 20, None, 30],      # EAC: suma de los meses
                   [None, None, None, "Nota", None, None, None, None, None],
                   ["2.1", "Básica", None, None, 10, 10, 10, None, None]]
    assert cr.last_paste_month(ws, 24) == 3


def test_migration_adds_the_eac_column_of_a_version_2_4_paste(tmp_path):
    """Libro de la versión 2.4: sin la columna EAC, con los meses desde la columna F."""
    old = tmp_path / "old.xlsx"
    make_book(old, config())
    wb = load_workbook(old)
    ws = wb[cr.SHEET_PASTE]
    ws.cell(row=cr.P_HEADER_ROW, column=cr.P_EAC, value=datetime(2027, 1, 1))       # encabezado del mes 1
    for k, line in enumerate([["1.1", "Dirección", None, None, 200, 100, 100], [None, None, "ABC01", "Pliegos"]]):
        for c, v in enumerate(line, start=1):
            ws.cell(row=cr.P_FIRST + k, column=c, value=v)
    wb.save(old)
    assert cr.paste_format(load_workbook(old)[cr.SHEET_PASTE]) == cr.FORMAT_NO_EAC
    assert cr.last_paste_month(load_workbook(old)[cr.SHEET_PASTE], 24) == 2
    new = tmp_path / "new.xlsx"
    rep = make_book(new, config(), prior_from_book(old, "Programa"))
    assert issues_with(rep, "no tenía la columna EAC", SEV_INFO)
    ws = load_workbook(new)[cr.SHEET_PASTE]
    got = [[ws.cell(row=cr.P_FIRST + k, column=c).value for c in range(1, 9)] for k in range(2)]
    assert got == [["1.1", "Dirección", None, None, 200, 200, 100, 100],
                   [None, None, "ABC01", "Pliegos", None, None, None, None]]
    assert ws.cell(row=cr.P_HEADER_ROW, column=cr.P_EAC).value == "EAC" and cr.last_paste_month(ws, 24) == 2


# ----------------------------------------------------------------------------
# Fórmulas de Excel (LibreOffice)
# ----------------------------------------------------------------------------
def status_and_notes(ws, n: int) -> tuple[list, list]:
    status_col = cr.P_FIRST_MONTH + cr.paste_months(24) + 1
    return ([ws.cell(row=cr.P_FIRST + k, column=status_col).value or None for k in range(n)],
            [ws.cell(row=cr.P_FIRST + k, column=status_col + 1).value or None for k in range(n)])


@pytest.mark.libreoffice
def test_client_lines_are_read_into_the_cost_rows(tmp_path, recalc):
    """Con los cálculos hechos: las filas «(coste)» leen los meses y el EAC de las líneas de EDT cliente."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config(cost_review_alignment=True))
    set_bac(path, "PV-EmpresaA", BAC_A)
    paste(path, PLAN)
    wb = load_workbook(path)
    wb[cr.SHEET_CHART][cr.G_SELECT] = "1.2.0"
    wb.save(path)
    wv = load_workbook(recalc(path), data_only=True)
    ws = wv[cr.SHEET_PASTE]
    status, notes = status_and_notes(ws, len(PLAN))
    assert status == PLAN_STATUS == python_status(PLAN, wb)
    assert not any(notes)
    controls = [ws.cell(row=r, column=cr.P_EAC).value for r in (cr.P_TOTAL_ROW, cr.P_ASSIGNED_ROW, cr.P_MISSING_ROW)]
    assert controls == [117 + 65 + 96, 0, 5 + 50 + 7]
    pv_a = wv["PV-EmpresaA"]
    r = rows(wb)

    def cost(code: str, n: int = 5) -> list:
        return [pv_a.cell(row=r[code] + 1, column=7 + i).value or 0 for i in range(n)]

    for code, expected in PLAN_COST.items():
        assert cost(code, len(expected)) == expected, code
        assert pv_a.cell(row=r[code] + 1, column=4).value == sum(expected), code      # EAC de la línea
    assert cost("2.1.0") == [0] * 5 and cost("1.0.0-D") == [0] * 5
    assert cost("1.2.0", 3) == [71, 60, 30] and pv_a.cell(row=r["1.2.0"] + 1, column=4).value == 65 + 96
    assert cost("1.0.0") == [92, 80, 60, 29, 17]
    total = max(r.values()) + 2
    assert pv_a.cell(row=total + 1, column=4).value == 117 + 65 + 96
    # Indicador: diferencia entre el EAC y el BAC, en % del BAC
    last = 7 + 24
    assert pv_a.cell(row=r["1.1.0"] + 1, column=last + 1).value == pytest.approx(0)
    assert pv_a.cell(row=r["1.2.2"] + 1, column=last + 1).value == pytest.approx(-0.04)
    chart = wv[cr.SHEET_CHART]
    assert chart["B5"].value == pytest.approx(165) and chart["B6"].value == pytest.approx(161)
    assert chart.cell(row=cr.G_FIRST + 2, column=cr.G_COSTC).value == pytest.approx(161)
    assert chart["C4"].value == "Contratos"


@pytest.mark.libreoffice
def test_line_notes_and_controls(tmp_path, recalc):
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    lines = [
        ["1.1", "Dirección", None, None, None, None, 100, "50", *[None] * 27, 20],   # texto; mes 30: fuera
        ["1.1.0", "Dirección (repetida)", None, None, None, None, 1],
        ["1.2.1", "Licitación", None, None, None, None, 6],
    ]
    paste(path, lines)
    ws = load_workbook(recalc(path), data_only=True)[cr.SHEET_PASTE]
    status, notes = status_and_notes(ws, 3)
    assert status == ["1.1.0", "1.1.0", "1.2.1"] == python_status(lines, load_workbook(path))
    assert "repetida" in notes[0] and "como texto" in notes[0] and "fuera del horizonte" in notes[0]
    assert "repetida" in notes[1] and notes[2] is None
    controls = [ws.cell(row=r, column=cr.P_EAC).value for r in (cr.P_TOTAL_ROW, cr.P_ASSIGNED_ROW, cr.P_MISSING_ROW)]
    assert controls == [107, 20, 0]


@pytest.mark.libreoffice
def test_plan_formulas_match_python(tmp_path, recalc):
    """procesar-costes --calculos: las fórmulas que escribe dan lo mismo que los cálculos en Python, y el BAC de cada
    EDT cliente es el de la hoja PV-EmpresaA."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    set_bac(path, "PV-EmpresaA", BAC_A)
    paste(path, PLAN_DATA)
    res = run_cost_plan(config(cost_review_calcs=True), path, backup=False)
    assert res.calculated == 16
    wv = load_workbook(recalc(path), data_only=True)
    ws = wv[cr.SHEET_PASTE]
    n = cr.P_FIRST_MONTH + 4
    got = [[ws.cell(row=cr.P_FIRST + k, column=c).value or None for c in range(cr.P_FIRST_MONTH, n + 1)]
           for k in range(len(PLAN))]
    assert got == [[v or None for v in (line + [None] * n)[cr.P_FIRST_MONTH - 1:n]] for line in PLAN]
    assert [ws.cell(row=cr.P_FIRST + k, column=cr.P_EAC).value for k in range(len(PLAN))] == PLAN_EAC
    bac = [ws.cell(row=cr.P_FIRST + k, column=cr.P_BAC).value for k in range(len(PLAN))]
    assert bac[:3] == [117 + 65 + 100 + 10, 117, 90] and bac[6] == 65 and bac[12] == 100 and bac[15] == 50
    assert {bac[11], bac[13], bac[16]} <= {"", None}                    # EDT cliente sin fila del libro
    assert status_and_notes(ws, len(PLAN))[0] == PLAN_STATUS
    assert ws.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value == 117 + 65 + 96


@pytest.mark.libreoffice
def test_example_forecast_is_fully_read(tmp_path, generator, recalc):
    """El ayudante del ejemplo escribe un plan de costes ficticio (hitos de pago, subtareas, EDT interna con
    componentes y EDT cliente de los resúmenes) que se lee completo."""
    xml = tmp_path / "base.xml"
    generator.write(generator.build_plans(), xml, None)
    cfg = config(milestone_weight_field="Peso hito", method_field="Método EVM", plan="INTERNO")
    cfg.start_month = None
    from evm.reader_xml import read_xml
    from evm.wbs import build_part, assemble_report
    project = read_xml(xml, cfg)
    rep = assemble_report([build_part(project, cfg)], cfg)
    path = tmp_path / "ejemplo.xlsx"
    write_workbook(rep, cfg, path)
    generator.fill_book(path)
    ws = load_workbook(path)[cr.SHEET_PASTE]
    codes = [ws.cell(row=cr.P_FIRST + k, column=cr.P_INT_CODE).value for k in range(cr.P_LINES)]
    assert {"PER01", "HP-01", "XRF01", "GPR03"} <= set(codes)
    wv = load_workbook(recalc(path), data_only=True)
    pa = wv[cr.SHEET_PASTE]
    assert pa.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value > 0
    assert pa.cell(row=cr.P_MISSING_ROW, column=cr.P_EAC).value == pytest.approx(0)
    status, notes = status_and_notes(pa, cr.P_LINES)
    assert not any(notes) and cr.STATUS_SUMMARY in status and cr.STATUS_NONE not in status
    total = max(rows(load_workbook(path)).values()) + 2
    assert wv["PV-EmpresaA"].cell(row=total + 1, column=4).value == pytest.approx(
        pa.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value)


# ----------------------------------------------------------------------------
# procesar-costes
# ----------------------------------------------------------------------------
def processed_book(tmp_path, lines=PLAN):
    """Libro de depuración con el BAC de EmpresaA y el plan de costes (sin procesar)."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    set_bac(path, "PV-EmpresaA", BAC_A)
    set_bac(path, "PV-EmpresaB", {"1.1.0": 100.0, "2.1.0": 100.0})
    paste(path, lines)
    return path


def test_procesar_costes_reviews_the_plan_without_changing_the_book(tmp_path):
    """Sin --calculos, procesar-costes revisa el plan y no modifica el libro: el % planificado es el de MS Project."""
    path = processed_book(tmp_path)
    data = path.read_bytes()
    res = run_cost_plan(config(), path)
    assert path.read_bytes() == data and not res.saved and res.backup is None and res.calculated is None
    assert res.sheet == "PV-EmpresaA" and [x[0] for x in res.rows] == ["1.1.0", "1.2.1", "1.2.2"]
    assert res.rows[2][2:] == (96, 100)                                     # EAC de la línea y BAC de la fila
    assert res.lines == 9 and res.counted == 3 and res.inputs == 5
    assert res.total == pytest.approx(117 + 65 + 96) and res.assigned == pytest.approx(117 + 65 + 96)
    by = {i.id: i.message for i in res.warnings}
    assert set(by) == {"DEP-014", "DEP-016"}
    assert "1.2.7 (fila 21): sin fila del libro" in by["DEP-014"]
    assert "2.0.0 (fila 23): resumen sin detalle" in by["DEP-014"]
    assert "(2): 1.0.0-D, 2.1.0. Su fila «(coste)» queda vacía." in by["DEP-016"]


def test_procesar_costes_with_calculations(tmp_path):
    """--calculos: escribe las fórmulas del plan en la hoja PV-depuración y usa sus resultados; el % planificado
    no cambia."""
    path = processed_book(tmp_path, PLAN_DATA)
    before = read_book(path)
    res = run_cost_plan(config(cost_review_calcs=True), path, backup=False)
    assert res.saved and res.calculated == 16 and [x[0] for x in res.rows] == ["1.1.0", "1.2.1", "1.2.2"]
    assert [x[2] for x in res.rows] == [117, 65, 96]
    book = read_book(path)
    assert book.pct_pv == before.pct_pv and all(not s.plan_pv for s in book.structure)
    wb = load_workbook(path)
    assert "procesar-costes" not in wb["% PV-EmpresaA"]["A1"].value
    assert wb[SHEET_STRUCT].max_column == 8                 # sin la columna de las filas procesadas (antes de 2.6.0)
    ws = wb[cr.SHEET_PASTE]
    first = cr.P_FIRST
    assert ws.cell(row=first, column=cr.P_FIRST_MONTH).value == f"=SUM(G{first + 1},G{first + 5})"
    assert ws.cell(row=first + 3, column=cr.P_EAC).value == f"=SUM(G{first + 3}:DV{first + 3})"
    assert ws.cell(row=first + 3, column=cr.P_FIRST_MONTH).value == 10                  # dato: no cambia
    assert ws.cell(row=first + 14, column=cr.P_EAC).value is None                       # línea solo con descripción
    # Sin --calculos y sin los cálculos hechos: solo 1.2.2 tiene importes en su línea de EDT cliente
    (tmp_path / "sin").mkdir()
    res = run_cost_plan(config(), processed_book(tmp_path / "sin", PLAN_DATA), backup=False)
    assert [x[0] for x in res.rows] == ["1.2.2"]
    assert "(2): 1.1.0, 1.2.1. Su fila «(coste)» queda en 0." in next(
        i.message for i in res.warnings if i.id == "DEP-017")


def test_procesar_costes_notices(tmp_path):
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    set_bac(path, "PV-EmpresaA", {"1.1.0": 300.0, "1.2.1": 10.0, "1.2.2": 10.0})
    paste(path, [
        ["1.1", "Dirección", None, None, None, None, *months((1, 100), (30, 50))],     # mes 30: fuera del horizonte
        ["1.1.0", "Dirección", None, None, None, None],                               # repetida
        ["1.2.1", "Licitación", None, None, None, None, 30, -50],                     # EAC negativo (abono)
        ["1.2.2", "Obras"],                                                           # sin importes
        ["2.1", "Básica", None, None, None, None, "40", "=10+5", 20],                  # texto, fórmula y sin BAC
    ])
    res = run_cost_plan(config(), path, backup=False)
    by = {i.id: i.message for i in res.warnings}
    assert set(by) == {"DEP-015", "DEP-016", "DEP-017", "DEP-018", "DEP-019"}
    assert "50,00" in by["DEP-015"] and "(1.1.0)" in by["DEP-015"] and "pero no de sus meses" in by["DEP-015"]
    assert "(1): 1.2.2" in by["DEP-017"] and "--calculos" in by["DEP-017"] and "1.1" in by["DEP-019"]
    assert "importes guardados como texto" in by["DEP-018"] and "1 celda tiene una fórmula" in by["DEP-018"]
    eac = {code: (e, bac) for code, _name, e, bac in res.rows}
    assert eac == {"1.1.0": (150, 300), "1.2.1": (-20, 10), "2.1.0": (20, 0)}


def test_procesar_costes_simulation_and_errors(tmp_path):
    path = processed_book(tmp_path)
    data = path.read_bytes()
    res = run_cost_plan(config(cost_review_calcs=True), path, dry_run=True)
    assert res.dry_run and not res.saved and res.backup is None and len(res.rows) == 3 and path.read_bytes() == data
    empty = tmp_path / "vacio.xlsx"
    make_book(empty, config())
    with pytest.raises(InputError) as exc:
        run_cost_plan(config(), empty)
    assert exc.value.id == "DEP-013"
    plain = tmp_path / "sin_depuracion.xlsx"
    make_book(plain, config(review=False))
    with pytest.raises(InputError) as exc:
        run_cost_plan(config(), plain)
    assert exc.value.id == "DEP-012"


def test_procesar_costes_command(tmp_path, capsys):
    path = processed_book(tmp_path, PLAN_DATA)
    assert main(["procesar-costes", "--libro", str(path)]) == 0
    out = capsys.readouterr().out
    assert "No se modificó el libro (los cálculos del plan se escriben con --calculos)." in out
    assert main(["procesar-costes", "--libro", str(path), "--sin-respaldo", "--calculos"]) == 0
    out = capsys.readouterr().out
    assert "Cálculos del plan de costes: 16 líneas" in out and "Plan de costes en 3 de 5 filas de ingreso" in out
    assert "DEP-014" in out and "1.1.0" in out and "Libro actualizado" in out
    assert "El % planificado (hoja '% PV-EmpresaA') no cambia: es el de MS Project." in out
    line = next(x for x in out.splitlines() if x.strip().startswith("1.2.2"))
    assert "EAC" in line and line.split("EAC")[1].split("BAC")[0].strip() == "96" and line.endswith(" 100")
    plain = tmp_path / "sin_depuracion.xlsx"
    make_book(plain, config(review=False))
    assert main(["procesar-costes", "--libro", str(plain)]) == 2
    assert "DEP-012" in capsys.readouterr().err


def legacy_processed_book(tmp_path):
    """Libro procesado con procesar-costes antes de la versión 2.6.0: el reparto del plan en la hoja % PV-EmpresaA
    (importe / EAC) y las filas marcadas en la columna 9 de _estructura."""
    path = processed_book(tmp_path)
    wb = load_workbook(path)
    r = rows(wb)
    ws_pct, ws_struct = wb["% PV-EmpresaA"], wb[SHEET_STRUCT]
    ws_struct.cell(row=1, column=9, value="pv_plan_costes")
    struct_row = {ws_struct.cell(row=k, column=3).value: k for k in range(2, ws_struct.max_row + 1)}
    for code, amounts in PLAN_COST.items():
        for i in range(24):
            v = amounts[i] / sum(amounts) if i < len(amounts) else None
            ws_pct.cell(row=r[code], column=7 + i).value = v
        ws_struct.cell(row=struct_row[code], column=9, value="PV-A")
    wb.save(path)
    return path


def test_legacy_processed_book_is_reported_and_migration_recovers_ms_project(tmp_path):
    fresh = tmp_path / "fresh.xlsx"                 # % planificado de MS Project
    make_book(fresh, config())
    ms_project = read_book(fresh)
    path = legacy_processed_book(tmp_path)
    old = read_book(path)
    assert {s.code: s.plan_pv for s in old.structure}["1.1.0"] == ("A",)
    r_old = rows(load_workbook(path))
    assert old.pct_pv["A"][r_old["1.2.2"]][:3] == pytest.approx([56 / 96, 40 / 96, 0])
    # procesar-costes y actualizar avisan; actualizar no compara esas filas con la línea base (ACT-047)
    res = run_cost_plan(config(), path)
    dep = next(i.message for i in res.warnings if i.id == "DEP-021")
    assert "3 filas de la hoja % PV-EmpresaA (1.1.0, 1.2.1, 1.2.2)" in dep and "--migrar-desde" in dep
    res = run_update(P(tasks()), config(), path, date(2027, 2, 28), backup=False)
    assert [w for w in res.warnings if w.id == "DEP-021"]
    assert not [w for w in res.warnings if w.id == "ACT-047"]
    # La otra empresa sí se compara: un % planificado distinto del de MS Project se informa
    wb = load_workbook(path)
    wb["% PV-EmpresaB"].cell(row=rows(wb)["1.1.0"], column=7, value=0.9)
    wb.save(path)
    res = run_update(P(tasks()), config(), path, date(2027, 2, 28), backup=False)
    drift = [w for w in res.warnings if w.id == "ACT-047"]
    assert len(drift) == 1 and "(1.1.0)" in drift[0].message
    # --migrar-desde: el libro nuevo (también el de entrega) tiene el % de MS Project y conserva el BAC
    for review, name in ((True, "new.xlsx"), (False, "final.xlsx")):
        new = tmp_path / name
        rep = make_book(new, config(review=review), prior_from_book(path, "Programa"))
        found = issues_with(rep, "3 filas de la hoja % PV-EmpresaA tenían en el libro anterior el reparto del plan "
                                 "de costes", SEV_INFO)
        assert found and found[0].id == "DEP-020"
        book = read_book(new)
        r_new, r_fresh = rows(load_workbook(new)), rows(load_workbook(fresh))
        for code in ("1.1.0", "1.2.1", "1.2.2", "2.1.0", "1.0.0-D"):
            assert book.pct_pv["A"][r_new[code]] == pytest.approx(ms_project.pct_pv["A"][r_fresh[code]]), code
            assert book.pct_pv["B"][r_new[code]] == pytest.approx(ms_project.pct_pv["B"][r_fresh[code]]), code
        assert all(not s.plan_pv for s in book.structure)
        assert book.bac_role["A"][r_new["1.2.2"]] == 100.0
        assert "procesar-costes" not in load_workbook(new)["% PV-EmpresaA"]["A1"].value
        assert not [w for w in run_update(P(tasks()), config(review=review), new, date(2027, 2, 28),
                                          backup=False).warnings if w.id in ("DEP-021", "ACT-047")]
