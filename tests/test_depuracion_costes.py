"""Depuración de costes: filas «(coste)» con la previsión pegada en la hoja PV-depuración."""
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
from evm.workbook import COLOR_COST, FILL_COST, KIND_COST, SHEET_STRUCT, read_structure, write_workbook

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
    """Columna 'Asignado a' según la asignación en Python de procesar-costes (cost_review.assign_lines)."""
    structure = read_structure(wb)
    code = {s.row: s.code for s in structure}
    out = []
    for x in cr.assign_lines(lines, cr.structure_table(structure)):
        if not (x.numeric or x.child):
            out.append(None)
        elif x.subtotal:
            out.append(cr.STATUS_SUBTOTAL)
        elif x.target > 0:
            out.append(code[x.target])
        else:
            out.append(cr.STATUS_SUMMARY if x.target == cr.NO_ROW else cr.STATUS_NONE)
    return out


# Plan de costes en tres niveles, como el de la empresa: EDT cliente, EDT interna (GPR00) y componentes de coste
# (solo la columna D)
PLAN = [
    [None, None, None, "Componente sin EDT cliente encima", None, None, 3],
    ["1.0.0", "Gestión", None, None, 317, None],                                 # subtotal: códigos más detallados
    ["1.1.0", "Dirección", None, None, 150, 157],                                # subtotal: EDT interna debajo
    [None, None, "GPR00", "Lanzamiento", 150, 157, 31, 30, 39, 35, 22],          # subtotal de sus componentes
    [None, None, None, "Costes indirectos X", 30, 40, 10, 9, 8, 7, 6],           # -> 1.1.0
    [None, None, None, "Costes de gestión Y", 60, 77, 11, 11, 22, 22, 11],
    [None, None, None, "Costes indirectos Y", 30, 19, 5, 5, 4, 3, 2],
    [None, None, None, "Otros", 30, 21, 5, 5, 5, 3, 3],
    ["1.2.1", "Licitación", None, None, 60, 60],                                 # subtotal
    [None, None, "GPR01", "Pliegos", 40, 40],                                    # subtotal, sin importes propios
    [None, None, None, "Personal", 30, 30, 10, 10, 10],                          # -> 1.2.1
    [None, None, None, "Viajes", 10, 10, None, 10],
    [None, None, "GPR02", "EDT interna sin componentes", 20, 20, None, None, 20],  # -> 1.2.1
    ["1.2.2", "Obras", "GPR03", "Obra civil", None, None, 99],                   # subtotal con importes propios
    ["1.2.2", "Obras", None, "Hormigón", None, None, 40, 40],                    # componente con la EDT cliente
    [None, None, None, "Acero", None, None, 9],                                  # -> 1.2.2
    [None, None, None, "Componente sin importes"],
    [None, None, "GPR04", None, None, None, 7],                                  # -> 1.2.2
    ["2.1", "Básica", None, None, 50, 50],                                       # subtotal: componente debajo
    [None, None, None, "Componente sin EDT interna", 50, 50, None, 25, 25],      # -> 2.1.0
]
PLAN_STATUS = [cr.STATUS_NONE, cr.STATUS_SUBTOTAL, cr.STATUS_SUBTOTAL, cr.STATUS_SUBTOTAL, "1.1.0", "1.1.0",
               "1.1.0", "1.1.0", cr.STATUS_SUBTOTAL, cr.STATUS_SUBTOTAL, "1.2.1", "1.2.1", "1.2.1",
               cr.STATUS_SUBTOTAL, "1.2.2", "1.2.2", None, "1.2.2", cr.STATUS_SUBTOTAL, "2.1.0"]
PLAN_COST = {"1.1.0": [31, 30, 39, 35, 22], "1.2.1": [10, 20, 30], "1.2.2": [56, 40], "2.1.0": [0, 25, 25]}


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
                    'depuracion_alineamiento = true\ndepuracion_umbral = 15\n\n[multiempresa]\nactivar = true\n',
                    encoding="utf-8")
    cfg = load_config(path)
    assert cfg.cost_review and cfg.cost_review_company == "EmpresaB" and cfg.cost_review_role == "B"
    assert cfg.cost_review_alignment and cfg.cost_review_threshold == 15
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
        assert pv_a.cell(row=row + 1, column=4).value.startswith("=SUM(")
    assert pv_a.cell(row=r["1.2.1"] + 1, column=7).value == cr.paste_sum(0, r["1.2.1"])
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
    assert entries["1"] == (cr.NO_ROW, by["1.0.0-D"])          # resumen: lo que depende va a «Otras tareas»
    assert entries["1.2"] == (cr.NO_ROW, cr.NO_ROW)            # resumen sin fila propia
    assert "1.0.0-D" not in entries and not nonnumeric and not repeated
    path = tmp_path / "libro.xlsx"                              # la misma tabla, con las filas de la hoja
    make_book(path, config())
    r = rows(load_workbook(path))
    table = cr.structure_table(read_structure(load_workbook(path)))
    assert table["1.2.1"] == (r["1.2.1"], r["1.2.1"]) and table["1"] == (cr.NO_ROW, r["1.0.0-D"])
    assert table["1.2"] == (cr.NO_ROW, cr.NO_ROW) and len(table) == len(entries)


def test_three_level_plan_is_assigned_in_python(tmp_path):
    """EDT cliente, EDT interna y componentes de coste: la EDT interna con componentes es un subtotal."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    wb = load_workbook(path)
    assert python_status(PLAN, wb) == PLAN_STATUS
    lines = cr.assign_lines(PLAN, cr.structure_table(read_structure(wb)))
    r = rows(wb)
    for code, expected in PLAN_COST.items():
        got = [sum((x.months[i] or 0) for x in lines if x.target == r[code] and i < len(x.months))
               for i in range(len(expected))]
        assert got == expected, code
    assert lines[3].subtotal and lines[3].int_code and lines[4].component and lines[4].internal_row == lines[3].row
    assert lines[12].int_code and not lines[12].subtotal          # EDT interna sin componentes: se suma
    assert not lines[16].counted                                   # componente sin importes: se ignora


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
    assert issues_with(rep, f"Se conservó el contenido de la hoja {cr.SHEET_PASTE} (2 líneas)")
    assert read_book(new).bac_role["A"][rows(load_workbook(new))["1.1.0"]] == 1234.0

    # Libro de entrega: sin depuración, conserva el BAC y no traslada la previsión
    final = tmp_path / "final.xlsx"
    rep = make_book(final, config(review=False), prior_from_book(new, "Programa"))
    assert cr.SHEET_PASTE not in load_workbook(final, read_only=True).sheetnames
    assert issues_with(rep, "no se trasladó: el libro nuevo no tiene depuración", SEV_INFO)
    book = read_book(final)
    assert book.bac_role["A"][rows(load_workbook(final))["1.1.0"]] == 1234.0
    assert sorted(s.row for s in book.structure) == list(range(5, 5 + len(book.structure)))


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
@pytest.mark.libreoffice
def test_pasted_lines_are_assigned_to_book_rows(tmp_path, recalc):
    path = tmp_path / "libro.xlsx"
    make_book(path, config(cost_review_alignment=True))
    set_bac(path, "PV-EmpresaA", {"1.1.0": 200.0, "1.2.1": 50.0, "1.2.2": 60.0, "1.0.0-D": 10.0, "2.1.0": 400.0})
    lines = [   # EDT cliente, actividad, EDT interna, actividad interna, BAC, EAC, meses
        [None, None, "ABC00", "EDT interna sin EDT cliente encima", None, None, *months((1, 3))],
        ["1", "Gestión (cabecera, sin importes)"],
        [None, None, "XRF01", "Subtarea suelta de Gestión", None, None, *months((3, 10))],     # -> 1.0.0-D
        ["1.1", "Dirección", None, None, 200, None, 100, 100, "50"],                          # 1.1 = 1.1.0; texto
        [1.2, "Contratos (subtotal con importes propios)", None, None, None, None, *months((1, 999))],
        ["1.2.1.5", "Licitación, tarea de detalle", None, None, None, None, *months((1, 50))],  # -> 1.2.1
        ["1.2.2", "Obras (BAC distinto del EAC: no se compara)", None, None, 70, 60, *months((2, 30), (3, 30))],
        [None, None, "ABC01", "Personal propio", 30, 30, *months((2, 30))],                   # -> 1.2.2
        [None, None, "ABC02", "Subcontratas", None, None, *months((3, 30))],
        [None, None, None, None, None],
        [None, None, None, "Nota sin código ni importes"],
        ["2.1", "Básica", None, None, 400, None, *months((2, 200), (30, 200))],             # EDT cliente repetida
        ["2.1", "Básica", "ABC03", "Ingeniería", None, None, *months((2, 150), (30, 200))],  # mes 30: fuera
        ["2.1", "Básica", "ABC04", "Revisión", None, None, *months((2, 50))],
        ["1.2.7", "Tarea de un resumen sin fila propia", None, None, None, None, *months((1, 5))],
        ["9.9", "Código inexistente", None, None, None, None, *months((1, 7))],
        ["1", "Gestión", "ABC05", "Gastos generales", None, None, *months((3, 4))],           # -> 1.0.0-D
        ["1.2.1", "Licitación", "ABC06", "Pliegos", 6, 6, *months((2, 6))],                   # sin línea de subtotal
        ["1.2.1", "Licitación", "ABC07", "Ofertas", 4, 4, *months((3, 4))],
    ]
    paste(path, lines)
    wb = load_workbook(path)
    wb[cr.SHEET_CHART][cr.G_SELECT] = "1.2.0"
    wb.save(path)
    wv = load_workbook(recalc(path), data_only=True)
    ws = wv[cr.SHEET_PASTE]
    status_col = cr.P_FIRST_MONTH + cr.paste_months(24) + 1
    status = [ws.cell(row=cr.P_FIRST + k, column=status_col).value or None for k in range(len(lines))]
    assert status == [cr.STATUS_NONE, cr.STATUS_SUBTOTAL, "1.0.0-D", "1.1.0", cr.STATUS_SUBTOTAL, "1.2.1",
                      cr.STATUS_SUBTOTAL, "1.2.2", "1.2.2", None, None, cr.STATUS_SUBTOTAL, "2.1.0", "2.1.0",
                      cr.STATUS_SUMMARY, cr.STATUS_NONE, "1.0.0-D", "1.2.1", "1.2.1"]
    assert status == python_status(lines, wb)
    obs = [ws.cell(row=cr.P_FIRST + k, column=status_col + 1).value for k in range(len(lines))]
    assert "Sin EDT cliente" in obs[0] and "como texto" in obs[3] and "Subtotal distinto" in obs[4]
    assert "fuera del horizonte" in obs[12]
    assert not any(obs[k] for k in (1, 2, 5, 6, 7, 8, 11, 13, 16, 17, 18))
    pasted = 3 + 10 + 200 + 50 + 60 + 400 + 5 + 7 + 4 + 10
    assert ws.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value == pytest.approx(pasted)
    assert ws.cell(row=cr.P_MISSING_ROW, column=cr.P_EAC).value == pytest.approx(3 + 200 + 5 + 7)

    pv_a = wv["PV-EmpresaA"]
    r = rows(wb)

    def cost(code: str) -> list:
        return [pv_a.cell(row=r[code] + 1, column=7 + i).value or 0 for i in range(3)]

    assert cost("1.1.0") == [100, 100, 0] and cost("1.2.1") == [50, 6, 4] and cost("1.2.2") == [0, 30, 30]
    assert cost("1.0.0-D") == [0, 0, 14] and cost("2.1.0") == [0, 200, 0]
    assert cost("1.2.0") == [50, 36, 34] and cost("1.0.0") == [150, 136, 48]
    total = max(r.values()) + 2
    assert pv_a.cell(row=total + 1, column=4).value == pytest.approx(pasted - 3 - 200 - 5 - 7)
    assert pv_a.cell(row=r["1.2.2"] + 1, column=4).value == 60
    # Indicador: diferencia de BAC y diferencia máxima acumulada en % del BAC
    last = 7 + 24
    assert pv_a.cell(row=r["1.1.0"] + 1, column=last + 1).value == pytest.approx(0)
    assert pv_a.cell(row=r["2.1.0"] + 1, column=last + 1).value == pytest.approx(-0.5)
    assert 0 < pv_a.cell(row=r["2.1.0"] + 1, column=last + 2).value <= 1
    assert pv_a.cell(row=r["1.1.0"], column=last + 1).value is None

    chart = wv[cr.SHEET_CHART]
    assert chart["B5"].value == pytest.approx(110) and chart["B6"].value == pytest.approx(120)
    assert chart.cell(row=cr.G_FIRST + 2, column=cr.G_COSTC).value == pytest.approx(120)
    assert chart["C4"].value == "Contratos"


@pytest.mark.libreoffice
def test_example_forecast_is_fully_assigned(tmp_path, generator, recalc):
    """El ayudante del ejemplo pega una previsión ficticia (hitos de pago y subtareas) que se asigna completa."""
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
    assert {"PER01", "HP-01", "XRF01"} <= set(codes)
    assert "GPR03" in codes                                     # EDT interna con componentes de coste
    wv = load_workbook(recalc(path), data_only=True)
    pa = wv[cr.SHEET_PASTE]
    assert pa.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value > 0
    assert pa.cell(row=cr.P_MISSING_ROW, column=cr.P_EAC).value == pytest.approx(0)
    obs_col = cr.P_FIRST_MONTH + cr.paste_months(24) + 2
    assert not any(pa.cell(row=cr.P_FIRST + k, column=obs_col).value for k in range(cr.P_LINES))   # todo cuadra
    total = max(rows(load_workbook(path)).values()) + 2
    assert wv["PV-EmpresaA"].cell(row=total + 1, column=4).value == pytest.approx(
        pa.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value)


@pytest.mark.libreoffice
def test_three_level_plan_formulas_match_python(tmp_path, recalc):
    """Las fórmulas de la hoja _depuracion y la asignación en Python de procesar-costes dan lo mismo."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    paste(path, PLAN)
    wv = load_workbook(recalc(path), data_only=True)
    ws = wv[cr.SHEET_PASTE]
    status_col = cr.P_FIRST_MONTH + cr.paste_months(24) + 1
    assert [ws.cell(row=cr.P_FIRST + k, column=status_col).value or None for k in range(len(PLAN))] == PLAN_STATUS
    obs = [ws.cell(row=cr.P_FIRST + k, column=status_col + 1).value for k in range(len(PLAN))]
    assert "Sin EDT cliente" in obs[0] and "Subtotal distinto" in obs[13]      # GPR03: 99 frente a 89
    assert not any(obs[k] for k in range(1, len(PLAN)) if k != 13)
    pv_a = wv["PV-EmpresaA"]
    r = rows(load_workbook(path))
    for code, expected in PLAN_COST.items():
        assert [pv_a.cell(row=r[code] + 1, column=7 + i).value or 0 for i in range(len(expected))] == expected, code
    assert ws.cell(row=cr.P_TOTAL_ROW, column=cr.P_EAC).value == pytest.approx(3 + 157 + 60 + 96 + 50)
    assert ws.cell(row=cr.P_MISSING_ROW, column=cr.P_EAC).value == pytest.approx(3)


# ----------------------------------------------------------------------------
# procesar-costes
# ----------------------------------------------------------------------------
def processed_book(tmp_path):
    """Libro de depuración con el BAC de EmpresaA y el plan de costes PLAN pegado (sin procesar)."""
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    set_bac(path, "PV-EmpresaA", {"1.1.0": 157.0, "1.2.1": 50.0, "1.2.2": 96.0, "2.1.0": 50.0, "1.0.0-D": 10.0})
    set_bac(path, "PV-EmpresaB", {"1.1.0": 100.0, "2.1.0": 100.0})
    paste(path, PLAN)
    return path


def test_procesar_costes_writes_the_plan_distribution(tmp_path):
    path = processed_book(tmp_path)
    before = read_book(path)
    res = run_cost_plan(config(), path)
    assert res.backup is not None and res.backup.exists()
    assert res.sheet == "PV-EmpresaA" and [c for c, _n, _t in res.rows] == ["1.1.0", "1.2.1", "1.2.2", "2.1.0"]
    assert res.lines == 19 and res.inputs == 5
    assert res.total == pytest.approx(3 + 157 + 60 + 96 + 50) and res.assigned == pytest.approx(157 + 60 + 96 + 50)
    book = read_book(path)
    r = rows(load_workbook(path))
    pct = book.pct_pv["A"]
    assert pct[r["1.1.0"]][:6] == pytest.approx([31 / 157, 30 / 157, 39 / 157, 35 / 157, 22 / 157, 0])
    # PV = BAC × %: con el BAC igual al EAC, el importe del plan; con otro BAC, el reparto del plan
    pv_a = {s.code: [book.bac_role["A"][s.row] * v for v in pct[s.row]] for s in book.input_rows}
    assert pv_a["1.1.0"][:5] == pytest.approx([31, 30, 39, 35, 22])
    assert pv_a["1.2.1"][:3] == pytest.approx([50 * 10 / 60, 50 * 20 / 60, 50 * 30 / 60])
    # Sin líneas en el plan: conserva el % de MS Project; la otra empresa no cambia
    assert pct[r["1.0.0-D"]] == before.pct_pv["A"][r["1.0.0-D"]] and book.pct_pv["B"] == before.pct_pv["B"]
    marks = {s.code: s.plan_pv for s in book.structure}
    assert marks["1.1.0"] == ("A",) and marks["1.0.0-D"] == ()
    wb = load_workbook(path)
    assert "procesar-costes" in wb["% PV-EmpresaA"]["A1"].value
    assert "procesar-costes" not in wb["% PV-EmpresaB"]["A1"].value
    assert wb["% PV-EmpresaA"].cell(row=r["1.1.0"], column=7).font.color.rgb.endswith(COLOR_COST)
    assert {"DEP-014", "DEP-016"} <= {i.id for i in res.warnings}
    assert "Componente sin EDT cliente encima (fila 8)" in next(i.message for i in res.warnings if i.id == "DEP-014")


def test_procesar_costes_notices(tmp_path):
    path = tmp_path / "libro.xlsx"
    make_book(path, config())
    paste(path, [
        ["1.1", "Dirección", None, None, None, None, *months((1, 100), (30, 50))],     # mes 30: fuera del horizonte
        ["1.2.1", "Licitación", None, None, None, None, 10, -10],                     # suma 0
        ["2.1", "Básica", None, None, None, None, "40", "=10+5", 20],                  # texto y fórmula sin valor
    ])
    res = run_cost_plan(config(), path, backup=False)
    by = {i.id: i.message for i in res.warnings}
    assert "50,00" in by["DEP-015"] and "(1.1.0)" in by["DEP-015"] and "1.2.1" in by["DEP-017"]
    assert "importes guardados como texto" in by["DEP-018"] and "1 celda tiene una fórmula" in by["DEP-018"]
    book = read_book(path)
    r = rows(load_workbook(path))
    assert sum(book.pct_pv["A"][r["1.1.0"]]) == pytest.approx(100 / 150)      # lo posterior no se reparte
    assert book.pct_pv["A"][r["2.1.0"]][:3] == pytest.approx([0, 0, 1])
    assert {s.code: s.plan_pv for s in book.structure}["1.2.1"] == ()


def test_procesar_costes_simulation_and_errors(tmp_path):
    path = processed_book(tmp_path)
    data = path.read_bytes()
    res = run_cost_plan(config(), path, dry_run=True)
    assert res.dry_run and res.backup is None and len(res.rows) == 4 and path.read_bytes() == data
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
    path = processed_book(tmp_path)
    assert main(["procesar-costes", "--libro", str(path), "--sin-respaldo"]) == 0
    out = capsys.readouterr().out
    assert "Reparto del plan de costes en 4 de 5 filas de ingreso" in out and "DEP-014" in out
    assert "1.1.0" in out and "Libro actualizado" in out
    plain = tmp_path / "sin_depuracion.xlsx"
    make_book(plain, config(review=False))
    assert main(["procesar-costes", "--libro", str(plain)]) == 2
    assert "DEP-012" in capsys.readouterr().err


def test_delivery_book_keeps_the_plan_distribution(tmp_path):
    path = processed_book(tmp_path)
    run_cost_plan(config(), path, backup=False)
    processed = read_book(path)
    old = rows(load_workbook(path))
    final = tmp_path / "final.xlsx"                 # libro de entrega: sin depuración
    rep = make_book(final, config(review=False), prior_from_book(path, "Programa"))
    assert issues_with(rep, "reparto del plan de costes (procesar-costes) en 4 filas de la hoja % PV-EmpresaA",
                       SEV_INFO)
    book = read_book(final)
    new = rows(load_workbook(final))
    for code in ("1.1.0", "1.2.1", "1.2.2", "2.1.0", "1.0.0-D"):    # 1.0.0-D: el de MS Project
        assert book.pct_pv["A"][new[code]] == pytest.approx(processed.pct_pv["A"][old[code]]), code
    assert {s.code: s.plan_pv for s in book.structure}["1.1.0"] == ("A",)
    assert "procesar-costes" in load_workbook(final)["% PV-EmpresaA"]["A1"].value
    # El horizonte empieza un mes después: el mes 1 anterior se suma en el mes 1 nuevo
    later = tmp_path / "later.xlsx"
    make_book(later, config(review=False, start_month=date(2027, 2, 1)), prior_from_book(final, "Programa"))
    pct = processed.pct_pv["A"][old["1.1.0"]]
    assert read_book(later).pct_pv["A"][rows(load_workbook(later))["1.1.0"]][:4] == pytest.approx(
        [pct[0] + pct[1], pct[2], pct[3], pct[4]])


def test_update_does_not_compare_the_plan_distribution_with_ms_project(tmp_path):
    path = processed_book(tmp_path)
    run_cost_plan(config(), path, backup=False)
    res = run_update(P(tasks()), config(), path, date(2027, 2, 28), backup=False)
    assert not [w for w in res.warnings if w.id == "ACT-047"]
    # La otra empresa sí se compara: un % planificado distinto del de MS Project se informa
    wb = load_workbook(path)
    wb["% PV-EmpresaB"].cell(row=rows(wb)["1.1.0"], column=7, value=0.9)
    wb.save(path)
    res = run_update(P(tasks()), config(), path, date(2027, 2, 28), backup=False)
    drift = [w for w in res.warnings if w.id == "ACT-047"]
    assert len(drift) == 1 and "(1.1.0)" in drift[0].message
