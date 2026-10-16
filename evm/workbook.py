"""Generación y lectura del libro Excel de Valor Ganado."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.cell.cell import ERROR_CODES, ILLEGAL_CHARACTERS_RE
from openpyxl.chart.axis import DateAxis
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline

from evm import __version__
from evm.config import Config
from evm.timeline import round_preserving_total
from evm.wbs import KIND_LEAF, KIND_RESIDUAL, KIND_SUMMARY, Report, ReportRow

# ----------------------------------------------------------------------------
# Disposición del libro (no modificar: el comando "actualizar" depende de ella)
# ----------------------------------------------------------------------------
SHEET_SUMMARY = "Resumen"
SHEET_PV = "PV"
SHEET_EV = "EV"
SHEET_AC = "AC"
SHEET_CURVE = "Curva S"
SHEET_IND = "Indicadores EDT"
SHEET_VALID = "Validación"
SHEET_LISTS = "Listas"
SHEET_STRUCT = "_estructura"
SHEET_META = "_meta"

TITLE_ROW, INFO_ROW, MONTHNUM_ROW, HEADER_ROW, FIRST_ROW = 1, 2, 3, 4, 5
COL_CODE, COL_NAME, COL_METHOD, COL_BAC, COL_START, COL_FINISH = 1, 2, 3, 4, 5, 6
FIRST_MONTH_COL = 7  # columna G

# Curva S (tabla vertical)
CURVE_HEADER_ROW = 4
CURVE_FIRST_ROW = 5
(C_NUM, C_DATE, C_PV, C_PVC, C_EV, C_EVC, C_AC, C_ACC,
 C_PVPCT, C_EVPCT, C_SPI, C_CPI, C_BAC, C_EVN, C_ACN) = range(1, 16)
CHART_ANCHOR = "Q4"

# Resumen
STATUS_CELL = "B5"
NAME_METHODS = "ListaMetodos"
START_CELL = "B7"
NAME_STATUS = "FechaCorte"
NAME_START = "InicioHorizonte"

# ----------------------------------------------------------------------------
# Estilos
# ----------------------------------------------------------------------------
DARK = "1F3864"
FILL_HEADER = PatternFill("solid", fgColor=DARK)
FILL_L1 = PatternFill("solid", fgColor="BDD7EE")
FILL_L2 = PatternFill("solid", fgColor="DDEBF7")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_TOTAL = PatternFill("solid", fgColor="D9D9D9")
# Colores de semáforo (fgColor y bgColor: Excel usa bgColor en el formato condicional)
FILL_RED = PatternFill("solid", fgColor="F8CBAD", bgColor="F8CBAD")
FILL_YELLOW = PatternFill("solid", fgColor="FFE699", bgColor="FFE699")
FILL_GREEN = PatternFill("solid", fgColor="C6E0B4", bgColor="C6E0B4")
FILL_SECTION = PatternFill("solid", fgColor="D6DCE4")
FONT_HEADER = Font(bold=True, color="FFFFFF")
FONT_TITLE = Font(bold=True, size=14, color=DARK)
FONT_BOLD = Font(bold=True)
FONT_NOTE = Font(italic=True, size=9, color="595959")
FONT_ITALIC = Font(italic=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_DATE = "dd/mm/yyyy"
FMT_MONTH = "mmm-yyyy"
FMT_RATIO = "0.00"
FMT_PCT = "0.0%"

METHOD_HELP = {
    "Por definir": "Aún no se asignó método. Para calcular el EV se usa % completado.",
    "0/100": "0 % al comenzar y 100 % al terminar. Para tareas que empiezan y terminan en el mismo mes.",
    "50/50": "50 % al comenzar y 50 % al terminar. Para tareas de hasta 2 meses.",
    "25/75": "25 % al comenzar y 75 % al terminar. Tareas de hasta 2 meses con costo concentrado al final.",
    "20/80": "20 % al comenzar y 80 % al terminar. Tareas de hasta 2 meses con costo concentrado al final.",
    "Hitos ponderados": "EV = BAC × suma de los pesos de los hitos cumplidos / suma de todos los pesos. Para "
                        "paquetes de 3 o más meses, con al menos un hito por mes.",
    "% Completado": "EV = BAC × % completado de MS Project (que mide % de duración), ponderado por el costo "
                    "de las subtareas.",
    "% Físico completado": "EV = BAC × '% físico completado' de MS Project (avance medido objetivamente, "
                           "p. ej. m³ ejecutados, planos emitidos).",
    "Unidades completadas": "EV = BAC × unidades terminadas / unidades totales. El cociente se registra en "
                            "'% físico completado' de MS Project.",
    "Avance lineal (LOE)": "Nivel de esfuerzo: EV = PV (gestión, soporte). No genera variación de plazo; "
                           "conviene limitarlo.",
    "CPTR de MS Project": "EV = valor acumulado (CPTR / BCWP) que calcula MS Project a la fecha de estado, según "
                          "el 'Método del valor acumulado' de cada tarea (% completado o % físico completado).",
}


def row_key(row: ReportRow) -> str:
    """Clave estable de una fila de ingreso: UniqueID de la tarea en MS Project."""
    return f"{row.task.uid}-D" if row.kind == KIND_RESIDUAL else str(row.task.uid)


def money_format(decimals: int) -> str:
    return "#,##0" if decimals <= 0 else "#,##0." + "0" * decimals


def col(c: int) -> str:
    return get_column_letter(c)


def clean_text(value):
    """Quita caracteres de control que Excel no admite."""
    return ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value


def put_text(ws, row: int, column: int, value):
    """Escribe un texto libre (nombre de tarea, código...) siempre como texto, nunca como fórmula."""
    cell = ws.cell(row=row, column=column, value=clean_text(value))
    if isinstance(cell.value, str) and (cell.value.startswith("=") or cell.value in ERROR_CODES):
        cell.data_type = "s"
    return cell


def _texts_in_last_row(ws) -> None:
    for cell in ws[ws.max_row]:
        if isinstance(cell.value, str):
            cell.value = clean_text(cell.value)
            if cell.value.startswith("=") or cell.value in ERROR_CODES:
                cell.data_type = "s"


MAX_FORMULA = 8000      # Excel admite fórmulas de hasta 8.192 caracteres


def _sum_formula(letter: str, rows: list[int], block: tuple[int, int] | None = None) -> str:
    """=SUM de las filas indicadas (las consecutivas se agrupan en rangos).

    Si la fórmula resultara demasiado larga, suma las filas de ingreso del bloque
    `block` con SUMIFS sobre el tipo de fila de la hoja _estructura.
    """
    if not rows:
        return "=0"
    rows = sorted(rows)
    refs: list[str] = []
    start = prev = rows[0]
    for r in rows[1:] + [None]:
        if r is not None and r == prev + 1:
            prev = r
            continue
        refs.append(f"{letter}{start}" if start == prev else f"{letter}{start}:{letter}{prev}")
        if r is not None:
            start = prev = r
    chunks = ["SUM(" + ",".join(refs[i:i + 250]) + ")" for i in range(0, len(refs), 250)]
    formula = "=" + (chunks[0] if len(chunks) == 1 else "SUM(" + ",".join(chunks) + ")")
    if len(formula) > MAX_FORMULA and block is not None:
        a, b = block
        off = FIRST_ROW - 2   # fila de datos r <-> fila r - off de _estructura
        formula = (f"=SUMIFS({letter}{a}:{letter}{b},'{SHEET_STRUCT}'!$E${a - off}:$E${b - off},"
                   f"\"<>{KIND_SUMMARY}\")")
    return formula


def _header(ws, row: int, labels: list[tuple[int, str]], height: float = 30) -> None:
    for c, text in labels:
        cell = ws.cell(row=row, column=c, value=text)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = height


def _naive(d):
    if isinstance(d, datetime):
        return d.replace(tzinfo=None)
    return d


# ----------------------------------------------------------------------------
# Escritura
# ----------------------------------------------------------------------------
class _Layout:
    def __init__(self, report: Report):
        self.months = report.horizon.months
        self.first_col = FIRST_MONTH_COL
        self.last_col = FIRST_MONTH_COL + self.months - 1
        self.total_col = self.last_col + 1
        self.row_of: dict[int, int] = {}
        for i, r in enumerate(report.rows):
            self.row_of[id(r)] = FIRST_ROW + i
        self.last_row = FIRST_ROW + len(report.rows) - 1
        self.total_row = self.last_row + 1
        self.block_end: dict[int, int] = {}

        def size(row) -> int:
            n = 1 + sum(size(c) for c in row.children)
            self.block_end[id(row)] = self.row_of[id(row)] + n - 1
            return n

        for root in report.roots:
            size(root)
        self.month_dates = [report.horizon.month_start(i) for i in range(self.months)]

    def r(self, row: ReportRow) -> int:
        return self.row_of[id(row)]

    def block(self, row: ReportRow) -> tuple[int, int]:
        """Filas de los descendientes de una fila resumen."""
        return self.row_of[id(row)] + 1, self.block_end[id(row)]

    def month_range(self, sheet: str, row: int) -> str:
        return f"'{sheet}'!${col(self.first_col)}{row}:${col(self.last_col)}{row}"

    def header_range(self, sheet: str) -> str:
        return f"'{sheet}'!${col(self.first_col)}${HEADER_ROW}:${col(self.last_col)}${HEADER_ROW}"


def _info_line(report: Report, cfg: Config, generated: datetime) -> str:
    p = report.project
    cur = cfg.currency or p.currency_symbol or "-"
    plan = f"Planificación: {report.plan}  |  " if report.plan else ""
    return (f"Proyecto: {p.name}  |  {plan}Origen: {Path(p.source).name}  |  {p.baseline_label}  |  "
            f"Moneda: {cur}  |  Generado: {generated:%d/%m/%Y %H:%M}")


def _write_data_sheet(ws, kind: str, report: Report, cfg: Config, lay: _Layout,
                      generated: datetime, prior=None) -> None:
    titles = {
        SHEET_PV: "Valor Planificado (PV) por mes",
        SHEET_EV: "Valor Ganado (EV) por mes  -  en las celdas amarillas se debe ingresar el valor DEL MES (no acumulado)",
        SHEET_AC: "Costo Real (AC) por mes  -  en las celdas amarillas se debe ingresar el valor DEL MES (no acumulado)",
    }
    money = money_format(cfg.decimals)
    ws.cell(row=TITLE_ROW, column=1, value=titles[kind]).font = FONT_TITLE
    put_text(ws, INFO_ROW, 1, _info_line(report, cfg, generated)).font = FONT_NOTE

    # Fila 3: número de mes; fila 4: encabezados y fecha de cada mes
    ws.cell(row=MONTHNUM_ROW, column=COL_FINISH, value="Mes N°").font = FONT_BOLD
    ws.cell(row=MONTHNUM_ROW, column=COL_FINISH).alignment = Alignment(horizontal="right")
    for i in range(lay.months):
        c = ws.cell(row=MONTHNUM_ROW, column=lay.first_col + i, value=i + 1)
        c.alignment = Alignment(horizontal="center")
        c.font = FONT_BOLD
    _header(ws, HEADER_ROW, [
        (COL_CODE, "Código EDT"), (COL_NAME, "Actividad"), (COL_METHOD, "Método de medición"),
        (COL_BAC, "BAC (Presupuesto)"), (COL_START, "Fecha inicio Plan (PV)"),
        (COL_FINISH, "Fecha fin Plan (PV)"), (lay.total_col, "Total horizonte"),
    ])
    for i, d in enumerate(lay.month_dates):
        c = ws.cell(row=HEADER_ROW, column=lay.first_col + i, value=d)
        c.number_format = FMT_MONTH
        c.fill = FILL_HEADER
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
    if kind == SHEET_PV and report.pv_before > 0.005:
        ws.cell(row=HEADER_ROW, column=lay.first_col).comment = Comment(
            f"El mes 1 incluye {report.pv_before:,.2f} de PV planificado antes del inicio del horizonte.",
            "evm_export")

    for row in report.rows:
        r = lay.r(row)
        put_text(ws, r, COL_CODE, row.code)
        name = put_text(ws, r, COL_NAME, row.name)
        name.alignment = Alignment(indent=max(0, row.level - 1))
        if kind == SHEET_PV:
            method = (prior.method_for(row_key(row)) if prior is not None and row.is_input else None) or row.method
            put_text(ws, r, COL_METHOD, method if row.is_input else None)
            if row.is_input:
                ws.cell(row=r, column=COL_BAC, value=round(row.bac, 2))
            else:
                ws.cell(row=r, column=COL_BAC,
                        value=_sum_formula(col(COL_BAC), [lay.r(c) for c in row.children], lay.block(row)))
            ws.cell(row=r, column=COL_START, value=_naive(row.start))
            ws.cell(row=r, column=COL_FINISH, value=_naive(row.finish))
        else:
            for c_ in (COL_METHOD, COL_BAC, COL_START, COL_FINISH):
                ref = f"'{SHEET_PV}'!{col(c_)}{r}"
                ws.cell(row=r, column=c_, value=f'=IF({ref}="","",{ref})')
        ws.cell(row=r, column=COL_BAC).number_format = money
        ws.cell(row=r, column=COL_START).number_format = FMT_DATE
        ws.cell(row=r, column=COL_FINISH).number_format = FMT_DATE

        # Columnas mensuales
        if row.is_input:
            if kind == SHEET_PV:
                values = round_preserving_total(row.vector, 2)
                for i, v in enumerate(values):
                    if v:
                        ws.cell(row=r, column=lay.first_col + i, value=v)
            elif prior is not None:
                for k, v in prior.values_for(row_key(row), "EV" if kind == SHEET_EV else "AC").items():
                    i = report.horizon.index(k)
                    if 0 <= i < lay.months:
                        ws.cell(row=r, column=lay.first_col + i, value=round(v, 2))
            if kind != SHEET_PV:
                for i in range(lay.months):
                    ws.cell(row=r, column=lay.first_col + i).fill = FILL_INPUT
        else:
            refs_rows = [lay.r(c) for c in row.children]
            for i in range(lay.months):
                L = col(lay.first_col + i)
                ws.cell(row=r, column=lay.first_col + i, value=_sum_formula(L, refs_rows, lay.block(row)))
        for i in range(lay.months):
            ws.cell(row=r, column=lay.first_col + i).number_format = money
        tot = ws.cell(row=r, column=lay.total_col,
                      value=f"=SUM({col(lay.first_col)}{r}:{col(lay.last_col)}{r})")
        tot.number_format = money

        # Estilo por nivel
        font = FONT_BOLD if row.kind == KIND_SUMMARY else (FONT_ITALIC if row.kind == KIND_RESIDUAL else None)
        fill = FILL_L1 if (row.kind == KIND_SUMMARY and row.level == 1) else (
            FILL_L2 if row.kind == KIND_SUMMARY else None)
        for c_ in range(1, lay.total_col + 1):
            cell = ws.cell(row=r, column=c_)
            if font:
                cell.font = font
            if fill and not (row.is_input and kind != SHEET_PV and c_ >= lay.first_col):
                cell.fill = fill
        ws.row_dimensions[r].outline_level = min(7, max(0, row.level - 1))

    # Fila de total
    tr = lay.total_row
    ws.cell(row=tr, column=COL_CODE, value="TOTAL")
    ws.cell(row=tr, column=COL_NAME, value="Total proyecto")
    root_rows = [lay.r(x) for x in report.roots]
    for c_ in [COL_BAC] + list(range(lay.first_col, lay.last_col + 1)):
        L = col(c_)
        ws.cell(row=tr, column=c_, value=_sum_formula(L, root_rows, (FIRST_ROW, lay.last_row))).number_format = money
    ws.cell(row=tr, column=lay.total_col,
            value=f"=SUM({col(lay.first_col)}{tr}:{col(lay.last_col)}{tr})").number_format = money
    for c_ in range(1, lay.total_col + 1):
        cell = ws.cell(row=tr, column=c_)
        cell.fill = FILL_TOTAL
        cell.font = FONT_BOLD
        cell.border = Border(top=Side(style="medium", color=DARK))

    if kind == SHEET_PV:
        # Resaltar filas cuyo total en el horizonte no coincide con el BAC
        rng = f"{col(lay.total_col)}{FIRST_ROW}:{col(lay.total_col)}{lay.total_row}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"ABS({col(lay.total_col)}{FIRST_ROW}-$D{FIRST_ROW})>0.5"], fill=FILL_RED))
        if report.rows:
            dv = DataValidation(type="list", formula1=NAME_METHODS,
                                allow_blank=True, showErrorMessage=True, errorStyle="warning",
                                errorTitle="Método de medición",
                                error="El valor no está en la lista de la hoja 'Listas'. ¿Mantener el valor?")
            ws.add_data_validation(dv)
            for row in report.rows:
                if row.is_input:
                    dv.add(f"{col(COL_METHOD)}{lay.r(row)}")

    widths = {COL_CODE: 12, COL_NAME: 48, COL_METHOD: 20, COL_BAC: 16, COL_START: 12, COL_FINISH: 12}
    for c_, w in widths.items():
        ws.column_dimensions[col(c_)].width = w
    for i in range(lay.months):
        ws.column_dimensions[col(lay.first_col + i)].width = 12
    ws.column_dimensions[col(lay.total_col)].width = 15
    ws.freeze_panes = ws.cell(row=FIRST_ROW, column=lay.first_col)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
    ws.sheet_format.outlineLevelRow = min(7, max([r.level - 1 for r in report.rows] + [0]))


def _method_list(cfg: Config) -> list[str]:
    return cfg.methods or ["Por definir"]


def _write_lists(ws, methods: list[str]) -> None:
    _header(ws, 1, [(1, "Métodos de medición"), (2, "Descripción / regla de cálculo")])
    for i, m in enumerate(methods, start=2):
        put_text(ws, i, 1, m)
        ws.cell(row=i, column=2, value=METHOD_HELP.get(m, ""))
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110
    ws.column_dimensions["D"].width = 70
    ws.cell(row=1, column=4, value="Se pueden agregar métodos al final de la columna A, sin dejar filas vacías: "
            "aparecen en el desplegable de la columna C de la hoja PV. Para calcular el EV se reconocen: "
            "fórmulas fijas X/Y (0/100, 50/50...), hitos ponderados, % completado, % físico completado, "
            "unidades completadas, avance lineal (LOE) y CPTR de MS Project. Otros textos se calculan "
            "como % completado.").font = FONT_NOTE
    ws.cell(row=1, column=4).alignment = Alignment(wrap_text=True, vertical="top")


def _write_curve(ws, report: Report, cfg: Config, lay: _Layout) -> None:
    money = money_format(cfg.decimals)
    ws.cell(row=1, column=1, value="Curva S  -  valores acumulados del proyecto").font = FONT_TITLE
    ws.cell(row=2, column=1, value="Tabla y gráfico calculados con fórmulas a partir de la fila TOTAL de las hojas "
            "PV, EV y AC. EV y AC se muestran desde el primer mes con datos cargados hasta la fecha de corte de la "
            "hoja 'Resumen'.").font = FONT_NOTE
    _header(ws, CURVE_HEADER_ROW, [
        (C_NUM, "Mes N°"), (C_DATE, "Mes"), (C_PV, "PV del mes"), (C_PVC, "PV acumulado"),
        (C_EV, "EV del mes"), (C_EVC, "EV acumulado"), (C_AC, "AC del mes"), (C_ACC, "AC acumulado"),
        (C_PVPCT, "% PV acum."), (C_EVPCT, "% EV acum."), (C_SPI, "SPI acum."), (C_CPI, "CPI acum."),
        (C_BAC, "BAC"), (C_EVN, "Celdas EV cargadas (acum.)"), (C_ACN, "Celdas AC cargadas (acum.)"),
    ])
    tr = lay.total_row
    f0 = CURVE_FIRST_ROW
    bac = f"'{SHEET_PV}'!$D${tr}"
    kinds = f"'{SHEET_STRUCT}'!$E$2:$E${len(report.rows) + 1}"
    for i in range(lay.months):
        r = f0 + i
        mc = col(lay.first_col + i)
        month_cells = f"{mc}${FIRST_ROW}:{mc}${lay.last_row}"
        prev_evn = f"+{col(C_EVN)}{r - 1}" if i else ""
        prev_acn = f"+{col(C_ACN)}{r - 1}" if i else ""
        # EV y AC acumulados: desde el primer mes con datos cargados hasta la fecha de corte
        cut_ev = f'OR({NAME_STATUS}="",{col(C_DATE)}{r}>{NAME_STATUS},{col(C_EVN)}{r}=0)'
        cut_ac = f'OR({NAME_STATUS}="",{col(C_DATE)}{r}>{NAME_STATUS},{col(C_ACN)}{r}=0)'
        cells = {
            C_NUM: i + 1,
            C_DATE: lay.month_dates[i],
            C_PV: f"='{SHEET_PV}'!{mc}{tr}",
            C_PVC: f"=SUM(${col(C_PV)}${f0}:{col(C_PV)}{r})",
            C_EV: f"='{SHEET_EV}'!{mc}{tr}",
            C_EVC: f"=IF({cut_ev},NA(),SUM(${col(C_EV)}${f0}:{col(C_EV)}{r}))",
            C_AC: f"='{SHEET_AC}'!{mc}{tr}",
            C_ACC: f"=IF({cut_ac},NA(),SUM(${col(C_AC)}${f0}:{col(C_AC)}{r}))",
            C_PVPCT: f"=IFERROR({col(C_PVC)}{r}/{bac},0)",
            C_EVPCT: f"=IFERROR({col(C_EVC)}{r}/{bac},NA())",
            C_SPI: f"=IFERROR({col(C_EVC)}{r}/{col(C_PVC)}{r},NA())",
            C_CPI: f"=IFERROR({col(C_EVC)}{r}/{col(C_ACC)}{r},NA())",
            C_BAC: f"={bac}",
            C_EVN: f"=COUNTIFS({kinds},\"<>{KIND_SUMMARY}\",'{SHEET_EV}'!{month_cells},\"<>\"){prev_evn}",
            C_ACN: f"=COUNTIFS({kinds},\"<>{KIND_SUMMARY}\",'{SHEET_AC}'!{month_cells},\"<>\"){prev_acn}",
        }
        for c_, v in cells.items():
            cell = ws.cell(row=r, column=c_, value=v)
            if c_ == C_DATE:
                cell.number_format = FMT_MONTH
            elif c_ in (C_PVPCT, C_EVPCT):
                cell.number_format = FMT_PCT
            elif c_ in (C_SPI, C_CPI):
                cell.number_format = FMT_RATIO
            elif c_ in (C_EVN, C_ACN):
                cell.number_format = "0"
            elif c_ != C_NUM:
                cell.number_format = money
    for c_ in (C_EVN, C_ACN):
        ws.column_dimensions[col(c_)].hidden = True
    for c_, w in {C_NUM: 7, C_DATE: 10}.items():
        ws.column_dimensions[col(c_)].width = w
    for c_ in range(C_PV, C_BAC + 1):
        ws.column_dimensions[col(c_)].width = 14
    ws.freeze_panes = ws.cell(row=CURVE_FIRST_ROW, column=C_PV)
    add_s_curve_chart(ws, lay.months, report.project.name, cfg.currency or report.project.currency_symbol)


def add_s_curve_chart(ws, months: int, project_name: str, currency: str = "") -> None:
    """Agrega (o vuelve a agregar) el gráfico de curva S en la hoja 'Curva S'.

    El gráfico toma sus datos de la tabla de la hoja, calculada con fórmulas de
    Excel, y abarca todo el horizonte: cualquier valor cargado en las hojas PV,
    EV o AC y cualquier cambio de la fecha de corte se reflejan al recalcular.
    """
    ws._charts = []  # openpyxl no conserva gráficos al reabrir: se recrea siempre
    last = CURVE_FIRST_ROW + max(1, months) - 1
    chart = LineChart()
    chart.title = f"Curva S - {project_name}" if project_name else "Curva S"
    chart.height = 15
    chart.width = 30
    chart.y_axis.title = f"Monto acumulado ({currency})" if currency else "Monto acumulado"
    chart.y_axis.number_format = "#,##0"
    chart.y_axis.crossAx = 500
    chart.x_axis = DateAxis(crossAx=100)
    chart.x_axis.number_format = "mmm-yy"
    chart.x_axis.majorTimeUnit = "months"
    chart.x_axis.title = "Mes"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"
    chart.display_blanks = "gap"
    series_spec = [(C_PVC, "1F4E79", None), (C_EVC, "70AD47", None), (C_ACC, "C00000", None),
                   (C_BAC, "7F7F7F", "dash")]
    for c_, color, dash in series_spec:
        ref = Reference(ws, min_col=c_, min_row=CURVE_HEADER_ROW, max_row=last)
        chart.add_data(ref, titles_from_data=True)
        s = chart.series[-1]
        s.smooth = False
        s.marker.symbol = "none"
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 28575 if c_ != C_BAC else 15875
        if dash:
            s.graphicalProperties.line.dashStyle = dash
    chart.set_categories(Reference(ws, min_col=C_DATE, min_row=CURVE_FIRST_ROW, max_row=last))
    ws.add_chart(chart, CHART_ANCHOR)


def _write_summary(ws, report: Report, cfg: Config, lay: _Layout, generated: datetime,
                   status: date | None) -> None:
    money = money_format(cfg.decimals)
    p = report.project
    tr = lay.total_row
    ws.cell(row=1, column=1, value="Informe de Valor Ganado (EVM)").font = FONT_TITLE
    put_text(ws, 2, 1, p.name).font = Font(bold=True, size=12)

    def section(r: int, text: str) -> None:
        for c_ in (1, 2, 3):
            ws.cell(row=r, column=c_).fill = FILL_SECTION
        ws.cell(row=r, column=1, value=text).font = FONT_BOLD

    section(4, "Parámetros")
    st = ws[STATUS_CELL]
    ws["A5"] = "Fecha de corte (fecha de estado)"
    st.value = status
    st.number_format = FMT_DATE
    st.fill = FILL_INPUT
    st.border = BORDER
    ws["C5"] = "Se debe ingresar el último día del mes informado (p. ej. 31/10/2026). Celda con nombre 'FechaCorte'."
    ws["A6"] = "Mes de corte (N°)"
    ws["B6"] = (f'=IF({NAME_STATUS}="","",(YEAR({NAME_STATUS})-YEAR({NAME_START}))*12'
                f'+MONTH({NAME_STATUS})-MONTH({NAME_START})+1)')
    ws["C6"] = "Mes del horizonte que contiene la fecha de corte."
    ws["A7"] = "Inicio del horizonte (mes 1)"
    ws[START_CELL] = report.horizon.start
    ws[START_CELL].number_format = FMT_MONTH
    ws["A8"] = "Meses del horizonte"
    ws["B8"] = report.horizon.months
    ws["A9"] = "Moneda"
    put_text(ws, 9, 2, cfg.currency or p.currency_symbol or "-")
    ws["A10"] = "Línea base"
    put_text(ws, 10, 2, p.baseline_label)
    ws["A11"] = "Origen"
    put_text(ws, 11, 2, str(p.source))
    ws["A12"] = "Fuente del PV mensual"
    put_text(ws, 12, 2, report.pv_source)
    ws["A13"] = "Generado"
    ws["B13"] = generated.replace(microsecond=0)
    ws["B13"].number_format = "dd/mm/yyyy hh:mm"
    ws["A14"] = "Planificación"
    put_text(ws, 14, 2, report.plan or "Todas (sin filtro)")

    pv_tot = lay.month_range(SHEET_PV, tr)
    ev_tot = lay.month_range(SHEET_EV, tr)
    ac_tot = lay.month_range(SHEET_AC, tr)
    hdr = lay.header_range(SHEET_PV)

    def cum(rng: str) -> str:
        return f'=IF({NAME_STATUS}="","",SUMIFS({rng},{hdr},"<="&{NAME_STATUS}))'

    section(15, "Indicadores a la fecha de corte")
    rows = [
        (16, "BAC - Presupuesto a la conclusión", f"='{SHEET_PV}'!D{tr}", money, "Suma del presupuesto de línea base."),
        (17, "PV - Valor planificado acumulado", cum(pv_tot), money, "Trabajo planificado hasta la fecha de corte."),
        (18, "EV - Valor ganado acumulado", cum(ev_tot), money, "Valor del trabajo realizado."),
        (19, "AC - Costo real acumulado", cum(ac_tot), money, "Costo incurrido."),
        (20, "SV = EV - PV", '=IFERROR(B18-B17,"")', money, "< 0: atraso."),
        (21, "CV = EV - AC", '=IFERROR(B18-B19,"")', money, "< 0: sobrecosto."),
        (22, "SPI = EV / PV", '=IFERROR(B18/B17,"")', FMT_RATIO, "< 1: atraso."),
        (23, "CPI = EV / AC", '=IFERROR(B18/B19,"")', FMT_RATIO, "< 1: sobrecosto."),
        (24, "EAC = BAC / CPI", '=IFERROR(B16/B23,"")', money, "Estimación a la conclusión si el desempeño de costo continúa."),
        (25, "EAC = AC + (BAC - EV)", '=IFERROR(B19+B16-B18,"")', money, "Si el trabajo restante se ejecuta según presupuesto."),
        (26, "EAC = AC + (BAC - EV) / (CPI × SPI)", '=IFERROR(B19+(B16-B18)/(B23*B22),"")', money,
         "Si costo y plazo siguen influyendo."),
        (27, "ETC = EAC - AC", '=IFERROR(B24-B19,"")', money, "Costo estimado para terminar (EAC = BAC / CPI)."),
        (28, "VAC = BAC - EAC", '=IFERROR(B16-B24,"")', money, "< 0: se espera terminar por encima del presupuesto."),
        (29, "TCPI = (BAC - EV) / (BAC - AC)", '=IFERROR((B16-B18)/(B16-B19),"")', FMT_RATIO,
         "Eficiencia de costo necesaria para terminar dentro del BAC (> 1: exigente)."),
        (30, "% Planificado = PV / BAC", '=IFERROR(B17/B16,"")', FMT_PCT, ""),
        (31, "% Completado = EV / BAC", '=IFERROR(B18/B16,"")', FMT_PCT, ""),
        (32, "% Gastado = AC / BAC", '=IFERROR(B19/B16,"")', FMT_PCT, ""),
    ]
    curve_pvc = (f"'{SHEET_CURVE}'!${col(C_PVC)}${CURVE_FIRST_ROW}:"
                 f"${col(C_PVC)}${CURVE_FIRST_ROW + lay.months - 1}")
    rows += [
        (34, "Duración planificada (meses)",
         f'=IFERROR(IF(INDEX({curve_pvc},{lay.months})<B16-0.005,"",COUNTIF({curve_pvc},"<"&(B16-0.005))+1),"")',
         "0.0", "Meses del horizonte hasta completar el BAC."),
        (35, "ES - Cronograma ganado (meses)",
         '=IFERROR(IF(B18="","",IF(B18>=B16-0.005,B34,IF(B41>B44,MIN(MAX(B36,IF(B18>0,B44+1,0)),B41),'
         'B41+(B18-B42)/(B43-B42)))),"")', "0.00",
         "Momento en que el plan preveía el EV actual."),
        (36, "AT - Tiempo transcurrido (meses)", '=B6', "0", "Meses desde el inicio del horizonte hasta el corte."),
        (37, "SV(t) = ES - AT (meses)", '=IFERROR(B35-B36,"")', "0.00", "< 0: meses de atraso."),
        (38, "SPI(t) = ES / AT", '=IFERROR(B35/B36,"")', FMT_RATIO, "No converge a 1 al final, a diferencia del SPI."),
        (39, "Duración estimada = duración plan / SPI(t)", '=IFERROR(B34/B38,"")', "0.0", "Meses."),
        (40, "Término estimado (aprox.)", f'=IFERROR(EOMONTH({NAME_START},ROUNDUP(B39,0)-1),"")', FMT_MONTH, ""),
        (41, "Auxiliar ES: meses con PV acumulado <= EV", f'=IFERROR(COUNTIF({curve_pvc},"<="&(B18+0.005)),"")',
         "0", ""),
        (42, "Auxiliar ES: PV acumulado en C", f'=IFERROR(IF(B41=0,0,INDEX({curve_pvc},B41)),"")', money, ""),
        (43, "Auxiliar ES: PV acumulado en C+1", f'=IFERROR(INDEX({curve_pvc},B41+1),"")', money, ""),
        (44, "Auxiliar ES: meses con PV acumulado < EV", f'=IFERROR(COUNTIF({curve_pvc},"<"&(B18-0.005)),"")',
         "0", ""),
    ]
    section(33, "Cronograma ganado (Earned Schedule)")
    for r, label, formula, fmt, note in rows:
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt
        c.border = BORDER
        if note:
            ws.cell(row=r, column=3, value=note).font = FONT_NOTE
    for r in (41, 42, 43, 44):
        ws.cell(row=r, column=1).font = FONT_NOTE

    for ref in ("B22", "B23", "B38"):
        _traffic_light(ws, ref, cfg)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 70
    ws.cell(row=46, column=1, value="Valores mensuales: hojas PV, EV y AC (celdas amarillas = ingreso). "
            "Curva S: hoja 'Curva S'. Indicadores por elemento EDT: hoja 'Indicadores EDT'.").font = FONT_NOTE


def _traffic_light(ws, rng: str, cfg: Config) -> None:
    first = rng.split(":")[0].replace("$", "")
    red, yellow = cfg.threshold_red, cfg.threshold_yellow
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}<{red})'], fill=FILL_RED, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first}),{first}<{yellow})'], fill=FILL_YELLOW, stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'ISNUMBER({first})'], fill=FILL_GREEN, stopIfTrue=True))


def _write_indicators(ws, report: Report, cfg: Config, lay: _Layout) -> None:
    money = money_format(cfg.decimals)
    ws.cell(row=TITLE_ROW, column=1, value="Indicadores por elemento EDT a la fecha de corte").font = FONT_TITLE
    ws.cell(row=INFO_ROW, column=1, value="Fecha de corte: celda 'FechaCorte' de la hoja Resumen. "
            "Valores acumulados hasta el mes de la fecha de corte inclusive.").font = FONT_NOTE
    heads = ["Código EDT", "Actividad", "Método de medición", "BAC", "PV acum.", "EV acum.", "AC acum.",
             "SV", "CV", "SPI", "CPI", "EAC", "VAC", "% Planificado", "% Completado", "% Gastado"]
    _header(ws, HEADER_ROW, list(enumerate(heads, start=1)))
    hdr = lay.header_range(SHEET_PV)
    all_rows = [(lay.r(x), x) for x in report.rows] + [(lay.total_row, None)]
    for r, row in all_rows:
        def cum(sheet: str) -> str:
            return f'=IF({NAME_STATUS}="","",SUMIFS({lay.month_range(sheet, r)},{hdr},"<="&{NAME_STATUS}))'
        f = {
            1: f"='{SHEET_PV}'!A{r}", 2: f"='{SHEET_PV}'!B{r}",
            3: f"=IF('{SHEET_PV}'!C{r}=\"\",\"\",'{SHEET_PV}'!C{r})", 4: f"='{SHEET_PV}'!D{r}",
            5: cum(SHEET_PV), 6: cum(SHEET_EV), 7: cum(SHEET_AC),
            8: f'=IFERROR(F{r}-E{r},"")', 9: f'=IFERROR(F{r}-G{r},"")',
            10: f'=IFERROR(F{r}/E{r},"")', 11: f'=IFERROR(F{r}/G{r},"")',
            12: f'=IFERROR(IF(OR(K{r}="",K{r}=0),G{r}+D{r}-F{r},D{r}/K{r}),"")', 13: f'=IFERROR(D{r}-L{r},"")',
            14: f'=IFERROR(E{r}/D{r},"")', 15: f'=IFERROR(F{r}/D{r},"")', 16: f'=IFERROR(G{r}/D{r},"")',
        }
        for c_, v in f.items():
            cell = ws.cell(row=r, column=c_, value=v)
            if c_ in (4, 5, 6, 7, 8, 9, 12, 13):
                cell.number_format = money
            elif c_ in (10, 11):
                cell.number_format = FMT_RATIO
            elif c_ >= 14:
                cell.number_format = FMT_PCT
        if row is None:
            for c_ in range(1, 17):
                ws.cell(row=r, column=c_).font = FONT_BOLD
                ws.cell(row=r, column=c_).fill = FILL_TOTAL
        else:
            ws.cell(row=r, column=2).alignment = Alignment(indent=max(0, row.level - 1))
            if row.kind == KIND_SUMMARY:
                for c_ in range(1, 17):
                    ws.cell(row=r, column=c_).font = FONT_BOLD
                    ws.cell(row=r, column=c_).fill = FILL_L1 if row.level == 1 else FILL_L2
            ws.row_dimensions[r].outline_level = min(7, max(0, row.level - 1))
    last = lay.total_row
    _traffic_light(ws, f"J{FIRST_ROW}:K{last}", cfg)
    for c_, w in {1: 12, 2: 46, 3: 20}.items():
        ws.column_dimensions[col(c_)].width = w
    for c_ in range(4, 17):
        ws.column_dimensions[col(c_)].width = 14
    ws.freeze_panes = ws.cell(row=FIRST_ROW, column=3)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
    ws.sheet_format.outlineLevelRow = min(7, max([r.level - 1 for r in report.rows] + [0]))


def _write_validation(ws, report: Report) -> None:
    ws.cell(row=1, column=1, value="Validación de la exportación").font = FONT_TITLE
    _header(ws, 3, [(1, "Severidad"), (2, "Código EDT"), (3, "Tarea"), (4, "Detalle")])
    if not report.issues:
        ws.cell(row=4, column=1, value="OK")
        ws.cell(row=4, column=4, value="No se encontraron observaciones.")
    for i, issue in enumerate(report.issues, start=4):
        ws.cell(row=i, column=1, value=issue.severity)
        put_text(ws, i, 2, issue.code)
        put_text(ws, i, 3, issue.name)
        put_text(ws, i, 4, issue.message).alignment = Alignment(wrap_text=True, vertical="top")
        if issue.severity != "Info":
            ws.cell(row=i, column=1).fill = FILL_YELLOW if issue.severity == "Advertencia" else FILL_RED
    for c_, w in {1: 13, 2: 14, 3: 40, 4: 110}.items():
        ws.column_dimensions[col(c_)].width = w
    ws.auto_filter.ref = f"A3:D{max(4, 3 + len(report.issues))}"


def _write_structure(ws, report: Report, lay: _Layout) -> None:
    ws.append(["fila", "clave", "edt", "nivel", "tipo", "fila_padre", "actividad", "metodo_origen"])
    for row in report.rows:
        ws.append([lay.r(row), row_key(row) if row.is_input else str(row.task.uid), row.code, row.level,
                   row.kind, lay.r(row.parent) if row.parent else None, row.name,
                   row.method if row.is_input else None])
        _texts_in_last_row(ws)


def write_meta(ws, meta: dict) -> None:
    ws.delete_rows(1, ws.max_row)
    ws.append(["clave", "valor"])
    for k, v in meta.items():
        ws.append([k, v])
        _texts_in_last_row(ws)


def write_workbook(report: Report, cfg: Config, path: str | Path, prior=None,
                   generated: datetime | None = None) -> dict:
    """Crea el libro Excel completo y lo guarda en `path`. Devuelve los metadatos."""
    generated = generated or datetime.now()
    lay = _Layout(report)
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = SHEET_SUMMARY
    ws_pv = wb.create_sheet(SHEET_PV)
    ws_ev = wb.create_sheet(SHEET_EV)
    ws_ac = wb.create_sheet(SHEET_AC)
    ws_curve = wb.create_sheet(SHEET_CURVE)
    ws_ind = wb.create_sheet(SHEET_IND)
    ws_val = wb.create_sheet(SHEET_VALID)
    ws_lists = wb.create_sheet(SHEET_LISTS)
    ws_struct = wb.create_sheet(SHEET_STRUCT)
    ws_meta = wb.create_sheet(SHEET_META)
    ws_struct.sheet_state = "hidden"
    ws_meta.sheet_state = "hidden"

    status = prior.status_date if prior else None
    methods = list(_method_list(cfg))
    if prior is not None:
        methods += [m for m in prior.methods_list if m not in methods]
    for r in report.rows:
        m = (prior.method_for(row_key(r)) if prior is not None and r.is_input else None) or r.method
        if r.is_input and m and m not in methods:
            methods.append(m)

    _write_lists(ws_lists, methods)
    _write_data_sheet(ws_pv, SHEET_PV, report, cfg, lay, generated, prior)
    _write_data_sheet(ws_ev, SHEET_EV, report, cfg, lay, generated, prior)
    _write_data_sheet(ws_ac, SHEET_AC, report, cfg, lay, generated, prior)
    _write_curve(ws_curve, report, cfg, lay)
    _write_summary(ws_sum, report, cfg, lay, generated, status)
    _write_indicators(ws_ind, report, cfg, lay)
    _write_validation(ws_val, report)
    _write_structure(ws_struct, report, lay)

    wb.defined_names[NAME_STATUS] = DefinedName(NAME_STATUS, attr_text=f"'{SHEET_SUMMARY}'!$B$5")
    wb.defined_names[NAME_START] = DefinedName(NAME_START, attr_text=f"'{SHEET_SUMMARY}'!$B$7")
    # Lista de métodos con tamaño dinámico (crece si se agregan métodos en la hoja Listas)
    wb.defined_names[NAME_METHODS] = DefinedName(
        NAME_METHODS, attr_text=f"OFFSET('{SHEET_LISTS}'!$A$2,0,0,MAX(1,COUNTA('{SHEET_LISTS}'!$A:$A)-1),1)")

    p = report.project
    meta = {
        "version": __version__,
        "proyecto": p.name,
        "origen": str(p.source),
        "motor": p.backend,
        "linea_base": p.baseline_label,
        "fuente_pv": report.pv_source,
        "planificacion": report.plan,
        "moneda": cfg.currency or p.currency_symbol or "",
        "horizonte_inicio": report.horizon.start,
        "meses": report.horizon.months,
        "primera_fila": FIRST_ROW,
        "fila_total": lay.total_row,
        "nivel_maximo": cfg.max_level,
        "decimales": cfg.decimals,
        "generado": generated.replace(microsecond=0),
    }
    write_meta(ws_meta, meta)

    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.properties.creator = "evm_export"
    wb.properties.title = f"EVM - {p.name}"
    save_workbook(wb, path)
    return meta


def save_workbook(wb, path: str | Path) -> None:
    path = Path(path)
    try:
        wb.save(path)
    except PermissionError as exc:
        raise PermissionError(
            f"No se pudo guardar '{path}'. Si el libro está abierto en Excel, se debe cerrar y volver a "
            "intentar.") from exc


# ----------------------------------------------------------------------------
# Lectura (para los comandos actualizar / curva / migración)
# ----------------------------------------------------------------------------
class WorkbookFormatError(ValueError):
    pass


@dataclass
class StructRow:
    row: int
    key: str
    code: str
    level: int
    kind: str
    parent_row: int | None
    name: str
    origin_method: str | None = None    # método que traía la fila al generarse (None en libros antiguos)

    @property
    def is_input(self) -> bool:
        return self.kind in (KIND_LEAF, KIND_RESIDUAL)


def read_meta(wb) -> dict:
    if SHEET_META not in wb.sheetnames or SHEET_STRUCT not in wb.sheetnames:
        raise WorkbookFormatError(
            "El libro no fue generado por este programa (faltan las hojas ocultas _meta/_estructura).")
    meta = {}
    for k, v in wb[SHEET_META].iter_rows(min_row=2, values_only=True):
        if k:
            meta[str(k)] = v
    for k in ("horizonte_inicio", "meses", "primera_fila", "fila_total"):
        if k not in meta:
            raise WorkbookFormatError(f"Falta el dato '{k}' en la hoja _meta.")
    hs = meta["horizonte_inicio"]
    meta["horizonte_inicio"] = hs.date() if isinstance(hs, datetime) else hs
    meta["meses"] = int(meta["meses"])
    meta["primera_fila"] = int(meta["primera_fila"])
    meta["fila_total"] = int(meta["fila_total"])
    return meta


def read_structure(wb) -> list[StructRow]:
    out = []
    for vals in wb[SHEET_STRUCT].iter_rows(min_row=2, values_only=True):
        if not vals or vals[0] is None:
            continue
        fila, key, code, level, kind, parent, name, origin = (list(vals) + [None] * 8)[:8]
        out.append(StructRow(int(fila), str(key), str(code or ""), int(level or 0), str(kind),
                             int(parent) if parent else None, str(name or ""),
                             None if len(vals) < 8 else str(origin or "")))
    return out
