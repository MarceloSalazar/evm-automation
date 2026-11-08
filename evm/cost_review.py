"""Depuración de costes: PV planificado de una empresa frente a su previsión de costes.

Mientras dura el ajuste del cronograma, la empresa escribe o pega en la hoja PV-depuración su plan de costes en tres
niveles: EDT cliente, EDT interna y componente de coste, con el BAC, el EAC y un importe por mes. Los importes se
suman hacia arriba (cálculos que hace el usuario, o procesar-costes con --calculos: calc_plan): los componentes en su
EDT interna, las EDT interna en su EDT cliente y cada EDT cliente en la de su código superior. El EAC de cada línea es
la suma de sus meses y el BAC de la EDT cliente se lee de la hoja PV de la empresa.

Solo se leen las líneas de EDT cliente (código numérico en la columna A, sin nada en C ni en D) cuyo código, sin los
ceros finales (1.2 = 1.2.0), es el de una fila de ingreso del libro. Con fórmulas de Excel, la hoja PV de la empresa
muestra debajo de cada fila su fila «(coste)», con los meses y el EAC de esa línea; la hoja 'Depuración costes' compara
las dos curvas acumuladas de un elemento EDT. El comando procesar-costes (evm/cost_plan.py) lee las mismas líneas
(read_lines) y lleva su importe de cada mes al % planificado. Al terminar, el libro de entrega se genera sin depuración
(con --migrar-desde, que conserva el BAC y el % planificado del plan de costes).

Estado de cada línea de EDT cliente:
- leída: su código es el de una fila de ingreso del libro;
- resumen: su código es el de una fila resumen del libro, que suma sus filas; si ninguna línea de un código inferior
  se lee, su importe queda sin leer;
- incluida: sin fila del libro, pero una línea de un código superior se lee (el plan tiene más detalle que el libro);
- sin fila del libro: su importe queda sin leer.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from openpyxl.chart import LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from evm.config import Config
from evm.fields import fmt_number
from evm.messages import Issue
from evm.wbs import KIND_RESIDUAL, Report, plural
from evm.workbook import (BORDER, COLOR_COST, FILL_COST, FILL_HEADER, FILL_INPUT, FILL_RED, FILL_TOTAL, FMT_MONTH,
                          FMT_PCT, FONT_BOLD, FONT_HEADER, FONT_NOTE, FONT_TITLE, HEADER_ROW, SHEET_PV, _header,
                          _Layout, col, company_sheet, money_format, money_format_hide_zero, put_text)

SHEET_PASTE = "PV-depuración"           # previsión pegada por la empresa
SHEET_CHART = "Depuración costes"       # gráfico comparativo
SHEET_HELPER = "_depuracion"            # hoja oculta con los cálculos de asignación

# Hoja PV-depuración
P_TOTAL_ROW, P_ASSIGNED_ROW, P_MISSING_ROW = 3, 4, 5
P_MONTHNUM_ROW, P_HEADER_ROW, P_FIRST = 6, 7, 8
P_LINES = 2000
P_LAST = P_FIRST + P_LINES - 1
P_CODE, P_DESC, P_INT_CODE, P_INT_DESC, P_BAC, P_EAC, P_FIRST_MONTH = 1, 2, 3, 4, 5, 6, 7
PASTE_MONTHS = 120                      # meses del plan de costes de la empresa (al menos los del horizonte)
EAC_HEADER = "EAC"
# Formatos de la hoja en libros anteriores, que --migrar-desde pasa al actual: sin la columna EAC (meses desde F,
# versión 2.4) y sin las columnas de EDT interna (código, descripción, BAC con este encabezado y meses desde D)
FORMAT_CURRENT, FORMAT_NO_EAC, FORMAT_LEGACY = "actual", "sin EAC", "sin EDT interna"
LEGACY_BAC_HEADER = "BAC (total)"
FIRST_MONTH_OF = {FORMAT_CURRENT: P_FIRST_MONTH, FORMAT_NO_EAC: 6, FORMAT_LEGACY: 4}

# Hoja _depuracion: columnas por línea pegada (filas P_FIRST a P_LAST, alineadas con PV-depuración)
(H_TEXT, H_NUM, H_NORM, H_CLIENT, H_TOTAL, H_HORIZON, H_MATCH, H_ROW, H_ABOVE, H_BELOW, H_REPEAT,
 H_STATE) = range(1, 13)
# Tabla de códigos del libro (desde la fila 2): código sin ceros finales, fila si es de ingreso (-1 si es un resumen)
# y fila de la hoja (para leer su BAC)
L_KEY, L_EXACT, L_ROW = 14, 15, 16
NO_ROW = -1
# Lista del selector de la hoja 'Depuración costes' (desde la fila 2) y fila elegida
S_CODE, S_ROW = 18, 19
SELECTED = "$U$1"
# Indicador de alineamiento: diferencia acumulada (previsión - PV planificado) por mes, en la fila de cada elemento
CUM_FIRST = 23

# Estado de cada línea (H_STATE): no es de EDT cliente, leída, resumen con líneas inferiores leídas, resumen sin
# ellas, incluida en una línea superior leída y sin fila del libro
STATE_OTHER, STATE_READ, STATE_SUMMARY, STATE_SUMMARY_EMPTY, STATE_INCLUDED, STATE_MISSING = range(6)
STATUS_SUMMARY = "Resumen: suma de sus filas"
STATUS_SUMMARY_EMPTY = "Resumen sin detalle en el plan"
STATUS_INCLUDED = "Incluida en su EDT cliente superior"
STATUS_NONE = "Sin fila del libro"
STATUS_OF = {STATE_SUMMARY: STATUS_SUMMARY, STATE_SUMMARY_EMPTY: STATUS_SUMMARY_EMPTY,
             STATE_INCLUDED: STATUS_INCLUDED, STATE_MISSING: STATUS_NONE}
UNREAD = (STATE_SUMMARY_EMPTY, STATE_MISSING)       # su importe no llega a ninguna fila del libro

# Hoja 'Depuración costes'
G_SELECT = "B4"
G_HEADER_ROW, G_FIRST = 11, 12
(G_NUM, G_DATE, G_PV, G_PVC, G_COST, G_COSTC, G_DIFF) = range(1, 8)
FILL_OUTSIDE = PatternFill("solid", fgColor="7F7F7F")     # meses posteriores al horizonte del libro
FILL_ALERT = PatternFill("solid", fgColor="F28B82", bgColor="F28B82")   # se distingue del naranja de «(coste)»
FILL_CALC = PatternFill("solid", fgColor="FFBF00")        # PV-depuración: cálculos del plan (naranja)


def paste_months(months: int) -> int:
    """Columnas de meses de la hoja PV-depuración: 120 o los meses del horizonte, si son más."""
    return max(PASTE_MONTHS, months)


def review_sheet(report: Report) -> str:
    """Hoja que se depura: 'PV' o 'PV-EmpresaA'."""
    if report.cost_review == "A":
        return company_sheet(SHEET_PV, report.company_a)
    if report.cost_review == "B":
        return company_sheet(SHEET_PV, report.company_b)
    return SHEET_PV


def review_company(report: Report) -> str:
    """Empresa que se depura ('' en un libro de una sola empresa)."""
    return {"A": report.company_a, "B": report.company_b}.get(report.cost_review or "", "")


def norm_code(code) -> str | None:
    """Código EDT numérico sin los ceros finales ('1.2.0' -> '1.2'); None si no es numérico (las mismas reglas que
    las fórmulas de la hoja _depuracion)."""
    text = str(code if code is not None else "").strip().replace(",", ".")
    parts = text.split(".")
    if not text or not all(p and p.isascii() and p.isdigit() for p in parts):
        return None
    while len(parts) > 1 and parts[-1] == "0":
        parts.pop()
    return ".".join(parts)


def code_table(rows) -> tuple[dict[str, tuple[object, object]], list, list[str]]:
    """Códigos del libro que pueden recibir el plan de costes: {código sin ceros finales: (fila si es de ingreso o
    NO_ROW si es un resumen, fila)}. rows: (código, tipo, ¿de ingreso?, fila) de cada fila de la EDT. Devuelve también
    las filas con código no numérico y los códigos repetidos. Las filas «Otras tareas» no tienen código propio."""
    entries: dict[str, tuple[object, object]] = {}
    nonnumeric = []
    repeated: list[str] = []
    for code, kind, is_input, row in rows:
        if kind == KIND_RESIDUAL:
            continue
        key = norm_code(code)
        if key is None:
            nonnumeric.append(row)
            continue
        if key in entries:
            repeated.append(code)
            continue
        entries[key] = (row if is_input else NO_ROW, row)
    return entries, nonnumeric, repeated


def lookup_rows(report: Report) -> tuple[dict[str, tuple[object, object]], list, list[str]]:
    """code_table con las filas de la EDT del informe (ReportRow)."""
    return code_table((r.code, r.kind, r.is_input, r) for r in report.rows)


def structure_table(structure) -> dict[str, tuple[int, int]]:
    """code_table con las filas de la hoja _estructura de un libro (StructRow): números de fila de las hojas."""
    return code_table((s.code, s.kind, s.is_input, s.row) for s in structure)[0]


def is_amount(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _cell_text(value) -> str:
    """Texto de una celda como en las fórmulas (A&"", con coma decimal cambiada por punto y sin espacios)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).replace(",", ".").strip()


def _cells(values: list) -> tuple[str, str, str]:
    """Columnas A (EDT cliente), C (EDT interna) y D (actividad interna o componente) de una línea, como texto."""
    v = list(values) + [None] * P_INT_DESC
    return _cell_text(v[P_CODE - 1]), _cell_text(v[P_INT_CODE - 1]), _cell_text(v[P_INT_DESC - 1])


# ----------------------------------------------------------------------------
# Lectura en Python (procesar-costes): las mismas reglas que las fórmulas de la hoja _depuracion
# ----------------------------------------------------------------------------
@dataclass
class PasteLine:
    row: int                            # fila de la hoja PV-depuración
    values: list                        # columnas A a la del último mes
    key: str | None = None              # EDT cliente sin ceros finales (H_NORM); None si la línea no es de EDT cliente
    match: int = 0                      # fila de ingreso del libro; NO_ROW: resumen; 0: ninguna (H_MATCH)
    book_row: int = 0                   # fila del libro con el código, de ingreso o resumen (H_ROW)
    state: int = STATE_OTHER            # H_STATE
    repeated: bool = False              # otra línea de EDT cliente con el mismo código (H_REPEAT)

    @property
    def client(self) -> bool:
        return self.key is not None

    @property
    def months(self) -> list:
        return self.values[P_FIRST_MONTH - 1:]

    @property
    def total(self) -> float:
        return sum(v for v in self.months if is_amount(v))


def read_lines(rows: list[list], table: dict[str, tuple[object, object]]) -> list[PasteLine]:
    """Líneas de la hoja PV-depuración (columnas A a la del último mes, desde la fila P_FIRST) con la fila del libro
    de las de EDT cliente y su estado, con las mismas reglas que las fórmulas de la hoja _depuracion. table:
    code_table del libro."""
    lines = [PasteLine(P_FIRST + k, list(values)) for k, values in enumerate(rows)]
    for x in lines:
        a, c, d = _cells(x.values)
        key = norm_code(a) if a else None
        if key is not None and not c and not d:
            x.key = key
            x.match, x.book_row = table.get(key, (0, 0))
    clients = [x for x in lines if x.client]
    read = {x.key for x in clients if x.match > 0}
    counts: dict[str, int] = {}
    for x in clients:
        counts[x.key] = counts.get(x.key, 0) + 1
    for x in clients:
        parts = x.key.split(".")
        x.repeated = counts[x.key] > 1
        if x.match > 0:
            x.state = STATE_READ
        elif x.match == NO_ROW:
            below = any(k.startswith(x.key + ".") for k in read)
            x.state = STATE_SUMMARY if below else STATE_SUMMARY_EMPTY
        else:
            above = any(".".join(parts[:k]) in read for k in range(1, len(parts)))
            x.state = STATE_INCLUDED if above else STATE_MISSING
    return lines


# ----------------------------------------------------------------------------
# Cálculos del plan de costes (procesar-costes --calculos)
# ----------------------------------------------------------------------------
KIND_CLIENT, KIND_INTERNAL, KIND_COMPONENT = "cliente", "interna", "componente"


@dataclass
class PlanLine:
    kind: str                           # KIND_CLIENT, KIND_INTERNAL, KIND_COMPONENT o "" (vacía o solo descripción)
    key: str | None = None              # EDT cliente sin ceros finales
    parent: int | None = None           # índice de la línea de la que depende
    children: list[int] = field(default_factory=list)


def plan_lines(rows: list[list]) -> list[PlanLine]:
    """Nivel de cada línea y la línea de la que depende: el componente de coste (solo la columna D), de la última EDT
    interna de arriba, de la misma EDT cliente, o de la EDT cliente si no tiene EDT interna; la EDT interna (código en
    C, o no numérico en A), de la última EDT cliente de arriba; la EDT cliente, de la del código superior más largo
    que esté en el plan (1.2.1 de 1.2 o de 1)."""
    out: list[PlanLine] = []
    client = internal = None
    first_of: dict[str, int] = {}
    for k, values in enumerate(rows):
        a, c, d = _cells(values)
        key = norm_code(a) if a else None
        if key is not None and not c and not d:
            out.append(PlanLine(KIND_CLIENT, key))
            first_of.setdefault(key, k)
            client, internal = k, None
        elif c or (a and key is None):
            out.append(PlanLine(KIND_INTERNAL, parent=client))
            internal = k
        elif d:
            out.append(PlanLine(KIND_COMPONENT, parent=internal if internal is not None else client))
        else:
            out.append(PlanLine(""))
    for x in out:
        if x.kind == KIND_CLIENT:
            parts = x.key.split(".")
            x.parent = next((first_of[".".join(parts[:j])] for j in range(len(parts) - 1, 0, -1)
                             if ".".join(parts[:j]) in first_of), None)
    for k, x in enumerate(out):
        if x.parent is not None:
            out[x.parent].children.append(k)
    return out


def calc_plan(rows: list[list]) -> tuple[list[list], list[PlanLine]]:
    """Cálculos del plan: los meses (y el BAC de la EDT interna) de cada línea con otras debajo son la suma de esas
    líneas, y el EAC de cada línea es la suma de sus meses. Devuelve las líneas con los importes calculados y el nivel
    de cada una (para escribir las fórmulas)."""
    plan = plan_lines(rows)
    width = max([len(v) for v in rows] + [P_FIRST_MONTH])
    out = [list(values) + [None] * (width - len(values)) for values in rows]
    done: set[int] = set()

    def amount(k: int, c: int) -> float:
        v = out[k][c]
        return v if is_amount(v) else 0.0

    def compute(k: int) -> None:
        if k in done:
            return
        done.add(k)
        x = plan[k]
        if not x.children:
            return
        for j in x.children:
            compute(j)
        columns = list(range(P_FIRST_MONTH - 1, width))
        if x.kind == KIND_INTERNAL:
            columns.insert(0, P_BAC - 1)
        for c in columns:
            out[k][c] = sum(amount(j, c) for j in x.children)

    for k in range(len(plan)):
        compute(k)
    for k, x in enumerate(plan):
        if x.kind:
            out[k][P_EAC - 1] = sum(amount(k, c) for c in range(P_FIRST_MONTH - 1, width))
    return out, plan


def _refs(letter: str, rows: list[int]) -> str:
    """Celdas de las filas indicadas, con las consecutivas en un rango: 'G12:G15,G17'."""
    refs: list[str] = []
    start = prev = None
    for r in sorted(rows) + [None]:
        if r is not None and prev is not None and r == prev + 1:
            prev = r
            continue
        if start is not None:
            refs.append(f"{letter}{start}" if start == prev else f"{letter}{start}:{letter}{prev}")
        start = prev = r
    return ",".join(refs)


def write_calcs(ws, plan: list[PlanLine], months: int, sheet: str, total_row: int, decimals: int) -> int:
    """Escribe en la hoja PV-depuración las fórmulas de los cálculos del plan (celdas naranjas): EAC = suma de los
    meses; meses (y BAC de la EDT interna) de cada línea con otras debajo = suma de esas líneas; BAC de la EDT cliente
    = el de su fila de la hoja PV de la empresa (sheet). Las celdas de ingreso quedan en amarillo. Devuelve la cantidad
    de líneas calculadas."""
    n_paste = paste_months(months)
    first, last = col(P_FIRST_MONTH), col(P_FIRST_MONTH + n_paste - 1)
    book_row = f"'{SHEET_HELPER}'!{col(H_ROW)}"
    money, month_money = money_format(decimals), money_format_hide_zero(decimals)
    done = 0
    for k, x in enumerate(plan):
        if not x.kind:
            continue
        done += 1
        r = P_FIRST + k
        rows = [P_FIRST + j for j in x.children]
        cells: dict[int, tuple[str | None, PatternFill]] = {P_EAC: (f"=SUM({first}{r}:{last}{r})", FILL_CALC)}
        if x.kind == KIND_CLIENT:
            cells[P_BAC] = (f"=IF({book_row}{r}>0,INDEX('{sheet}'!$D$1:$D${total_row + 1},{book_row}{r}),\"\")",
                            FILL_CALC)
        elif rows and x.kind == KIND_INTERNAL:
            cells[P_BAC] = (f"=SUM({_refs(col(P_BAC), rows)})", FILL_CALC)
        else:
            cells[P_BAC] = (None, FILL_INPUT)
        for i in range(n_paste):
            c_ = P_FIRST_MONTH + i
            cells[c_] = (f"=SUM({_refs(col(c_), rows)})", FILL_CALC) if rows else (None, FILL_INPUT)
        for c_, (formula, fill) in cells.items():
            cell = ws.cell(row=r, column=c_)
            if formula is not None:
                cell.value = formula
            cell.fill = fill
            cell.number_format = month_money if c_ >= P_FIRST_MONTH else money
    return done


def review_issues(report: Report) -> list[Issue]:
    """Observaciones de la hoja Validación de un libro con depuración de costes."""
    if report.cost_review is None:
        return []
    company = review_company(report)
    sheet = review_sheet(report)
    issues = [Issue("DEP-001", "", "",
                    f"Depuración de costes{' de ' + company if company else ''}: se debe pegar la previsión mensual "
                    f"en la hoja {SHEET_PASTE} (desde la celda A{P_FIRST}); las filas «(coste)» de la hoja {sheet} y "
                    f"la hoja {SHEET_CHART} la comparan con el PV planificado. El libro de entrega se debe generar sin "
                    "depuracion_costes, con --migrar-desde para conservar el BAC.")]
    _, nonnumeric, repeated = lookup_rows(report)
    if nonnumeric:
        codes = ", ".join(r.code or "(sin código)" for r in nonnumeric[:10])
        more = f" y {len(nonnumeric) - 10} más" if len(nonnumeric) > 10 else ""
        issues.append(Issue("DEP-002", "", "",
                            f"Depuración de costes: {plural(len(nonnumeric), 'fila tiene', 'filas tienen')} código EDT "
                            f"no numérico ({codes}{more}) y no pueden recibir la previsión de la hoja {SHEET_PASTE}."))
    if repeated:
        issues.append(Issue("DEP-003", "", "",
                            "Depuración de costes: códigos EDT repetidos en el libro (" + ", ".join(repeated[:10]) +
                            f"); la previsión de la hoja {SHEET_PASTE} se asigna a la primera fila de cada código."))
    return issues


def _read_sum(column: int, r: int) -> str:
    """Suma de una columna de la hoja PV-depuración en las líneas de EDT cliente de la fila de ingreso r."""
    L = col(column)
    return (f"=SUMIFS('{SHEET_PASTE}'!${L}${P_FIRST}:${L}${P_LAST},"
            f"'{SHEET_HELPER}'!${col(H_MATCH)}${P_FIRST}:${col(H_MATCH)}${P_LAST},{r})")


def paste_sum(i: int, r: int) -> str:
    """Fórmula del mes i (base 0) de la fila «(coste)» de la fila de ingreso r: el importe del mes de su línea de EDT
    cliente en la hoja PV-depuración."""
    return _read_sum(P_FIRST_MONTH + i, r)


def paste_eac(r: int) -> str:
    """Fórmula de la columna D de la fila «(coste)» de la fila de ingreso r: el EAC de su línea de EDT cliente."""
    return _read_sum(P_EAC, r)


# ----------------------------------------------------------------------------
# Indicador de alineamiento (depuracion_alineamiento)
# ----------------------------------------------------------------------------
def _threshold(cfg: Config) -> float:
    return cfg.cost_review_threshold / 100


def alignment_cells(ws, cfg: Config, lay: _Layout, r: int, c: int) -> None:
    """Columnas del indicador en la fila «(coste)» c: diferencia entre el EAC y el BAC y diferencia máxima acumulada
    entre las dos curvas, en % del BAC de la fila r."""
    first, last = col(CUM_FIRST), col(CUM_FIRST + lay.months - 1)
    cum = f"'{SHEET_HELPER}'!${first}{r}:${last}{r}"
    values = {lay.total_col + 1: f'=IF(N($D{r})=0,"",($D{c}-$D{r})/$D{r})',
              lay.total_col + 2: f'=IF(N($D{r})=0,"",MAX(MAX({cum}),-MIN({cum}))/$D{r})'}
    for c_, v in values.items():
        cell = ws.cell(row=c, column=c_, value=v)
        cell.number_format = FMT_PCT
        cell.fill = FILL_COST
        cell.font = Font(italic=True, color=COLOR_COST)


def alignment_format(ws, report: Report, cfg: Config, lay: _Layout) -> None:
    """Encabezados del indicador y resaltado: meses cuya diferencia acumulada supera el umbral y diferencias de BAC o
    acumuladas que lo superan."""
    th = cfg.cost_review_threshold
    _header(ws, HEADER_ROW, [(lay.total_col + 1, "Dif. EAC (% del BAC)"),
                             (lay.total_col + 2, f"Dif. máx. acumulada (umbral {fmt_number(th, trim=True)} %)")])
    ws.row_dimensions[HEADER_ROW].height = 30
    for c_ in (lay.total_col + 1, lay.total_col + 2):
        ws.column_dimensions[col(c_)].width = 14
    rows = lay.cost_rows(report)
    if not rows:
        return
    t = _threshold(cfg)
    c0 = rows[0]
    first, last = col(lay.first_col), col(lay.last_col)
    months = " ".join(f"{first}{c}:{last}{c}" for c in rows)
    ws.conditional_formatting.add(months, FormulaRule(
        formula=[f"AND(N($D{c0 - 1})>0,ABS(SUM(${first}{c0}:{first}{c0})-SUM(${first}{c0 - 1}:{first}{c0 - 1}))"
                 f">{t!r}*$D{c0 - 1})"], fill=FILL_ALERT))
    x = col(lay.total_col + 1)
    y = col(lay.total_col + 2)
    ws.conditional_formatting.add(" ".join(f"{x}{c}:{y}{c}" for c in rows), FormulaRule(
        formula=[f"AND(ISNUMBER({x}{c0}),ABS({x}{c0})>{t!r})"], fill=FILL_ALERT))


# ----------------------------------------------------------------------------
# Hojas
# ----------------------------------------------------------------------------
def _digits_removed(ref: str) -> str:
    out = ref
    for ch in "0123456789.":
        out = f'SUBSTITUTE({out},"{ch}","")'
    return out


def _prefix_part(ref: str, k: int) -> str:
    """Los k primeros segmentos del código ref ('1.2.3', k = 2 -> '1.2'); error si tiene k segmentos o menos."""
    return f'LEFT({ref},FIND("|",SUBSTITUTE({ref},".","|",{k}))-1)'


def write_helper(ws, report: Report, lay: _Layout, cfg: Config) -> None:
    """Hoja oculta _depuracion: fila del libro y estado de cada línea pegada, tabla de códigos del libro, lista del
    selector y diferencias acumuladas del indicador."""
    rs = review_sheet(report)
    n_paste = paste_months(lay.months)
    last_paste = col(P_FIRST_MONTH + n_paste - 1)
    last_horizon = col(P_FIRST_MONTH + lay.months - 1)
    paste = f"'{SHEET_PASTE}'"
    heads = {H_TEXT: "codigo", H_NUM: "numerico", H_NORM: "codigo_normalizado", H_CLIENT: "edt_cliente",
             H_TOTAL: "total_linea", H_HORIZON: "total_horizonte", H_MATCH: "fila_ingreso", H_ROW: "fila_libro",
             H_ABOVE: "superior_leida", H_BELOW: "inferior_leida", H_REPEAT: "repetida", H_STATE: "estado"}
    for c_, text in heads.items():
        ws.cell(row=P_FIRST - 1, column=c_, value=text)
    tables = {L_KEY: "edt", L_EXACT: "fila_ingreso", L_ROW: "fila", S_CODE: "selector", S_ROW: "fila",
              CUM_FIRST - 1: "fila_libro"}
    for c_, text in tables.items():
        ws.cell(row=1, column=c_, value=text)

    entries, _nonnumeric, _repeated = lookup_rows(report)

    def row_number(value) -> int:
        return value if value == NO_ROW else lay.r(value)

    for i, (key, (exact, row)) in enumerate(entries.items(), start=2):
        put_text(ws, i, L_KEY, key)
        ws.cell(row=i, column=L_EXACT, value=row_number(exact))
        ws.cell(row=i, column=L_ROW, value=row_number(row))
    n_keys = max(2, len(entries) + 1)
    keys = f"${col(L_KEY)}$2:${col(L_KEY)}${n_keys}"
    exact_rng = f"${col(L_EXACT)}$2:${col(L_EXACT)}${n_keys}"
    row_rng = f"${col(L_ROW)}$2:${col(L_ROW)}${n_keys}"
    depth = max([key.count(".") + 1 for key in entries] + [1])

    def rng(c_: int) -> str:
        return f"${col(c_)}${P_FIRST}:${col(c_)}${P_LAST}"

    for r in range(P_FIRST, P_LAST + 1):
        a, b, cn, cl, m, above, below = (f"{col(x)}{r}" for x in (H_TEXT, H_NUM, H_NORM, H_CLIENT, H_MATCH, H_ABOVE,
                                                                   H_BELOW))
        # Una línea de código superior leída (comparación de textos: COUNTIFS tomaría '1.10' como 1,1)
        read_above = "+".join(f"IFERROR(SUMPRODUCT(({rng(H_NORM)}={_prefix_part(cn, k)})*({rng(H_MATCH)}>0)),0)"
                              for k in range(1, depth + 1))
        cells = {
            H_TEXT: f'=TRIM(SUBSTITUTE({paste}!A{r}&"",",","."))',
            H_NUM: (f'=IF({a}="",0,IF(AND(LEN({_digits_removed(a)})=0,LEFT({a},1)<>".",RIGHT({a},1)<>".",'
                    f'ISERROR(FIND("..",{a}))),1,0))'),
            H_NORM: (f'=IF({b}=1,LEFT({a},LEN({a})-2*SUMPRODUCT(--(RIGHT({a},{{2,4,6,8,10,12,14,16,18}})'
                     f'=REPT(".0",{{1,2,3,4,5,6,7,8,9}})))),"")'),
            # Línea de EDT cliente: código numérico en A, sin EDT interna ni componente
            H_CLIENT: (f'=IF(AND({b}=1,LEN(TRIM({paste}!{col(P_INT_CODE)}{r}&""))=0,'
                       f'LEN(TRIM({paste}!{col(P_INT_DESC)}{r}&""))=0),1,0)'),
            H_TOTAL: f"=SUM({paste}!{col(P_FIRST_MONTH)}{r}:{last_paste}{r})",
            H_HORIZON: f"=SUM({paste}!{col(P_FIRST_MONTH)}{r}:{last_horizon}{r})",
            H_MATCH: f"=IF({cl}=1,IFERROR(INDEX({exact_rng},MATCH({cn},{keys},0)),0),0)",
            H_ROW: f"=IF({cl}=1,IFERROR(INDEX({row_rng},MATCH({cn},{keys},0)),0),0)",
            H_ABOVE: f"=IF(AND({cl}=1,{m}=0),IF({read_above}>0,1,0),0)",
            H_BELOW: f'=IF({m}={NO_ROW},IF(COUNTIFS({rng(H_NORM)},{cn}&".*",{rng(H_MATCH)},">0")>0,1,0),0)',
            H_REPEAT: f"=IF({cl}=1,IF(SUMPRODUCT(({rng(H_NORM)}={cn})*({rng(H_CLIENT)}=1))>1,1,0),0)",
            H_STATE: (f"=IF({cl}=0,{STATE_OTHER},IF({m}>0,{STATE_READ},IF({m}={NO_ROW},"
                      f"IF({below}=1,{STATE_SUMMARY},{STATE_SUMMARY_EMPTY}),"
                      f"IF({above}=1,{STATE_INCLUDED},{STATE_MISSING}))))"),
        }
        for c_, v in cells.items():
            ws.cell(row=r, column=c_, value=v)

    # Lista del selector: TOTAL y los elementos EDT del libro, con su fila en la hoja que se depura
    put_text(ws, 2, S_CODE, "TOTAL")
    ws.cell(row=2, column=S_ROW, value=lay.total_row)
    for i, row in enumerate(report.rows, start=3):
        put_text(ws, i, S_CODE, row.code)
        ws.cell(row=i, column=S_ROW, value=lay.r(row))
    n_list = len(report.rows) + 2
    ws[SELECTED] = (f"=IFERROR(INDEX(${col(S_ROW)}$2:${col(S_ROW)}${n_list},MATCH('{SHEET_CHART}'!${G_SELECT[0]}$"
                    f"{G_SELECT[1:]},${col(S_CODE)}$2:${col(S_CODE)}${n_list},0)),{lay.total_row})")

    if cfg.cost_review_alignment:
        rows = [lay.r(x) for x in report.rows] + [lay.total_row]
        for r in rows:
            ws.cell(row=r, column=CUM_FIRST - 1, value=r)
            for i in range(lay.months):
                L = col(lay.first_col + i)
                prev = f"{col(CUM_FIRST + i - 1)}{r}+" if i else ""
                ws.cell(row=r, column=CUM_FIRST + i, value=f"={prev}('{rs}'!{L}{r + 1}-'{rs}'!{L}{r})")


def selector_range(report: Report) -> str:
    return f"'{SHEET_HELPER}'!${col(S_CODE)}$2:${col(S_CODE)}${len(report.rows) + 2}"


def write_paste(ws, report: Report, cfg: Config, lay: _Layout, rows: list[list] | None = None,
                calcs: bool = False) -> None:
    """Hoja PV-depuración: el bloque del plan de costes de la empresa (desde A8), los encabezados de los meses y los
    controles de la lectura. rows: líneas trasladadas del libro anterior (--migrar-desde); calcs: la hoja anterior
    tenía fórmulas, que se vuelven a escribir con write_calcs."""
    company = review_company(report)
    rs = review_sheet(report)
    money = money_format(cfg.decimals)
    n_paste = paste_months(lay.months)
    status_col = P_FIRST_MONTH + n_paste + 1
    obs_col = status_col + 1
    helper = f"'{SHEET_HELPER}'"
    amounts, last_paste = col(P_BAC), col(P_FIRST_MONTH + n_paste - 1)      # BAC, EAC y meses de cada línea
    control = col(P_EAC)                # controles: importes de los meses, en la columna del EAC
    ws.cell(row=1, column=1, value="Plan de costes para la depuración del PV" +
            (f"  -  {company}" if company else "")).font = FONT_TITLE
    ws.cell(row=2, column=1, value=(
        f"Pegar o escribir desde la celda A{P_FIRST}: EDT cliente, actividad EDT, EDT interna, actividad interna o "
        f"componente de coste, BAC, EAC y un importe por mes, con el mes 1 en la columna {col(P_FIRST_MONTH)}. Los "
        "importes se suman hacia arriba: los componentes de coste en su EDT interna y las EDT interna en su EDT "
        f"cliente; el EAC es la suma de los meses y el BAC de la EDT cliente, el de la hoja {rs} (celdas naranjas: "
        "cálculos, que hace procesar-costes --calculos; amarillas: datos). Solo se leen las líneas de EDT cliente "
        f"cuyo código es una fila de la hoja {rs}.")).font = FONT_NOTE

    def by_state(column: int, *states: int) -> str:
        return "+".join(f"SUMIFS({helper}!${col(column)}${P_FIRST}:${col(column)}${P_LAST},"
                        f"{helper}!${col(H_STATE)}${P_FIRST}:${col(H_STATE)}${P_LAST},{s})" for s in states)

    labels = [(P_TOTAL_ROW, "Leído en filas del libro (meses del horizonte)", "=" + by_state(H_HORIZON, STATE_READ)),
              (P_ASSIGNED_ROW, "Leído después del horizonte (no forma parte del PV)",
               f"={by_state(H_TOTAL, STATE_READ)}-{control}{P_TOTAL_ROW}"),
              (P_MISSING_ROW, "Sin leer: EDT cliente sin fila del libro o resumen sin detalle",
               "=" + by_state(H_TOTAL, *UNREAD))]
    for r, label, formula in labels:
        ws.cell(row=r, column=P_DESC, value=label).font = FONT_BOLD
        cell = ws.cell(row=r, column=P_EAC, value=formula)
        cell.number_format = money
        cell.border = BORDER
    ws.conditional_formatting.add(f"{control}{P_ASSIGNED_ROW}:{control}{P_MISSING_ROW}", FormulaRule(
        formula=[f"ABS({control}{P_ASSIGNED_ROW})>0.5"], fill=FILL_RED))
    ws.cell(row=P_TOTAL_ROW, column=P_FIRST_MONTH,
            value=f"Resultado: filas «(coste)» de la hoja {rs} y hoja '{SHEET_CHART}'; el comando procesar-costes "
                  "lleva el importe de cada mes al % planificado.").font = FONT_NOTE
    ws.cell(row=P_MISSING_ROW, column=P_FIRST_MONTH,
            value=f"Fila del libro de cada línea: columnas {col(status_col)} ('Asignado a') y {col(obs_col)} "
                  "('Observación'), a la derecha de los meses.").font = FONT_NOTE

    ws.cell(row=P_MONTHNUM_ROW, column=P_EAC, value="Mes N°").font = FONT_BOLD
    ws.cell(row=P_MONTHNUM_ROW, column=P_EAC).alignment = Alignment(horizontal="right")
    _header(ws, P_HEADER_ROW, [(P_CODE, "EDT cliente"), (P_DESC, "Actividad EDT"), (P_INT_CODE, "EDT interna"),
                               (P_INT_DESC, "Actividad interna / componente de coste"), (P_BAC, "BAC"),
                               (P_EAC, EAC_HEADER), (status_col, "Asignado a"), (obs_col, "Observación")])
    for i in range(n_paste):
        c_ = P_FIRST_MONTH + i
        num = ws.cell(row=P_MONTHNUM_ROW, column=c_, value=i + 1)
        num.alignment = Alignment(horizontal="center")
        num.font = FONT_BOLD
        head = ws.cell(row=P_HEADER_ROW, column=c_, value=report.horizon.month_start(i))
        head.number_format = FMT_MONTH
        head.font = FONT_HEADER
        head.alignment = Alignment(horizontal="center", vertical="center")
        head.border = BORDER
        head.fill = FILL_HEADER if i < lay.months else FILL_OUTSIDE
        ws.column_dimensions[col(c_)].width = 11
        ws.column_dimensions[col(c_)].number_format = money
    if n_paste > lay.months:
        ws.cell(row=P_HEADER_ROW, column=P_FIRST_MONTH + lay.months).comment = Comment(
            f"Meses posteriores al horizonte del libro ({lay.months} meses): no forman parte del PV.", "evm_export")
    ws.column_dimensions[col(P_CODE)].number_format = "@"
    ws.column_dimensions[col(P_INT_CODE)].number_format = "@"
    ws.column_dimensions[col(P_BAC)].number_format = money
    ws.column_dimensions[col(P_EAC)].number_format = money

    for k, values in enumerate((rows or [])[:P_LINES]):
        for c_, v in enumerate(values, start=1):
            if v is not None and c_ < status_col:
                cell = put_text(ws, P_FIRST + k, c_, v)
                if c_ >= P_BAC and is_amount(v):
                    cell.number_format = money
    if rows and calcs:
        write_calcs(ws, plan_lines(rows[:P_LINES]), lay.months, rs, lay.total_row, cfg.decimals)

    for r in range(P_FIRST, P_LAST + 1):
        h = {x: f"{helper}!{col(x)}{r}" for x in (H_STATE, H_MATCH, H_REPEAT, H_TOTAL, H_HORIZON)}
        choices = [(STATE_READ, f"INDEX('{rs}'!$A$1:$A${lay.total_row + 1},{h[H_MATCH]})")]
        choices += [(s, f'"{text}"') for s, text in STATUS_OF.items()]
        formula = '""'                  # STATE_OTHER: líneas de EDT interna y componentes, que no se leen
        for s, value in reversed(choices):
            formula = f"IF({h[H_STATE]}={s},{value},{formula})"
        ws.cell(row=r, column=status_col, value=f"={formula}")
        ws.cell(row=r, column=obs_col, value=(
            f'=IF({h[H_STATE]}<>{STATE_READ},"",TRIM(IF({h[H_REPEAT]}=1,"EDT cliente repetida: se suman sus '
            f'líneas. ","")&IF(ABS({h[H_TOTAL]}-{h[H_HORIZON]})>0.5,"Importes fuera del horizonte del libro. ","")'
            f'&IF(SUMPRODUCT(ISTEXT({amounts}{r}:{last_paste}{r})*(LEN({amounts}{r}:{last_paste}{r})>0))>0,'
            f'"Importes guardados como texto: no se suman. ","")))'))
    status = f"${col(status_col)}{P_FIRST}"
    obs = f"${col(obs_col)}{P_FIRST}"
    area = f"{col(status_col)}{P_FIRST}:{col(obs_col)}{P_LAST}"
    unread = ",".join(f'{status}="{STATUS_OF[s]}"' for s in UNREAD)
    ws.conditional_formatting.add(area, FormulaRule(formula=[f'OR({unread},{obs}<>"")'], fill=FILL_RED))
    ws.conditional_formatting.add(f"A{P_FIRST}:{col(P_INT_DESC)}{P_LAST}", FormulaRule(
        formula=[f'OR({status}="{STATUS_SUMMARY}",{status}="{STATUS_INCLUDED}")'], fill=FILL_TOTAL,
        font=Font(italic=True, color="595959")))

    for c_, w in {P_CODE: 12, P_DESC: 34, P_INT_CODE: 12, P_INT_DESC: 34, P_BAC: 14, P_EAC: 14, status_col - 1: 3,
                  status_col: 34, obs_col: 60}.items():
        ws.column_dimensions[col(c_)].width = w
    ws.row_dimensions[P_HEADER_ROW].height = 30
    ws.freeze_panes = ws.cell(row=P_FIRST, column=P_FIRST_MONTH)


def write_chart_sheet(ws, report: Report, cfg: Config, lay: _Layout) -> None:
    """Hoja 'Depuración costes': curvas acumuladas del PV planificado (BAC × %) y de la previsión del elemento EDT
    elegido en la celda B4 (por defecto, TOTAL)."""
    company = review_company(report)
    rs = review_sheet(report)
    money = money_format(cfg.decimals)
    sel = f"'{SHEET_HELPER}'!{SELECTED}"
    ws.cell(row=1, column=1, value="Depuración de costes  -  PV planificado frente a la previsión" +
            (f" de {company}" if company else "")).font = FONT_TITLE
    ws.cell(row=2, column=1, value=(
        f"Se elige el elemento EDT en la celda {G_SELECT} (TOTAL o un código de la hoja {rs}). PV planificado: BAC × "
        f"% planificado. Previsión: fila «(coste)», con los meses y el EAC de la EDT cliente en la hoja {SHEET_PASTE}. "
        "El objetivo es que el EAC coincida con el BAC y que las curvas se parezcan: se ajusta el plan de costes, o en "
        "MS Project las tareas, sus fechas y el peso de los hitos (y se regenera el libro con --migrar-desde); "
        "procesar-costes lleva el plan al % planificado.")).font = FONT_NOTE
    ws.cell(row=2, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)
    ws.row_dimensions[2].height = 45

    ws.cell(row=4, column=1, value="Elemento EDT").font = FONT_BOLD
    pick = ws[G_SELECT]
    pick.value = "TOTAL"
    pick.fill = FILL_INPUT
    pick.border = BORDER
    pick.alignment = Alignment(horizontal="center")
    ws.cell(row=4, column=3, value=f"=IFERROR(INDEX('{rs}'!$B$1:$B${lay.total_row + 1},{sel}),\"\")")
    dv = DataValidation(type="list", formula1=f"={selector_range(report)}", allow_blank=False,
                        showErrorMessage=True, errorTitle="Elemento EDT",
                        error="Se debe elegir TOTAL o un código EDT de la lista.")
    ws.add_data_validation(dv)
    dv.add(G_SELECT)

    end = G_FIRST + lay.months - 1
    diff = f"${col(G_DIFF)}${G_FIRST}:${col(G_DIFF)}${end}"
    rows = [
        (5, "BAC (fila de porcentaje)", f"=N(INDEX('{rs}'!$D$1:$D${lay.total_row + 1},{sel}))", money),
        (6, "EAC (fila «(coste)»)", f"=N(INDEX('{rs}'!$D$1:$D${lay.total_row + 1},{sel}+1))", money),
        (7, "Diferencia: EAC - BAC", "=B6-B5", money),
        (8, "Diferencia máxima acumulada", f"=MAX(MAX({diff}),-MIN({diff}))", money),
    ]
    for r, label, formula, fmt in rows:
        ws.cell(row=r, column=1, value=label)
        cell = ws.cell(row=r, column=2, value=formula)
        cell.number_format = fmt
        cell.border = BORDER
    for r in (7, 8):
        pct = ws.cell(row=r, column=3, value=f'=IF(B5=0,"",B{r}/B5)')
        pct.number_format = FMT_PCT
    ws.cell(row=7, column=4, value="(% del BAC)").font = FONT_NOTE
    t = _threshold(cfg)
    for r in (7, 8):
        ws.conditional_formatting.add(f"C{r}", FormulaRule(formula=[f"AND(ISNUMBER(C{r}),ABS(C{r})>{t!r})"],
                                                           fill=FILL_RED))

    _header(ws, G_HEADER_ROW, [(G_NUM, "Mes N°"), (G_DATE, "Mes"), (G_PV, "PV planificado del mes"),
                               (G_PVC, "PV planificado acumulado"), (G_COST, "Previsión del mes"),
                               (G_COSTC, "Previsión acumulada"), (G_DIFF, "Diferencia acumulada")])
    block = f"'{rs}'!${col(lay.first_col)}$1:${col(lay.last_col)}${lay.total_row + 1}"
    for i in range(lay.months):
        r = G_FIRST + i
        cells = {
            G_NUM: i + 1,
            G_DATE: lay.month_dates[i],
            G_PV: f"=N(INDEX({block},{sel},{i + 1}))",
            G_PVC: f"=SUM(${col(G_PV)}${G_FIRST}:{col(G_PV)}{r})",
            G_COST: f"=N(INDEX({block},{sel}+1,{i + 1}))",
            G_COSTC: f"=SUM(${col(G_COST)}${G_FIRST}:{col(G_COST)}{r})",
            G_DIFF: f"={col(G_COSTC)}{r}-{col(G_PVC)}{r}",
        }
        for c_, v in cells.items():
            cell = ws.cell(row=r, column=c_, value=v)
            cell.number_format = FMT_MONTH if c_ == G_DATE else ("0" if c_ == G_NUM else money)
            if c_ in (G_NUM, G_DATE):
                cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    for c_ in range(G_PV, G_DIFF + 1):
        ws.column_dimensions[col(c_)].width = 15
    ws.freeze_panes = ws.cell(row=G_FIRST, column=G_PV)


# Eje de montos del gráfico: la escala depende del elemento EDT elegido (millones, miles o unidades)
AXIS_FORMAT = '[>=1000000]#,##0.0,,"M";[>=1000]#,##0,"k";#,##0'


def add_chart(wb, months: int, chart_months: int, currency: str) -> None:
    """Agrega (o vuelve a agregar: openpyxl no conserva los gráficos) el gráfico de la hoja 'Depuración costes'."""
    from evm.dashboard import _date_axis, _frame, _legend, _line, _style_axis
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.marker import Marker
    ws = wb[SHEET_CHART]
    ws._charts = []
    n = max(1, min(months, chart_months))
    last = G_FIRST + n - 1
    line = LineChart()
    for c_, color, dash in ((G_PVC, "1F4E79", None), (G_COSTC, "C55A11", "dash")):
        line.add_data(Reference(ws, min_col=c_, min_row=G_HEADER_ROW, max_row=last), titles_from_data=True)
        s = line.series[-1]
        s.smooth = False
        s.marker = Marker(symbol="none")
        s.graphicalProperties = GraphicalProperties(ln=_line(color, 2.25, dash))
    x = _date_axis(n)
    line.x_axis = x
    line.y_axis.crossAx = x.axId
    _style_axis(line.y_axis, title=currency or "Monto", fmt=AXIS_FORMAT, grid=True)
    line.y_axis.scaling.min = 0
    line.set_categories(Reference(ws, min_col=G_DATE, min_row=G_FIRST, max_row=last))
    _legend(line)
    _frame(line, "PV planificado y previsión de costes acumulados", 760, 400)
    ws.add_chart(line, "I4")


def paste_format(ws) -> str:
    """Formato de la hoja PV-depuración: FORMAT_CURRENT, FORMAT_NO_EAC (versión 2.4) o FORMAT_LEGACY (sin EDT
    interna)."""
    if ws.cell(row=P_HEADER_ROW, column=P_INT_CODE).value == LEGACY_BAC_HEADER:
        return FORMAT_LEGACY
    if ws.cell(row=P_HEADER_ROW, column=P_EAC).value == EAC_HEADER:
        return FORMAT_CURRENT
    return FORMAT_NO_EAC


def last_paste_month(ws, months: int) -> int:
    """Último mes (N°, dentro del horizonte) con algún importe pegado en la hoja PV-depuración."""
    first = FIRST_MONTH_OF[paste_format(ws)]
    last = 0
    for values in ws.iter_rows(min_row=P_FIRST, max_row=min(ws.max_row, P_LAST), min_col=first,
                               max_col=first + months - 1, values_only=True):
        for i, v in enumerate(values):
            if is_amount(v) and v and i + 1 > last:
                last = i + 1
    return last


def read_paste(ws, months: int) -> tuple[list[list], str]:
    """Líneas pegadas en la hoja PV-depuración (columnas A a la del último mes), sin las filas vacías del final, y el
    formato de la hoja. Las líneas de un formato anterior se devuelven en el actual: sin las columnas de EDT interna,
    un código numérico y su descripción quedan como EDT cliente y actividad, y un código no numérico (XRF01) y su
    descripción, como EDT interna y actividad interna; sin la columna EAC, se agrega con la suma de los meses."""
    fmt = paste_format(ws)
    n = FIRST_MONTH_OF[fmt] + paste_months(months) - 1
    rows = [list(v) for v in ws.iter_rows(min_row=P_FIRST, max_row=min(ws.max_row, P_LAST), max_col=n,
                                          values_only=True)]
    while rows and all(v is None or v == "" for v in rows[-1]):
        rows.pop()
    if fmt == FORMAT_LEGACY:
        rows = [([code, desc, None, None] if norm_code(code) is not None else [None, None, code, desc]) + rest
                for code, desc, *rest in rows]
    if fmt != FORMAT_CURRENT:
        rows = [values[:P_EAC - 1] + [_months_total(values[P_EAC - 1:])] + values[P_EAC - 1:] for values in rows]
    return rows, fmt


def resolve_formulas(path, rows: list[list], months: int) -> tuple[list[list], int, list[tuple[int, int]]]:
    """Cambia las fórmulas de las líneas de la hoja PV-depuración (read_paste) por el valor que calculó Excel.
    Devuelve las líneas, la cantidad de celdas con fórmula y las celdas (línea, columna, desde 0) sin valor calculado
    (el libro no se guardó en Excel después de escribirlas)."""
    cells = [(k, c) for k, values in enumerate(rows) for c, v in enumerate(values)
             if isinstance(v, str) and v.startswith("=")]
    if not cells:
        return rows, 0, []
    from openpyxl import load_workbook
    wv = load_workbook(path, data_only=True, read_only=True)
    try:
        cached, _fmt = read_paste(wv[SHEET_PASTE], months)
    finally:
        wv.close()
    missing = []
    for k, c in cells:
        value = cached[k][c] if k < len(cached) and c < len(cached[k]) else None
        if value is None:
            missing.append((k, c))
        rows[k][c] = value
    return rows, len(cells), missing


def _months_total(values: list) -> float | None:
    amounts = [v for v in values if is_amount(v)]
    return sum(amounts) if amounts else None


def shift_rows(rows: list[list], offset: int, n_paste: int) -> tuple[list[list], float]:
    """Desplaza los meses de las líneas pegadas cuando cambia el mes 1 del horizonte (offset: meses que avanza el
    inicio). Devuelve las líneas y el importe que queda fuera de las columnas de meses."""
    lost = 0.0
    out = []
    for values in rows:
        head = (list(values[:P_FIRST_MONTH - 1]) + [None] * (P_FIRST_MONTH - 1))[:P_FIRST_MONTH - 1]
        months = list(values[P_FIRST_MONTH - 1:])
        new = [None] * n_paste
        for i, v in enumerate(months):
            j = i - offset
            if 0 <= j < n_paste:
                new[j] = v
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                lost += abs(v)
        out.append(head + new)
    return out, lost


def month_offset(old: date, new: date) -> int:
    return (new.year - old.year) * 12 + (new.month - old.month)
