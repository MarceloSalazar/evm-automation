"""Depuración de costes: PV planificado de una empresa frente a su previsión de costes.

Mientras dura el ajuste del cronograma, la empresa pega en la hoja PV-depuración su plan de costes (EDT cliente,
actividad EDT, EDT interna, actividad interna o componente de coste, BAC, EAC y un importe por mes). Con fórmulas de
Excel, cada línea pegada se asigna a una fila del libro y la hoja PV de la empresa muestra, debajo de cada fila, su
fila «(coste)» con la previsión. La hoja 'Depuración costes' compara las dos curvas acumuladas de un elemento EDT. El
comando procesar-costes (evm/cost_plan.py) lleva el reparto mensual del plan al % planificado de cada fila, con las
mismas reglas (assign_lines). Al terminar, el libro de entrega se genera sin depuración (con --migrar-desde, que
conserva el BAC y el % planificado del plan de costes).

Reglas de asignación de cada línea pegada, en tres niveles:
- EDT cliente (columna A) con código numérico (2.3.4.1): su fila del libro; un código más detallado que el libro
  (2.3.4.1.5) va a la fila de la que depende. Los ceros finales no cuentan (1.2 = 1.2.0) y se admite el código
  convertido en número por Excel;
- EDT interna (código en la columna C, como GPR01, o código no numérico en la columna A, como XRF01): pertenece al
  último código numérico de las líneas anteriores;
- componente de coste (solo la columna D, sin código en C): pertenece a la última EDT interna de arriba, de la misma
  EDT cliente, o a la EDT cliente si no tiene EDT interna;
- una línea de EDT interna o un componente que repite la EDT cliente en la columna A pertenece a esa EDT cliente,
  como si la columna A estuviera vacía;
- una línea con otras que dependen de ella es un subtotal y no se suma: la EDT cliente (sin EDT interna) con EDT
  interna, componentes o códigos más detallados debajo, y la EDT interna con componentes de coste debajo.
"""
from __future__ import annotations

from dataclasses import dataclass
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
                          _Layout, col, company_sheet, money_format, put_text)

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
(H_TEXT, H_NUM, H_NORM, H_TOTAL, H_CHILD, H_GROUP, H_EFF, H_PARENT, H_COUNT, H_TRUNC, H_TARGET, H_HORIZON,
 H_SUBDIFF, H_INT, H_ICODE, H_COMP, H_SUB) = range(1, 18)
# Tabla de códigos del libro (desde la fila 2): código sin ceros finales, fila si el código coincide y fila si la
# línea depende del código. Un resumen del libro sin fila propia lleva -1 (la línea queda sin asignar).
L_KEY, L_EXACT, L_PREFIX = 19, 20, 21
NO_ROW = -1
# Lista del selector de la hoja 'Depuración costes' (desde la fila 2) y fila elegida
S_CODE, S_ROW = 23, 24
SELECTED = "$Z$1"
# Indicador de alineamiento: diferencia acumulada (previsión - PV planificado) por mes, en la fila de cada elemento
CUM_FIRST = 28

# Hoja 'Depuración costes'
G_SELECT = "B4"
G_HEADER_ROW, G_FIRST = 11, 12
(G_NUM, G_DATE, G_PV, G_PVC, G_COST, G_COSTC, G_DIFF) = range(1, 8)
STATUS_SUBTOTAL = "Subtotal (no se suma)"
STATUS_SUMMARY = "Resumen sin fila propia"
STATUS_NONE = "Sin fila del libro"
FILL_OUTSIDE = PatternFill("solid", fgColor="7F7F7F")     # meses posteriores al horizonte del libro
FILL_ALERT = PatternFill("solid", fgColor="F28B82", bgColor="F28B82")   # se distingue del naranja de «(coste)»


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
    """Códigos del libro que pueden recibir la previsión: {código sin ceros finales: (fila exacta, fila de lo que
    depende del código)}, con la fila o NO_ROW. rows: (código, tipo, ¿de ingreso?, fila, fila «Otras tareas» de un
    resumen o None) de cada fila de la EDT. Devuelve también las filas con código no numérico y los códigos
    repetidos."""
    entries: dict[str, tuple[object, object]] = {}
    nonnumeric = []
    repeated: list[str] = []
    for code, kind, is_input, row, residual in rows:
        if kind == KIND_RESIDUAL:
            continue                    # recibe lo que depende de su resumen sin fila propia
        key = norm_code(code)
        if key is None:
            nonnumeric.append(row)
            continue
        value = (row, row) if is_input else (NO_ROW, NO_ROW if residual is None else residual)
        if key in entries:
            repeated.append(code)
            continue
        entries[key] = value
    return entries, nonnumeric, repeated


def lookup_rows(report: Report) -> tuple[dict[str, tuple[object, object]], list, list[str]]:
    """code_table con las filas de la EDT del informe (ReportRow)."""
    return code_table((r.code, r.kind, r.is_input, r, next((c for c in r.children if c.kind == KIND_RESIDUAL), None))
                      for r in report.rows)


def structure_table(structure) -> dict[str, tuple[int, int]]:
    """code_table con las filas de la hoja _estructura de un libro (StructRow): números de fila de las hojas."""
    residual: dict[int, int] = {}
    for s in structure:
        if s.kind == KIND_RESIDUAL and s.parent_row is not None:
            residual.setdefault(s.parent_row, s.row)
    return code_table((s.code, s.kind, s.is_input, s.row, residual.get(s.row)) for s in structure)[0]


# ----------------------------------------------------------------------------
# Asignación en Python (procesar-costes): las mismas reglas que las fórmulas de la hoja _depuracion
# ----------------------------------------------------------------------------
@dataclass
class PasteLine:
    row: int                            # fila de la hoja PV-depuración
    values: list                        # columnas A a la del último mes
    text: str = ""                      # EDT cliente como texto (H_TEXT)
    key: str | None = None              # EDT cliente numérica sin ceros finales (H_NORM); None si no es numérica
    child: bool = False                 # línea con contenido que no es una EDT cliente numérica (H_CHILD)
    internal: bool = False              # algo en la columna C o en la D (H_INT)
    int_code: bool = False              # EDT interna: código en la columna C o no numérico en la A (H_ICODE)
    component: bool = False             # componente de coste (H_COMP)
    group: int = 0                      # fila de la EDT cliente de la que depende (H_GROUP)
    internal_row: int = 0               # fila de la EDT interna de la que depende (H_SUB)
    effective: str = ""                 # código con que se asigna (H_EFF)
    subtotal: bool = False              # H_PARENT
    counted: bool = False               # H_COUNT
    target: int = 0                     # fila del libro; NO_ROW: resumen sin fila propia; 0: ninguna (H_TARGET)

    @property
    def numeric(self) -> bool:
        return self.key is not None

    @property
    def months(self) -> list:
        return self.values[P_FIRST_MONTH - 1:]

    @property
    def total(self) -> float:
        return sum(v for v in self.months if is_amount(v))


def is_amount(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _cell_text(value) -> str:
    """Texto de una celda como en las fórmulas (A&"", con coma decimal cambiada por punto y sin espacios)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).replace(",", ".").strip()


def _prefix_row(key: str, table: dict[str, tuple[object, object]]) -> object:
    """Fila de la que depende un código más detallado que el libro: la del código más largo de la tabla que es un
    prefijo, por segmentos ('2.3.4.1.5' -> '2.3.4.1' -> '2.3.4'...); 0 si no hay (_trunc_chain)."""
    parts = key.split(".")
    for k in range(len(parts) - 1, 0, -1):
        found = table.get(".".join(parts[:k]))
        if found is not None:
            return found[1]
    return 0


def assign_lines(rows: list[list], table: dict[str, tuple[object, object]]) -> list[PasteLine]:
    """Asigna las líneas pegadas (columnas A a la del último mes, desde la fila P_FIRST) a las filas del libro, con
    las mismas reglas que las fórmulas de la hoja _depuracion. table: code_table del libro."""
    lines = [PasteLine(P_FIRST + k, list(values)) for k, values in enumerate(rows)]
    by_row = {x.row: x for x in lines}
    prev: PasteLine | None = None
    for x in lines:
        v = x.values + [None] * (P_FIRST_MONTH - 1 - len(x.values))
        x.text = _cell_text(v[P_CODE - 1])
        x.key = norm_code(x.text) if x.text else None
        int_code = bool(_cell_text(v[P_INT_CODE - 1]))
        x.internal = int_code or bool(_cell_text(v[P_INT_DESC - 1]))
        x.child = not x.numeric and (bool(x.text) or int_code or any(is_amount(a) for a in v[P_BAC - 1:]))
        x.int_code = int_code or (not x.numeric and bool(x.text))
        x.component = not x.int_code and x.internal and (x.child or x.numeric)
        prev_group = prev.group if prev else 0
        if not x.numeric:
            x.group = prev_group
        elif x.internal and prev_group > 0 and by_row[prev_group].key == x.key:
            x.group = prev_group        # EDT interna o componente que repite la EDT cliente de su grupo
        else:
            x.group = x.row
        if x.int_code:
            x.internal_row = x.row
        elif prev is not None and x.group == prev_group:
            x.internal_row = prev.internal_row
        x.effective = x.key if x.numeric else (by_row[x.group].key if x.group > 0 else "")
        prev = x
    for x in lines:
        if x.int_code:
            x.subtotal = any(y.internal_row == x.row and y.component for y in lines)
        elif x.numeric and not x.internal:
            x.subtotal = any(y.group == x.row and (y.child or (y.numeric and y.internal)) for y in lines) or any(
                y.key is not None and y.key.startswith(x.key + ".") for y in lines)
        x.counted = not x.subtotal and (x.numeric or x.child)
        if x.counted and x.effective:
            found = table.get(x.effective)
            if found is None:
                x.target = _prefix_row(x.effective, table)
            else:
                x.target = found[0] if x.numeric and not x.internal else found[1]
    return lines


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


def paste_sum(i: int, r: int) -> str:
    """Fórmula del mes i (base 0) de la fila «(coste)» de la fila de ingreso r: suma de las líneas pegadas asignadas
    a la fila."""
    L = col(P_FIRST_MONTH + i)
    return (f"=SUMIFS('{SHEET_PASTE}'!${L}${P_FIRST}:${L}${P_LAST},"
            f"'{SHEET_HELPER}'!${col(H_TARGET)}${P_FIRST}:${col(H_TARGET)}${P_LAST},{r})")


# ----------------------------------------------------------------------------
# Indicador de alineamiento (depuracion_alineamiento)
# ----------------------------------------------------------------------------
def _threshold(cfg: Config) -> float:
    return cfg.cost_review_threshold / 100


def alignment_cells(ws, cfg: Config, lay: _Layout, r: int, c: int) -> None:
    """Columnas del indicador en la fila «(coste)» c: diferencia de BAC y diferencia máxima acumulada entre las dos
    curvas, en % del BAC de la fila r."""
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
    _header(ws, HEADER_ROW, [(lay.total_col + 1, "Dif. BAC (% del BAC)"),
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


def _trunc_chain(ref: str, depth: int, keys: str, prefix: str) -> str:
    """Fila de la que depende el código ref (sin contar el código completo): el código más largo de la tabla del
    libro que es un prefijo de ref, por segmentos ('2.3.4.1.5' -> '2.3.4.1' -> '2.3.4' -> ...); 0 si no hay."""
    out = "0"
    for k in range(1, depth + 1):
        part = f'LEFT({ref},FIND("|",SUBSTITUTE({ref},".","|",{k}))-1)'
        out = f"IFERROR(INDEX({prefix},MATCH({part},{keys},0)),{out})"
    return out


def write_helper(ws, report: Report, lay: _Layout, cfg: Config) -> None:
    """Hoja oculta _depuracion: asignación de cada línea pegada, tabla de códigos del libro, lista del selector y
    diferencias acumuladas del indicador."""
    rs = review_sheet(report)
    n_paste = paste_months(lay.months)
    last_paste = col(P_FIRST_MONTH + n_paste - 1)
    last_horizon = col(P_FIRST_MONTH + lay.months - 1)
    paste = f"'{SHEET_PASTE}'"
    heads = {H_TEXT: "codigo", H_NUM: "numerico", H_NORM: "codigo_normalizado", H_TOTAL: "total_linea",
             H_CHILD: "subtarea", H_GROUP: "fila_grupo", H_EFF: "codigo_efectivo", H_PARENT: "subtotal",
             H_COUNT: "se_suma", H_TRUNC: "fila_por_dependencia", H_TARGET: "fila_libro", H_HORIZON: "total_horizonte",
             H_SUBDIFF: "dif_subtotal", H_INT: "columna_c_o_d", H_ICODE: "edt_interna", H_COMP: "componente",
             H_SUB: "fila_edt_interna"}
    for c_, text in heads.items():
        ws.cell(row=P_FIRST - 1, column=c_, value=text)
    tables = {L_KEY: "edt", L_EXACT: "fila_exacta", L_PREFIX: "fila_dependencia", S_CODE: "selector", S_ROW: "fila",
              CUM_FIRST - 1: "fila_libro"}
    for c_, text in tables.items():
        ws.cell(row=1, column=c_, value=text)

    entries, _nonnumeric, _repeated = lookup_rows(report)

    def row_number(value) -> int:
        return value if value == NO_ROW else lay.r(value)

    for i, (key, (exact, prefix)) in enumerate(entries.items(), start=2):
        put_text(ws, i, L_KEY, key)
        ws.cell(row=i, column=L_EXACT, value=row_number(exact))
        ws.cell(row=i, column=L_PREFIX, value=row_number(prefix))
    n_keys = max(2, len(entries) + 1)
    keys = f"${col(L_KEY)}$2:${col(L_KEY)}${n_keys}"
    exact_rng = f"${col(L_EXACT)}$2:${col(L_EXACT)}${n_keys}"
    prefix_rng = f"${col(L_PREFIX)}$2:${col(L_PREFIX)}${n_keys}"
    depth = max([key.count(".") + 1 for key in entries] + [1])

    def rng(c_: int) -> str:
        return f"${col(c_)}${P_FIRST}:${col(c_)}${P_LAST}"

    norm = f"${col(H_NORM)}$1:${col(H_NORM)}${P_LAST}"
    for r in range(P_FIRST, P_LAST + 1):
        a, b, cn, d, e, f, g, h, i_, j, k, n, ic = (
            f"{col(x)}{r}" for x in (H_TEXT, H_NUM, H_NORM, H_TOTAL, H_CHILD, H_GROUP, H_EFF, H_PARENT, H_COUNT,
                                     H_TRUNC, H_TARGET, H_INT, H_ICODE))
        prev = f"{col(H_GROUP)}{r - 1}"
        int_code = f'LEN(TRIM({paste}!{col(P_INT_CODE)}{r}&""))>0'
        cells = {
            H_TEXT: f'=TRIM(SUBSTITUTE({paste}!A{r}&"",",","."))',
            H_NUM: (f'=IF({a}="",0,IF(AND(LEN({_digits_removed(a)})=0,LEFT({a},1)<>".",RIGHT({a},1)<>".",'
                    f'ISERROR(FIND("..",{a}))),1,0))'),
            H_NORM: (f'=IF({b}=1,LEFT({a},LEN({a})-2*SUMPRODUCT(--(RIGHT({a},{{2,4,6,8,10,12,14,16,18}})'
                     f'=REPT(".0",{{1,2,3,4,5,6,7,8,9}})))),"")'),
            H_TOTAL: f"=SUM({paste}!{col(P_FIRST_MONTH)}{r}:{last_paste}{r})",
            H_CHILD: (f'=IF(AND({b}=0,OR({a}<>"",{int_code},'
                      f'COUNT({paste}!{col(P_BAC)}{r}:{last_paste}{r})>0)),1,0)'),
            # Grupo: la línea numérica de la que dependen las líneas siguientes. Una línea de EDT interna que repite
            # la EDT cliente de su grupo sigue en ese grupo.
            H_GROUP: (f"=IF({b}=1,ROW(),0)" if r == P_FIRST else
                      f"=IF({b}=1,IF(AND({n}=1,{prev}>0),IF(INDEX({norm},{prev})={cn},{prev},ROW()),ROW()),{prev})"),
            H_EFF: f'=IF({b}=1,{cn},IF({f}>0,INDEX({norm},{f}),""))',
            # Subtotal: la EDT interna con componentes de coste debajo, o la EDT cliente con líneas que dependen de
            # ella (EDT interna, componentes o códigos más detallados)
            H_PARENT: (f'=IF({ic}=1,IF(COUNTIFS({rng(H_SUB)},ROW(),{rng(H_COMP)},1)>0,1,0),'
                       f'IF(AND({b}=1,{n}=0),IF(OR(COUNTIFS({rng(H_GROUP)},ROW(),{rng(H_CHILD)},1)>0,'
                       f'COUNTIFS({rng(H_GROUP)},ROW(),{rng(H_NUM)},1,{rng(H_INT)},1)>0,'
                       f'COUNTIF({rng(H_NORM)},{cn}&".*")>0),1,0),0))'),
            H_COUNT: f"=IF({h}=1,0,IF(OR({b}=1,{e}=1),1,0))",
            H_TRUNC: f'=IF({g}="",0,{_trunc_chain(g, depth, keys, prefix_rng)})',
            H_TARGET: (f'=IF({i_}=0,0,IF({g}="",0,IF(AND({b}=1,{n}=0),'
                       f'IFERROR(INDEX({exact_rng},MATCH({g},{keys},0)),{j}),'
                       f'IFERROR(INDEX({prefix_rng},MATCH({g},{keys},0)),{j}))))'),
            H_HORIZON: f"=SUM({paste}!{col(P_FIRST_MONTH)}{r}:{last_horizon}{r})",
            H_SUBDIFF: (f'=IF({h}=1,{d}-IF({ic}=1,SUMIFS({rng(H_TOTAL)},{rng(H_COUNT)},1,{rng(H_SUB)},ROW()),'
                        f'SUMIFS({rng(H_TOTAL)},{rng(H_COUNT)},1,{rng(H_GROUP)},ROW())'
                        f'+SUMIFS({rng(H_TOTAL)},{rng(H_COUNT)},1,{rng(H_EFF)},{cn}&".*")),0)'),
            H_INT: f'=IF(OR({int_code},LEN(TRIM({paste}!{col(P_INT_DESC)}{r}&""))>0),1,0)',
            H_ICODE: f'=IF(OR({int_code},AND({b}=0,{a}<>"")),1,0)',
            H_COMP: f"=IF(AND({ic}=0,{n}=1,OR({e}=1,{b}=1)),1,0)",
            # Fila de la última EDT interna de arriba, en la misma EDT cliente
            H_SUB: (f"=IF({ic}=1,ROW(),0)" if r == P_FIRST else
                    f"=IF({ic}=1,ROW(),IF({f}={prev},{col(H_SUB)}{r - 1},0))"),
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


def write_paste(ws, report: Report, cfg: Config, lay: _Layout, rows: list[list] | None = None) -> None:
    """Hoja PV-depuración: el bloque de la previsión de la empresa (desde A8), los encabezados de los meses y los
    controles de la asignación. rows: líneas trasladadas del libro anterior (--migrar-desde)."""
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
        f"Pegar desde la celda A{P_FIRST} (pegado especial > valores): EDT cliente, actividad EDT, EDT interna, "
        f"actividad interna o componente de coste, BAC, EAC y un importe por mes, con el mes 1 en la columna "
        f"{col(P_FIRST_MONTH)}. Cada línea se suma en la última EDT cliente de arriba. Las líneas con otras debajo son "
        "subtotales y no se suman: la EDT cliente con EDT interna o códigos más detallados, y la EDT interna (p. ej. "
        "GPR01) con componentes de coste (solo la columna D).")).font = FONT_NOTE

    labels = [(P_TOTAL_ROW, "Total pegado (sin subtotales)",
               f"=SUMIFS({helper}!${col(H_TOTAL)}${P_FIRST}:${col(H_TOTAL)}${P_LAST},"
               f"{helper}!${col(H_COUNT)}${P_FIRST}:${col(H_COUNT)}${P_LAST},1)"),
              (P_ASSIGNED_ROW, "Asignado a filas del libro (meses del horizonte)",
               f"=SUMIFS({helper}!${col(H_HORIZON)}${P_FIRST}:${col(H_HORIZON)}${P_LAST},"
               f"{helper}!${col(H_TARGET)}${P_FIRST}:${col(H_TARGET)}${P_LAST},\">0\")"),
              (P_MISSING_ROW, "Sin asignar (sin fila del libro o fuera del horizonte)",
               f"={control}{P_TOTAL_ROW}-{control}{P_ASSIGNED_ROW}")]
    for r, label, formula in labels:
        ws.cell(row=r, column=P_DESC, value=label).font = FONT_BOLD
        cell = ws.cell(row=r, column=P_EAC, value=formula)
        cell.number_format = money
        cell.border = BORDER
    ws.conditional_formatting.add(f"{control}{P_MISSING_ROW}", FormulaRule(
        formula=[f"ABS({control}{P_MISSING_ROW})>0.5"], fill=FILL_RED))
    ws.cell(row=P_TOTAL_ROW, column=P_FIRST_MONTH,
            value=f"Resultado: filas «(coste)» de la hoja {rs} y hoja '{SHEET_CHART}'; el comando procesar-costes "
                  "lleva el reparto mensual al % planificado.").font = FONT_NOTE
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
            f"Meses posteriores al horizonte del libro ({lay.months} meses): no se asignan a las filas del libro.",
            "evm_export")
    ws.column_dimensions[col(P_CODE)].number_format = "@"
    ws.column_dimensions[col(P_INT_CODE)].number_format = "@"
    ws.column_dimensions[col(P_BAC)].number_format = money
    ws.column_dimensions[col(P_EAC)].number_format = money

    for k, values in enumerate((rows or [])[:P_LINES]):
        for c_, v in enumerate(values, start=1):
            if v is not None and c_ < status_col:
                cell = put_text(ws, P_FIRST + k, c_, v)
                if c_ >= P_BAC and isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = money

    for r in range(P_FIRST, P_LAST + 1):
        h = {x: f"{helper}!{col(x)}{r}" for x in (H_NUM, H_CHILD, H_GROUP, H_PARENT, H_TARGET, H_TOTAL, H_SUBDIFF,
                                                  H_HORIZON)}
        ws.cell(row=r, column=status_col, value=(
            f'=IF(AND({h[H_NUM]}=0,{h[H_CHILD]}=0),"",IF({h[H_PARENT]}=1,"{STATUS_SUBTOTAL}",IF({h[H_TARGET]}>0,'
            f'INDEX(\'{rs}\'!$A$1:$A${lay.total_row + 1},{h[H_TARGET]}),IF({h[H_TARGET]}={NO_ROW},'
            f'"{STATUS_SUMMARY}","{STATUS_NONE}"))))'))
        s = f"{col(status_col)}{r}"
        ws.cell(row=r, column=obs_col, value=(
            f'=IF({s}="","",TRIM(IF(AND({h[H_NUM]}=0,{h[H_GROUP]}=0),"Sin EDT cliente en las líneas anteriores. ","")'
            f'&IF(AND({h[H_PARENT]}=1,{h[H_TOTAL]}<>0,ABS({h[H_SUBDIFF]})>0.5),'
            f'"Subtotal distinto de la suma de sus líneas. ","")'
            f'&IF(AND({h[H_TARGET]}>0,ABS({h[H_TOTAL]}-{h[H_HORIZON]})>0.5),'
            f'"Importes fuera del horizonte del libro. ","")'
            f'&IF(SUMPRODUCT(ISTEXT({amounts}{r}:{last_paste}{r})*(LEN({amounts}{r}:{last_paste}{r})>0))>0,'
            f'"Importes guardados como texto: no se suman. ","")))'))
    status = f"{col(status_col)}{P_FIRST}"
    obs = f"{col(obs_col)}{P_FIRST}"
    area = f"{col(status_col)}{P_FIRST}:{col(obs_col)}{P_LAST}"
    ws.conditional_formatting.add(area, FormulaRule(
        formula=[f'OR(${status}="{STATUS_NONE}",${status}="{STATUS_SUMMARY}",${obs}<>"")'], fill=FILL_RED))
    ws.conditional_formatting.add(f"A{P_FIRST}:{col(P_EAC)}{P_LAST}", FormulaRule(
        formula=[f'${status}="{STATUS_SUBTOTAL}"'], fill=FILL_TOTAL, font=Font(italic=True, color="595959")))

    for c_, w in {P_CODE: 12, P_DESC: 34, P_INT_CODE: 12, P_INT_DESC: 34, P_BAC: 14, P_EAC: 14, status_col - 1: 3,
                  status_col: 20, obs_col: 60}.items():
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
        f"% planificado de MS Project. Previsión: fila «(coste)», con lo pegado en la hoja {SHEET_PASTE}. El objetivo "
        "es que los dos BAC coincidan y que las curvas se parezcan; se ajustan en MS Project las tareas, sus fechas "
        "y el peso de los hitos, y se regenera el libro con --migrar-desde.")).font = FONT_NOTE
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
        (6, "Previsión (fila «(coste)»)", f"=N(INDEX('{rs}'!$D$1:$D${lay.total_row + 1},{sel}+1))", money),
        (7, "Diferencia: previsión - BAC", "=B6-B5", money),
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
