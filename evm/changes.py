"""Registro de cambios: BAC y valores de los meses ya informados (norma EIA-748-E, guías 24 a 26).

'actualizar' y 'analizar' guardan al terminar, en la hoja oculta _registro, la referencia de cada fila de ingreso (y
de cada empresa en multi-empresa): su BAC y el % ganado y el AC de cada mes hasta la fecha de corte. En la ejecución
siguiente comparan el libro con esa referencia:

- BAC distinto: cambio de presupuesto. Como el PV y el EV de cada mes son el BAC × el %, el cambio recalcula los
  meses ya informados (cambio retroactivo) si la fila tiene PV o EV en meses anteriores al mes de corte.
- % ganado o AC distinto en un mes anterior al mes de corte: cambio de un valor ya informado. Un valor vacío cuenta
  como 0, y cargar un valor en el último mes informado (el AC del mes se suele cargar después de 'actualizar') no es
  un cambio.

Cada cambio se agrega a la hoja visible Registro de cambios, con columnas para el motivo y el número de cambio
aprobado; sus filas no se borran. Una fila sin referencia (primera ejecución o fila nueva de --migrar-desde) solo
guarda la suya. Las series se guardan por mes calendario, de modo que la referencia sigue valiendo si
--migrar-desde cambia el mes 1 del horizonte.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from openpyxl.styles import Alignment

from evm.consistency import sentence
from evm.fields import fmt_number
from evm.messages import SEV_INFO, SEV_WARN, Issue
from evm.timeline import add_months, first_of_month
from evm.workbook import (BORDER, FILL_INPUT, FMT_DATE, FMT_MONTH, FONT_NOTE, FONT_TITLE, _header, col, money_format,
                          put_text)

SHEET_LOG = "Registro de cambios"
SHEET_REF = "_registro"                 # referencia de la última ejecución (oculta)
HEAD_ROW = 4
FIRST = 5
(L_DETECTED, L_CUTOFF, L_CODE, L_NAME, L_COMPANY, L_DATA, L_MONTH, L_OLD, L_NEW, L_DELTA, L_AFFECTED, L_REASON,
 L_APPROVAL, L_KEY) = range(1, 15)
HEADERS = {
    L_DETECTED: "Detectado", L_CUTOFF: "Fecha de corte", L_CODE: "Código EDT", L_NAME: "Actividad",
    L_COMPANY: "Empresa", L_DATA: "Dato", L_MONTH: "Mes del valor", L_OLD: "Valor anterior", L_NEW: "Valor nuevo",
    L_DELTA: "Variación", L_AFFECTED: "Meses informados afectados", L_REASON: "Motivo",
    L_APPROVAL: "N.º de cambio aprobado", L_KEY: "clave",
}
WIDTHS = {L_DETECTED: 16, L_CUTOFF: 12, L_CODE: 11, L_NAME: 34, L_COMPANY: 12, L_DATA: 10, L_MONTH: 10, L_OLD: 14,
          L_NEW: 14, L_DELTA: 14, L_AFFECTED: 11, L_REASON: 40, L_APPROVAL: 14}
INPUT_COLS = (L_REASON, L_APPROVAL)
DATA_BAC, DATA_EV, DATA_AC = "BAC", "% ganado", "AC"
FMT_DETECTED = "dd/mm/yyyy hh:mm"
FMT_PCT_LOG = "0.00%"
EPS_PCT = 1e-9
EPS_MONEY = 0.005
SHOWN = 6                               # cambios de una fila que se detallan en su observación


@dataclass
class RowRef:
    """Referencia de una fila (y empresa): BAC y valores de cada mes desde el inicio de la referencia."""

    bac: float
    ev: list[float]                     # % ganado de cada mes (vacío = 0)
    ac: list[float]


@dataclass
class Reference:
    start: date                         # mes de la primera posición de las series
    cutoff: date                        # fecha de corte con que se guardó
    rows: dict[tuple[str, str], RowRef]     # (clave de la fila, empresa: "", "A" o "B") -> referencia


@dataclass
class Change:
    key: str
    code: str
    name: str
    company: str                        # nombre de la empresa ("" en un libro de una sola empresa)
    data: str                           # DATA_BAC, DATA_EV o DATA_AC
    month: date | None                  # mes del valor (None en el BAC)
    old: float
    new: float
    affected: int | None = None         # BAC: meses informados con PV o EV recalculados


# ----------------------------------------------------------------------------
# Referencia (hoja oculta _registro)
# ----------------------------------------------------------------------------
def _pack(values) -> str:
    return ";".join("" if not v else repr(float(v)) for v in values)


def _unpack(text) -> list[float]:
    out = []
    for part in str(text or "").split(";"):
        try:
            out.append(float(part) if part.strip() else 0.0)
        except ValueError:
            out.append(0.0)
    return out


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def reference_from_book(book, idx: int, status: date) -> Reference:
    """Referencia con los valores del libro hasta el mes idx del horizonte."""
    rows = {}
    for s in book.input_rows:
        for role, _company in book.roles():
            rows[(s.key, role)] = RowRef(book.bac_role[role][s.row],
                                         [v or 0.0 for v in book.pct_ev[role][s.row][: idx + 1]],
                                         [v or 0.0 for v in book.ac_of(role)[s.row][: idx + 1]])
    return Reference(book.horizon.start, status, rows)


def read_reference(wb) -> Reference | None:
    """Referencia guardada por la ejecución anterior (None si el libro no la tiene)."""
    if SHEET_REF not in wb.sheetnames:
        return None
    ws = wb[SHEET_REF]
    head = list(next(ws.iter_rows(min_row=1, max_row=1, max_col=4, values_only=True), ())) + [None] * 4
    start, cutoff = _as_date(head[1]), _as_date(head[3])
    if start is None or cutoff is None:
        return None
    rows = {}
    for vals in ws.iter_rows(min_row=3, max_col=5, values_only=True):
        vals = list(vals) + [None] * (5 - len(vals))
        key = str(vals[0] or "").strip()
        if not key:
            continue
        bac = vals[2] if isinstance(vals[2], (int, float)) and not isinstance(vals[2], bool) else 0.0
        rows[(key, str(vals[1] or "").strip())] = RowRef(float(bac), _unpack(vals[3]), _unpack(vals[4]))
    return Reference(first_of_month(start), cutoff, rows)


def write_reference(wb, ref: Reference) -> None:
    """Guarda la referencia en la hoja oculta _registro (la crea si no existe)."""
    if SHEET_REF in wb.sheetnames:
        ws = wb[SHEET_REF]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(SHEET_REF)
    ws.sheet_state = "hidden"
    ws.append(["inicio", ref.start, "corte", ref.cutoff])
    ws.append(["clave", "empresa", "BAC", "% EV", "AC"])
    for r, ((key, role), x) in enumerate(ref.rows.items(), start=3):
        put_text(ws, r, 1, key)
        put_text(ws, r, 2, role)
        ws.cell(row=r, column=3, value=x.bac)
        ws.cell(row=r, column=4, value=_pack(x.ev))
        ws.cell(row=r, column=5, value=_pack(x.ac))


# ----------------------------------------------------------------------------
# Detección de cambios
# ----------------------------------------------------------------------------
def detect(book, ref: Reference | None, idx: int) -> list[Change]:
    """Cambios del libro frente a la referencia de la ejecución anterior, para la fecha de corte del mes idx."""
    if ref is None:
        return []
    last = first_of_month(ref.cutoff)               # último mes informado
    out: list[Change] = []
    for s in book.input_rows:
        for role, company in book.roles():
            r = ref.rows.get((s.key, role))
            if r is None:
                continue
            bac = book.bac_role[role][s.row]
            # El primer BAC de una fila (la referencia no tenía) no es un cambio: se carga después del primer
            # 'actualizar' (ACT-022)
            if abs(r.bac) > EPS_MONEY and abs(bac - r.bac) > EPS_MONEY:
                pv, ev = book.pct_pv[role][s.row], book.pct_ev[role][s.row]
                affected = sum(1 for i in range(min(idx, len(pv)))
                               if abs(pv[i]) > EPS_PCT or abs(ev[i] or 0.0) > EPS_PCT)
                out.append(Change(s.key, s.code, s.name, company, DATA_BAC, None, r.bac, bac, affected))
            for data, old_series, new_series, eps in ((DATA_EV, r.ev, book.pct_ev[role][s.row], EPS_PCT),
                                                      (DATA_AC, r.ac, book.ac_of(role)[s.row], EPS_MONEY)):
                for i, old in _by_book_month(old_series, ref.start, idx, book.horizon):
                    month = book.horizon.month_start(i)
                    new = (new_series[i] or 0.0) if i < len(new_series) else 0.0
                    if abs(new - old) <= eps or (abs(old) <= eps and month == last):
                        continue
                    out.append(Change(s.key, s.code, s.name, company, data, month, old, new))
    return out


def _by_book_month(series: list[float], start: date, idx: int, horizon) -> list[tuple[int, float]]:
    """Valores de la referencia por mes del horizonte del libro, de los meses anteriores al mes de corte (idx). Los
    meses anteriores al mes 1 (o posteriores al último) se suman en el mes 1 (o en el último), como hace
    --migrar-desde con el historial cuando cambia el horizonte (MIG-011 y MIG-012)."""
    out: dict[int, float] = {}
    for j, value in enumerate(series):
        i = min(max(horizon.index(add_months(start, j)), 0), horizon.months - 1)
        out[i] = out.get(i, 0.0) + value
    return sorted((i, v) for i, v in out.items() if i < idx)


def first_reference() -> Issue:
    return Issue("ACT-038", message="Se guardó la referencia del registro de cambios (BAC, % ganado y AC de cada "
                                    "fila): a partir de la próxima ejecución, los cambios se agregan a la hoja "
                                    "Registro de cambios.")


def _value(data: str, v: float, decimals: int) -> str:
    return f"{fmt_number(v * 100, 2, trim=True)} %" if data == DATA_EV else fmt_number(v, decimals)


def _signed(data: str, v: float, decimals: int) -> str:
    text = _value(data, v, decimals)
    return text if text.startswith("-") else f"+{text}"


def issues(changes: list[Change], decimals: int, registered: bool = True) -> list[Issue]:
    """Observaciones de la consola: una por cambio de BAC (ACT-036) y una por fila con valores de meses ya informados
    modificados (ACT-037)."""
    where = ("Se agregó a la hoja Registro de cambios." if registered else
             "Se registrará en la hoja Registro de cambios al actualizar sin --simular.")
    out: list[Issue] = []
    values: dict[tuple[str, str], list[Change]] = {}
    for c in changes:
        prefix = f"{c.company}: " if c.company else ""
        if c.data == DATA_BAC:
            if c.affected:
                effect = (f"se recalcularon el PV y el EV de {c.affected} "
                          f"{'mes ya informado' if c.affected == 1 else 'meses ya informados'} (cambio retroactivo)")
            else:
                effect = "la fila no tiene meses informados"
            out.append(Issue("ACT-036", c.code, c.name,
                             f"{prefix}BAC {_value(c.data, c.old, decimals)} → {_value(c.data, c.new, decimals)} "
                             f"({_signed(c.data, c.new - c.old, decimals)}): {effect}. {where}",
                             SEV_WARN if c.affected else SEV_INFO))
        else:
            values.setdefault((c.key, c.company), []).append(c)
    for (_key, company), lst in values.items():
        first = lst[0]
        detail = "; ".join(f"{c.data} de {c.month:%m/%Y}: {_value(c.data, c.old, decimals)} → "
                           f"{_value(c.data, c.new, decimals)}" for c in lst[:SHOWN])
        more = f" y {len(lst) - SHOWN} más" if len(lst) > SHOWN else ""
        prefix = f"{company}: " if company else ""
        out.append(Issue("ACT-037", first.code, first.name, sentence(
            prefix, f"{'cambió un valor' if len(lst) == 1 else 'cambiaron valores'} de meses ya informados "
                    f"({detail}{more}). {where}")))
    return out


# ----------------------------------------------------------------------------
# Hoja Registro de cambios
# ----------------------------------------------------------------------------
def _last_row(ws) -> int:
    """Última fila con datos del registro (HEAD_ROW si no tiene filas)."""
    last = HEAD_ROW
    for r in range(FIRST, ws.max_row + 1):
        if any(ws.cell(row=r, column=c_).value not in (None, "") for c_ in range(1, L_KEY + 1)):
            last = r
    return last


def _format_row(ws, r: int, data: str, decimals: int) -> None:
    number = FMT_PCT_LOG if data == DATA_EV else money_format(decimals)
    for c_ in range(1, L_KEY):
        cell = ws.cell(row=r, column=c_)
        cell.border = BORDER
        cell.alignment = Alignment(vertical="top", wrap_text=c_ in (L_NAME, L_REASON))
        if c_ in INPUT_COLS:
            cell.fill = FILL_INPUT
    ws.cell(row=r, column=L_DETECTED).number_format = FMT_DETECTED
    ws.cell(row=r, column=L_CUTOFF).number_format = FMT_DATE
    ws.cell(row=r, column=L_MONTH).number_format = FMT_MONTH
    for c_ in (L_OLD, L_NEW, L_DELTA):
        ws.cell(row=r, column=c_).number_format = number
    ws.cell(row=r, column=L_AFFECTED).alignment = Alignment(horizontal="center", vertical="top")


def _finish(ws, last: int) -> None:
    ws.auto_filter.ref = f"A{HEAD_ROW}:{col(L_KEY)}{max(last, FIRST)}"


def write_sheet(wb, rows: list[list] | None = None, decimals: int = 0):
    """Crea la hoja Registro de cambios (después de Análisis de desviaciones) con las filas indicadas, o la deja
    como está si ya existe. rows: valores de las columnas A a N de cada fila (de un libro anterior)."""
    if SHEET_LOG in wb.sheetnames:
        return wb[SHEET_LOG]
    from evm.analysis import SHEET_ANALYSIS
    index = wb.sheetnames.index(SHEET_ANALYSIS) + 1 if SHEET_ANALYSIS in wb.sheetnames else len(wb.sheetnames)
    ws = wb.create_sheet(SHEET_LOG, index)
    ws.cell(row=1, column=1, value=SHEET_LOG).font = FONT_TITLE
    ws.cell(row=2, column=1, value=(
        "Cambios del BAC y de los valores de meses ya informados (% ganado y AC) que detectan actualizar y "
        "analizar. Un cambio del BAC recalcula el PV y el EV de todos los meses de la fila. Las columnas amarillas "
        "se completan a mano: el motivo y el número de cambio aprobado. Las filas no se borran.")).font = FONT_NOTE
    _header(ws, HEAD_ROW, list(HEADERS.items()))
    for c_, w in WIDTHS.items():
        ws.column_dimensions[col(c_)].width = w
    ws.column_dimensions[col(L_KEY)].hidden = True
    ws.freeze_panes = ws.cell(row=FIRST, column=L_NAME + 1)
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"
    ws.page_setup.orientation = "landscape"
    r = HEAD_ROW
    for r, vals in enumerate(rows or [], start=FIRST):
        for c_, v in enumerate(vals[:L_KEY], start=1):
            if isinstance(v, str):
                put_text(ws, r, c_, v)
            else:
                ws.cell(row=r, column=c_, value=v)
        _format_row(ws, r, str(vals[L_DATA - 1] or ""), decimals)
    _finish(ws, r)
    return ws


def append(wb, changes: list[Change], detected: datetime, cutoff: date, decimals: int) -> None:
    """Agrega los cambios al final de la hoja Registro de cambios (la crea en los libros anteriores)."""
    ws = write_sheet(wb, decimals=decimals)
    r = _last_row(ws)
    for c in changes:
        r += 1
        ws.cell(row=r, column=L_DETECTED, value=detected)
        ws.cell(row=r, column=L_CUTOFF, value=cutoff)
        put_text(ws, r, L_CODE, c.code)
        put_text(ws, r, L_NAME, c.name)
        put_text(ws, r, L_COMPANY, c.company)
        put_text(ws, r, L_DATA, c.data)
        ws.cell(row=r, column=L_MONTH, value=c.month)
        ws.cell(row=r, column=L_OLD, value=c.old)
        ws.cell(row=r, column=L_NEW, value=c.new)
        ws.cell(row=r, column=L_DELTA, value=c.new - c.old)
        ws.cell(row=r, column=L_AFFECTED, value=c.affected)
        put_text(ws, r, L_KEY, c.key)
        _format_row(ws, r, c.data, decimals)
    _finish(ws, r)


def read_rows(ws) -> list[list]:
    """Filas del registro (columnas A a N), para trasladarlas con --migrar-desde."""
    out = []
    for vals in ws.iter_rows(min_row=FIRST, max_col=L_KEY, values_only=True):
        vals = list(vals) + [None] * (L_KEY - len(vals))
        if any(v not in (None, "") for v in vals):
            out.append(vals)
    return out
