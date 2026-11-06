"""Comando procesar-costes: % planificado de cada fila con el reparto mensual del plan de costes.

Lee la hoja PV-depuración de un libro con depuración de costes, asigna cada línea a una fila del libro con las mismas
reglas que las fórmulas de la hoja _depuracion (cost_review.assign_lines) y escribe en la hoja '% PV' de la empresa que
se depura el % planificado de cada fila de ingreso con líneas del plan: el importe de cada mes dividido por el total de
sus líneas (el EAC de la fila). El PV del mes sigue siendo el BAC × el %: si el BAC de la fila coincide con su EAC, el
PV es el importe del plan; si no, el plan da el reparto en el tiempo y el total es el BAC.

Las filas procesadas quedan marcadas en la hoja _estructura: 'actualizar' no compara su % planificado con la línea base
de MS Project y --migrar-desde lo conserva en el libro nuevo (también en el libro de entrega, sin depuración). Las
filas sin líneas en el plan conservan su % planificado.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from evm import cost_review as cr
from evm import progress
from evm.config import Config
from evm.messages import InputError, Issue
from evm.update import (BookData, _refresh_charts, _save, _write_fixable, check_sheets, fixable_warning,
                        read_book)
from evm.wbs import fmt_money, plural
from evm.workbook import (FIRST_MONTH_COL, FIRST_ROW, FMT_PCT_MONTH, FONT_COST, PCT_PREFIX, SHEET_PV, SHEET_STRUCT,
                          STRUCT_PLAN_COL, STRUCT_PLAN_HEAD, TITLE_ROW, company_sheet, pct_pv_title, plan_text,
                          put_text, styled)

TOLERANCE = 0.005                       # importe que se considera 0


@dataclass
class CostPlanResult:
    sheet: str                          # hoja PV de la empresa que se procesó ('PV-EmpresaA' o 'PV')
    lines: int                          # líneas pegadas con contenido
    counted: int                        # líneas que se suman (sin los subtotales)
    total: float                        # importe de las líneas que se suman
    assigned: float                     # parte del total asignada a filas procesadas, dentro del horizonte
    rows: list[tuple[str, str, float]]  # filas procesadas: código, actividad y EAC (total de sus líneas)
    inputs: int                         # filas de ingreso del libro
    warnings: list[Issue]
    backup: Path | None
    dry_run: bool
    book: BookData


def _company_role(book: BookData, sheet: str) -> tuple[str, str]:
    """Empresa ('A', 'B' o '') y nombre de la hoja PV que se depura."""
    for role, company in book.roles():
        if (company_sheet(SHEET_PV, company) if company else SHEET_PV) == sheet:
            return role, company
    raise InputError("DEP-012", f"El libro no tiene la hoja {sheet} que indica su depuración de costes: se debe "
                                "generar de nuevo con depuracion_costes = true.")


def _paste_rows(path: Path, ws, months: int) -> tuple[list[list], int]:
    """Líneas de la hoja PV-depuración en el formato actual. Las celdas con fórmula (pegadas sin 'valores') toman el
    valor que calculó Excel; devuelve también cuántas no tienen valor calculado."""
    rows, _fmt = cr.read_paste(ws, months)
    cells = [(k, c) for k, values in enumerate(rows) for c, v in enumerate(values)
             if isinstance(v, str) and v.startswith("=")]
    if not cells:
        return rows, 0
    wv = load_workbook(path, data_only=True, read_only=True)
    try:
        cached, _fmt = cr.read_paste(wv[cr.SHEET_PASTE], months)
    finally:
        wv.close()
    missing = 0
    for k, c in cells:
        value = cached[k][c] if k < len(cached) and c < len(cached[k]) else None
        missing += value is None
        rows[k][c] = value
    return rows, missing


def _label(x: cr.PasteLine) -> str:
    """Línea pegada en los mensajes: su código (EDT cliente, EDT interna o componente) y su fila."""
    values = x.values + [None] * cr.P_FIRST_MONTH
    text = next((str(values[c - 1]).strip() for c in (cr.P_CODE, cr.P_INT_CODE, cr.P_INT_DESC)
                 if str(values[c - 1] or "").strip()), "")
    return f"{text} (fila {x.row})" if text else f"fila {x.row}"


def _listed(items: list[str], limit: int = 10) -> str:
    return ", ".join(items[:limit]) + (f" y {len(items) - limit} más" if len(items) > limit else "")


def run_cost_plan(cfg: Config, book_path: str | Path, backup: bool = True, dry_run: bool = False) -> CostPlanResult:
    """Comando 'procesar-costes': escribe en la hoja '% PV' de la empresa que se depura el reparto mensual del plan
    de costes pegado en la hoja PV-depuración."""
    book_path = Path(book_path)
    progress.stage("libro", "Lectura del libro")
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    if cr.SHEET_PASTE not in wb.sheetnames:
        raise InputError("DEP-012", f"El libro no tiene la hoja {cr.SHEET_PASTE}: procesar-costes usa un libro "
                                    "generado con depuracion_costes = true, con el plan de costes pegado en esa hoja.")
    check_sheets(wb.sheetnames, book.meta, charts=True)
    sheet = str(book.meta.get("depuracion_hoja") or SHEET_PV)
    role, company = _company_role(book, sheet)
    months = book.horizon.months
    warnings = fixable_warning(book, dry_run)

    progress.stage("procesar", "Reparto del plan de costes")
    rows, missing = _paste_rows(book_path, wb[cr.SHEET_PASTE], months)
    lines = cr.assign_lines(rows, cr.structure_table(book.structure))
    counted = [x for x in lines if x.counted]
    amounts: dict[int, list[float]] = {}            # fila del libro -> importe de cada mes del plan
    for x in counted:
        if x.target > 0:
            vec = amounts.setdefault(x.target, [0.0] * len(x.months))
            for i, v in enumerate(x.months):
                if cr.is_amount(v):
                    vec[i] += v
    inputs = book.input_rows
    if not any(s.row in amounts for s in inputs):
        raise InputError("DEP-013", f"Ninguna línea de la hoja {cr.SHEET_PASTE} se asigna a una fila del libro: se "
                                    f"debe pegar el plan de costes desde la celda A{cr.P_FIRST} y revisar la columna "
                                    "'Asignado a'.")

    processed: list[tuple] = []                     # (fila, EAC, % de cada mes)
    no_lines, zero = [], []
    for s in inputs:
        vec = amounts.get(s.row)
        if vec is None:
            no_lines.append(s.code)
        elif abs(sum(vec)) < TOLERANCE:
            zero.append(s.code)
        else:
            total = sum(vec)
            processed.append((s, total, [v / total for v in vec[:months]]))
    code_of = {s.row: s.code for s in book.structure}
    warnings += _warnings(lines, counted, amounts, months, missing, no_lines, zero, sheet, code_of)

    if not dry_run:
        ws_pct = wb[PCT_PREFIX + sheet]
        ws_struct = wb[SHEET_STRUCT]
        ws_struct.cell(row=1, column=STRUCT_PLAN_COL, value=STRUCT_PLAN_HEAD)
        for s, _total, pct in processed:
            for i, p in enumerate(pct):
                cell = ws_pct.cell(row=s.row, column=FIRST_MONTH_COL + i)
                cell.value = p if abs(p) > 1e-12 else None
                if cell.value is not None:
                    styled(cell, number_format=FMT_PCT_MONTH, font=FONT_COST)
            put_text(ws_struct, s.row - (FIRST_ROW - 2), STRUCT_PLAN_COL, plan_text(set(s.plan_pv) | {role}))
            book.pct_pv[role][s.row] = list(pct)
        ws_pct.cell(row=TITLE_ROW, column=1, value=pct_pv_title(company, True))
        _write_fixable(wb, book)
        idx = book.horizon.index(book.status_date) if book.status_date is not None else None
        _refresh_charts(wb, book, idx if idx is not None and 0 <= idx < months else None, book.meta)
        progress.stage("guardar", "Guardado del libro")
    backup_path = None if dry_run else _save(wb, book_path, backup, cfg.fast_save)
    total = sum(x.total for x in counted)
    rows_done = {s.row for s, _t, _p in processed}
    assigned = sum(sum(v[:months]) for r, v in amounts.items() if r in rows_done)
    return CostPlanResult(sheet, sum(1 for x in lines if x.numeric or x.child), len(counted), total, assigned,
                          [(s.code, s.name, t) for s, t, _p in processed], len(inputs), warnings, backup_path,
                          dry_run, book)


def _warnings(lines: list[cr.PasteLine], counted: list[cr.PasteLine], amounts: dict[int, list[float]], months: int,
              missing: int, no_lines: list[str], zero: list[str], sheet: str, code_of: dict[int, str]) -> list[Issue]:
    """Observaciones de procesar-costes: líneas sin fila del libro, meses fuera del horizonte, filas sin reparto e
    importes que no se pueden sumar."""
    out: list[Issue] = []
    unassigned = [x for x in counted if x.target <= 0 and abs(x.total) >= TOLERANCE]
    if unassigned:
        why = {cr.NO_ROW: "resumen sin fila propia"}
        items = [f"{_label(x)}: {why.get(x.target, 'sin EDT cliente encima' if not x.effective else 'sin fila')}"
                 for x in unassigned]
        out.append(Issue("DEP-014", message=(
            f"{plural(len(unassigned), 'línea', 'líneas')} de la hoja {cr.SHEET_PASTE} sin fila del libro, por "
            f"{fmt_money(sum(x.total for x in unassigned))}: {_listed(items)}. Su importe no forma parte del % "
            "planificado.")))
    later = {r: sum(v[months:]) for r, v in amounts.items() if abs(sum(v[months:])) >= TOLERANCE}
    if later:
        out.append(Issue("DEP-015", message=(
            f"Hay {fmt_money(sum(later.values()))} del plan de costes en meses posteriores al horizonte del libro, en "
            f"{plural(len(later), 'fila', 'filas')} ({_listed([code_of.get(r, str(r)) for r in sorted(later)])}): "
            "no forman parte del % planificado, que en esas filas no suma 100 %.")))
    if no_lines:
        out.append(Issue("DEP-016", message=(
            f"Filas de ingreso de la hoja {sheet} sin líneas en el plan de costes ({len(no_lines)}): "
            f"{_listed(no_lines)}. Conservan su % planificado (de MS Project o de un procesado anterior).")))
    if zero:
        out.append(Issue("DEP-017", message=(
            f"Filas cuyas líneas del plan de costes suman 0 ({len(zero)}): {_listed(zero)}. No hay reparto que "
            "aplicar: conservan su % planificado.")))
    texts = [x for x in lines if (x.numeric or x.child) and any(
        isinstance(v, str) and v.strip() for v in x.values[cr.P_BAC - 1:])]
    if texts or missing:
        parts = []
        if texts:
            parts.append(f"{plural(len(texts), 'línea tiene', 'líneas tienen')} importes guardados como texto, que no "
                         f"se suman ({_listed([_label(x) for x in texts])})")
        if missing:
            parts.append(f"{plural(missing, 'celda tiene', 'celdas tienen')} una fórmula sin valor calculado: se debe "
                         "pegar el plan con pegado especial > valores, o abrir y guardar el libro en Excel")
        out.append(Issue("DEP-018", message=f"Hoja {cr.SHEET_PASTE}: " + "; ".join(parts) + "."))
    return out
