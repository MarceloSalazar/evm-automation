"""Comando procesar-costes: revisión del plan de costes de la hoja PV-depuración y, con --calculos, sus cálculos.

Lee la hoja PV-depuración de un libro con depuración de costes: solo las líneas de EDT cliente cuyo código es el de una
fila de ingreso de la hoja PV de la empresa que se depura (cost_review.read_lines, las mismas reglas que las fórmulas
de la hoja _depuracion, que llevan esas líneas a las filas «(coste)»). Informa las líneas que no se leen, los importes
después del horizonte, las EDT cliente repetidas y las filas sin plan, y lista el EAC y el BAC de cada fila con plan.

Con --calculos ([exportacion] depuracion_calculos) escribe en la hoja PV-depuración las fórmulas de los cálculos del
plan (cost_review.calc_plan y write_calcs): las sumas de los componentes en su EDT interna y de las EDT interna en su
EDT cliente, el EAC y el BAC de la EDT cliente. Sin esa opción, usa los importes que tiene la hoja y no modifica el
libro.

No modifica el % planificado: el PV de todas las filas es el de MS Project, y el plan de costes se compara con él en las
filas «(coste)» y en la hoja 'Depuración costes'. Antes de la versión 2.6.0, procesar-costes escribía el reparto del
plan en la hoja '% PV' y marcaba esas filas en la hoja _estructura: en esos libros se avisa (DEP-021) y --migrar-desde
les devuelve el % de MS Project (DEP-020).
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
                        plan_rows_warning, read_book)
from evm.wbs import fmt_money, plural
from evm.workbook import SHEET_PV, company_sheet

TOLERANCE = 0.005                       # importe que se considera 0


@dataclass
class CostPlanResult:
    sheet: str                          # hoja PV de la empresa que se procesó ('PV-EmpresaA' o 'PV')
    lines: int                          # líneas de EDT cliente de la hoja PV-depuración
    counted: int                        # líneas de EDT cliente que se leen (con fila de ingreso del libro)
    total: float                        # importe de las líneas que se leen
    assigned: float                     # parte de ese importe en las filas con plan, dentro del horizonte
    rows: list[tuple[str, str, float, float | None]]  # filas con plan: código, actividad, EAC de su línea y BAC
    inputs: int                         # filas de ingreso del libro
    calculated: int | None              # líneas con los cálculos del plan escritos (None: sin --calculos)
    warnings: list[Issue]
    backup: Path | None
    dry_run: bool
    saved: bool                         # el libro se guardó (solo con los cálculos, fuera de la simulación)
    book: BookData


def _company_role(book: BookData, sheet: str) -> tuple[str, str]:
    """Empresa ('A', 'B' o '') y nombre de la hoja PV que se depura."""
    for role, company in book.roles():
        if (company_sheet(SHEET_PV, company) if company else SHEET_PV) == sheet:
            return role, company
    raise InputError("DEP-012", f"El libro no tiene la hoja {sheet} que indica su depuración de costes: se debe "
                                "generar de nuevo con depuracion_costes = true.")


def _label(x: cr.PasteLine) -> str:
    """Línea de EDT cliente en los mensajes: su código, como está escrito, y su fila."""
    return f"{str(x.values[cr.P_CODE - 1]).strip()} (fila {x.row})"


def _listed(items: list[str], limit: int = 10) -> str:
    return ", ".join(items[:limit]) + (f" y {len(items) - limit} más" if len(items) > limit else "")


def run_cost_plan(cfg: Config, book_path: str | Path, backup: bool = True, dry_run: bool = False) -> CostPlanResult:
    """Comando 'procesar-costes': revisa el plan de costes de la hoja PV-depuración y, con cfg.cost_review_calcs,
    escribe sus cálculos. No modifica el % planificado."""
    book_path = Path(book_path)
    progress.stage("libro", "Lectura del libro")
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    if cr.SHEET_PASTE not in wb.sheetnames:
        raise InputError("DEP-012", f"El libro no tiene la hoja {cr.SHEET_PASTE}: procesar-costes usa un libro "
                                    "generado con depuracion_costes = true, con el plan de costes en esa hoja.")
    check_sheets(wb.sheetnames, book.meta, charts=True)
    sheet = str(book.meta.get("depuracion_hoja") or SHEET_PV)
    role, _company = _company_role(book, sheet)
    months = book.horizon.months
    save = cfg.cost_review_calcs and not dry_run
    warnings = fixable_warning(book, dry_run) if cfg.cost_review_calcs else []

    progress.stage("procesar", "Plan de costes")
    ws_paste = wb[cr.SHEET_PASTE]
    rows, _fmt = cr.read_paste(ws_paste, months)
    rows, _formulas, missing_cells = cr.resolve_formulas(book_path, rows, months)
    calculated = None
    if cfg.cost_review_calcs:
        rows, plan = cr.calc_plan(rows)
        missing_cells = []              # las celdas calculadas toman el valor de los cálculos
        calculated = sum(1 for x in plan if x.kind)
        if save:
            cr.write_calcs(ws_paste, plan, months, sheet, int(book.meta["fila_total"]),
                           int(book.meta.get("decimales") or 0))
    lines = cr.read_lines(rows, cr.structure_table(book.structure))
    read = [x for x in lines if x.state == cr.STATE_READ]
    read_rows = {x.row - cr.P_FIRST for x in read}
    missing = sum(1 for k, c in missing_cells if k in read_rows and c >= cr.P_FIRST_MONTH - 1)
    amounts: dict[int, list[float]] = {}            # fila de ingreso del libro -> importe de cada mes del plan
    for x in read:
        vec = amounts.setdefault(x.match, [0.0] * len(x.months))
        for i, v in enumerate(x.months):
            if cr.is_amount(v):
                vec[i] += v
    inputs = book.input_rows
    if not any(s.row in amounts for s in inputs):
        raise InputError("DEP-013", f"Ninguna línea de EDT cliente de la hoja {cr.SHEET_PASTE} tiene el código de una "
                                    f"fila de ingreso de la hoja {sheet}: se debe escribir el plan de costes desde la "
                                    f"celda A{cr.P_FIRST} y revisar la columna 'Asignado a'.")

    with_plan: list[tuple] = []                     # (fila, EAC de su línea)
    no_lines, zero = [], []
    for s in inputs:
        vec = amounts.get(s.row)
        if vec is None:
            no_lines.append(s.code)
        elif all(abs(v) < TOLERANCE for v in vec):
            zero.append(s.code)
        else:
            with_plan.append((s, sum(vec)))
    code_of = {s.row: s.code for s in book.structure}
    warnings += _warnings(lines, amounts, months, missing, no_lines, zero, sheet, code_of, cfg.cost_review_calcs)
    warnings += plan_rows_warning(book)

    backup_path = None
    if save:
        _write_fixable(wb, book)
        idx = book.horizon.index(book.status_date) if book.status_date is not None else None
        _refresh_charts(wb, book, idx if idx is not None and 0 <= idx < months else None, book.meta)
        progress.stage("guardar", "Guardado del libro")
        backup_path = _save(wb, book_path, backup, cfg.fast_save)
    rows_done = {s.row for s, _t in with_plan}
    assigned = sum(sum(v[:months]) for r, v in amounts.items() if r in rows_done)
    return CostPlanResult(sheet, sum(1 for x in lines if x.client), len(read), sum(x.total for x in read), assigned,
                          [(s.code, s.name, t, book.bac_role[role][s.row]) for s, t in with_plan], len(inputs),
                          calculated, warnings, backup_path, dry_run, save, book)


def _warnings(lines: list[cr.PasteLine], amounts: dict[int, list[float]], months: int, missing: int,
              no_lines: list[str], zero: list[str], sheet: str, code_of: dict[int, str], calcs: bool) -> list[Issue]:
    """Observaciones de procesar-costes: líneas sin leer, meses fuera del horizonte, filas sin plan e importes que no
    se pueden sumar."""
    out: list[Issue] = []
    unread = [x for x in lines if x.state in cr.UNREAD and abs(x.total) >= TOLERANCE]
    if unread:
        items = [f"{_label(x)}: {cr.STATUS_OF[x.state].lower()}" for x in unread]
        out.append(Issue("DEP-014", message=(
            f"{plural(len(unread), 'línea', 'líneas')} de EDT cliente de la hoja {cr.SHEET_PASTE} sin leer, por "
            f"{fmt_money(sum(x.total for x in unread))}: {_listed(items)}. Su importe no llega a ninguna fila "
            "«(coste)».")))
    repeated = sorted({x.key for x in lines if x.state == cr.STATE_READ and x.repeated})
    if repeated:
        out.append(Issue("DEP-019", message=(
            f"EDT cliente repetidas en la hoja {cr.SHEET_PASTE} ({_listed(repeated)}): se suman sus líneas.")))
    later = {r: sum(v[months:]) for r, v in amounts.items() if abs(sum(v[months:])) >= TOLERANCE}
    if later:
        out.append(Issue("DEP-015", message=(
            f"Hay {fmt_money(sum(later.values()))} del plan de costes en meses posteriores al horizonte del libro, en "
            f"{plural(len(later), 'fila', 'filas')} ({_listed([code_of.get(r, str(r)) for r in sorted(later)])}): "
            "forman parte del EAC de su fila «(coste)», pero no de sus meses.")))
    if no_lines:
        out.append(Issue("DEP-016", message=(
            f"Filas de ingreso de la hoja {sheet} sin línea de EDT cliente en el plan de costes ({len(no_lines)}): "
            f"{_listed(no_lines)}. Su fila «(coste)» queda vacía.")))
    if zero:
        hint = "" if calcs else " Si faltan los cálculos del plan, se puede ejecutar procesar-costes --calculos."
        out.append(Issue("DEP-017", message=(
            f"Filas cuya línea de EDT cliente no tiene importes en el plan de costes ({len(zero)}): {_listed(zero)}. "
            f"Su fila «(coste)» queda en 0.{hint}")))
    texts = [x for x in lines if x.state == cr.STATE_READ and any(
        isinstance(v, str) and v.strip() for v in x.months)]
    if texts or missing:
        parts = []
        if texts:
            parts.append(f"{plural(len(texts), 'línea tiene', 'líneas tienen')} importes guardados como texto, que no "
                         f"se suman ({_listed([_label(x) for x in texts])})")
        if missing:
            parts.append(f"{plural(missing, 'celda tiene', 'celdas tienen')} una fórmula sin valor calculado: se debe "
                         "abrir y guardar el libro en Excel, o ejecutar procesar-costes --calculos")
        out.append(Issue("DEP-018", message=f"Hoja {cr.SHEET_PASTE}: " + "; ".join(parts) + "."))
    return out
