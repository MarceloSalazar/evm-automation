"""Comando procesar-costes: % planificado de cada fila con el importe mensual del plan de costes.

Lee la hoja PV-depuración de un libro con depuración de costes: solo las líneas de EDT cliente cuyo código es el de una
fila de ingreso de la hoja PV de la empresa que se depura (cost_review.read_lines, las mismas reglas que las fórmulas
de la hoja _depuracion). En la hoja '% PV' de la empresa escribe el % planificado de cada una de esas filas: el importe
de cada mes dividido por el BAC de la fila, de modo que el PV (BAC × %) es el importe del plan.

Con --calculos ([exportacion] depuracion_calculos), antes escribe en la hoja PV-depuración las fórmulas de los
cálculos del plan (cost_review.calc_plan y write_calcs): las sumas de los componentes en su EDT interna y de las EDT
interna en su EDT cliente, el EAC y el BAC de la EDT cliente. Sin esa opción, usa los importes que tiene la hoja.

Las filas procesadas quedan marcadas en la hoja _estructura: 'actualizar' no compara su % planificado con la línea base
de MS Project y --migrar-desde lo conserva en el libro nuevo (también en el libro de entrega, sin depuración). Las
filas sin línea en el plan, sin BAC o con el plan en 0 conservan su % planificado.
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
                          STRUCT_PLAN_COL, STRUCT_PLAN_HEAD, TITLE_ROW, company_sheet, pct_pv_title, plan_text, put_text,
                          styled)

TOLERANCE = 0.005                       # importe que se considera 0


@dataclass
class CostPlanResult:
    sheet: str                          # hoja PV de la empresa que se procesó ('PV-EmpresaA' o 'PV')
    lines: int                          # líneas de EDT cliente de la hoja PV-depuración
    counted: int                        # líneas de EDT cliente que se leen (con fila de ingreso del libro)
    total: float                        # importe de las líneas que se leen
    assigned: float                     # parte de ese importe en las filas procesadas, dentro del horizonte
    rows: list[tuple[str, str, float]]  # filas procesadas: código, actividad y EAC (importe de su línea)
    inputs: int                         # filas de ingreso del libro
    calculated: int | None              # líneas con los cálculos del plan escritos (None: sin --calculos)
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


def _label(x: cr.PasteLine) -> str:
    """Línea de EDT cliente en los mensajes: su código, como está escrito, y su fila."""
    return f"{str(x.values[cr.P_CODE - 1]).strip()} (fila {x.row})"


def _listed(items: list[str], limit: int = 10) -> str:
    return ", ".join(items[:limit]) + (f" y {len(items) - limit} más" if len(items) > limit else "")


def run_cost_plan(cfg: Config, book_path: str | Path, backup: bool = True, dry_run: bool = False) -> CostPlanResult:
    """Comando 'procesar-costes': escribe en la hoja '% PV' de la empresa que se depura el importe mensual del plan de
    costes de la hoja PV-depuración, dividido por el BAC de cada fila. cfg.cost_review_calcs: escribe antes los
    cálculos del plan."""
    book_path = Path(book_path)
    progress.stage("libro", "Lectura del libro")
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    if cr.SHEET_PASTE not in wb.sheetnames:
        raise InputError("DEP-012", f"El libro no tiene la hoja {cr.SHEET_PASTE}: procesar-costes usa un libro "
                                    "generado con depuracion_costes = true, con el plan de costes en esa hoja.")
    check_sheets(wb.sheetnames, book.meta, charts=True)
    sheet = str(book.meta.get("depuracion_hoja") or SHEET_PV)
    role, company = _company_role(book, sheet)
    months = book.horizon.months
    warnings = fixable_warning(book, dry_run)

    progress.stage("procesar", "Plan de costes")
    ws_paste = wb[cr.SHEET_PASTE]
    rows, _fmt = cr.read_paste(ws_paste, months)
    rows, _formulas, missing_cells = cr.resolve_formulas(book_path, rows, months)
    calculated = None
    if cfg.cost_review_calcs:
        rows, plan = cr.calc_plan(rows)
        missing_cells = []              # las celdas calculadas toman el valor de los cálculos
        calculated = sum(1 for x in plan if x.kind)
        if not dry_run:
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

    processed: list[tuple] = []                     # (fila, EAC, % de cada mes)
    no_lines, zero, no_bac = [], [], []
    for s in inputs:
        vec = amounts.get(s.row)
        bac = book.bac_role[role][s.row]
        if vec is None:
            no_lines.append(s.code)
        elif all(abs(v) < TOLERANCE for v in vec):
            zero.append(s.code)
        elif not bac or bac <= 0:
            no_bac.append(s.code)
        else:
            processed.append((s, sum(vec), [v / bac for v in vec[:months]]))
    code_of = {s.row: s.code for s in book.structure}
    warnings += _warnings(lines, amounts, months, missing, no_lines, zero, no_bac, sheet, code_of,
                          cfg.cost_review_calcs)

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
    rows_done = {s.row for s, _t, _p in processed}
    assigned = sum(sum(v[:months]) for r, v in amounts.items() if r in rows_done)
    return CostPlanResult(sheet, sum(1 for x in lines if x.client), len(read), sum(x.total for x in read), assigned,
                          [(s.code, s.name, t) for s, t, _p in processed], len(inputs), calculated, warnings,
                          backup_path, dry_run, book)


def _warnings(lines: list[cr.PasteLine], amounts: dict[int, list[float]], months: int, missing: int,
              no_lines: list[str], zero: list[str], no_bac: list[str], sheet: str, code_of: dict[int, str],
              calcs: bool) -> list[Issue]:
    """Observaciones de procesar-costes: líneas sin leer, meses fuera del horizonte, filas sin procesar e importes
    que no se pueden sumar."""
    out: list[Issue] = []
    unread = [x for x in lines if x.state in cr.UNREAD and abs(x.total) >= TOLERANCE]
    if unread:
        items = [f"{_label(x)}: {cr.STATUS_OF[x.state].lower()}" for x in unread]
        out.append(Issue("DEP-014", message=(
            f"{plural(len(unread), 'línea', 'líneas')} de EDT cliente de la hoja {cr.SHEET_PASTE} sin leer, por "
            f"{fmt_money(sum(x.total for x in unread))}: {_listed(items)}. Su importe no forma parte del % "
            "planificado.")))
    repeated = sorted({x.key for x in lines if x.state == cr.STATE_READ and x.repeated})
    if repeated:
        out.append(Issue("DEP-020", message=(
            f"EDT cliente repetidas en la hoja {cr.SHEET_PASTE} ({_listed(repeated)}): se suman sus líneas.")))
    later = {r: sum(v[months:]) for r, v in amounts.items() if abs(sum(v[months:])) >= TOLERANCE}
    if later:
        out.append(Issue("DEP-015", message=(
            f"Hay {fmt_money(sum(later.values()))} del plan de costes en meses posteriores al horizonte del libro, en "
            f"{plural(len(later), 'fila', 'filas')} ({_listed([code_of.get(r, str(r)) for r in sorted(later)])}): "
            "no forman parte del % planificado.")))
    if no_lines:
        out.append(Issue("DEP-016", message=(
            f"Filas de ingreso de la hoja {sheet} sin línea de EDT cliente en el plan de costes ({len(no_lines)}): "
            f"{_listed(no_lines)}. Conservan su % planificado (de MS Project o de un procesado anterior).")))
    if zero:
        hint = "" if calcs else " Si faltan los cálculos del plan, se puede ejecutar procesar-costes --calculos."
        out.append(Issue("DEP-017", message=(
            f"Filas cuya línea de EDT cliente no tiene importes en el plan de costes ({len(zero)}): {_listed(zero)}. "
            f"Conservan su % planificado.{hint}")))
    if no_bac:
        out.append(Issue("DEP-019", message=(
            f"Filas sin BAC en la hoja {sheet} ({len(no_bac)}): {_listed(no_bac)}. El % planificado es el importe del "
            "plan dividido por el BAC: conservan su % planificado hasta que se ingrese el BAC y se ejecute de nuevo "
            "procesar-costes.")))
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
