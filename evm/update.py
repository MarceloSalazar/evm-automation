"""Comando 'actualizar' (incorpora EV y AC de un mes) y lectura de libros existentes."""
from __future__ import annotations

import shutil
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.properties import CalcProperties

from evm.config import Config
from evm.methods import ev_fraction, parse_method
from evm.model import ProjectData, Task
from evm.timeline import Horizon
from evm.wbs import build_tree, fmt_money, iter_subtree, pad_code, prune_inactive, tolerance
from evm.workbook import (COL_BAC, COL_METHOD, FIRST_MONTH_COL, SHEET_AC, SHEET_CURVE, SHEET_EV, SHEET_META,
                          SHEET_PV, SHEET_SUMMARY, STATUS_CELL, StructRow, WorkbookFormatError, add_s_curve_chart,
                          col, read_meta, read_structure, save_workbook, write_meta)


# ----------------------------------------------------------------------------
# Lectura de libros existentes
# ----------------------------------------------------------------------------
@dataclass
class BookData:
    path: Path
    meta: dict
    structure: list[StructRow]
    horizon: Horizon
    pv: dict[int, list[float]]           # fila -> valores mensuales (solo filas de ingreso)
    ev: dict[int, list[float | None]]
    ac: dict[int, list[float | None]]
    bac: dict[int, float]
    method: dict[int, str]
    status_date: date | None
    # Celdas de ingreso de EV/AC que no son números: (hoja, fila, columna, valor numérico, motivo)
    fixable: list[tuple[str, int, int, float, str]] = field(default_factory=list)

    @property
    def input_rows(self) -> list[StructRow]:
        return [s for s in self.structure if s.is_input]


def _num(value, cached, where: str, fixable: list | None = None, key=None) -> float | None:
    """Valor numérico de una celda.

    Fórmula: usa el valor calculado que guardó Excel. Texto con un número: Excel no
    lo suma, así que se rechaza (o, si se indica `fixable`, se anota para convertirlo).
    """
    if isinstance(value, str) and value.startswith("="):
        if cached is None or isinstance(cached, str):
            raise WorkbookFormatError(
                f"La celda {where} tiene una fórmula sin valor numérico calculado. Abra y guarde el libro en "
                "Excel, o reemplace la fórmula por un número.")
        if fixable is not None and key is not None and not isinstance(cached, bool):
            fixable.append((*key, float(cached), "fórmula reemplazada por su valor"))
        value = cached
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise WorkbookFormatError(f"Valor no numérico en {where}: {value!r}")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        number = float(text.replace(" ", "").replace(",", "."))
    except ValueError as exc:
        raise WorkbookFormatError(f"Valor no numérico en {where}: {value!r}") from exc
    if fixable is None or key is None:
        raise WorkbookFormatError(
            f"La celda {where} contiene el número {value!r} guardado como texto: Excel no lo suma. "
            "Ingréselo como número.")
    fixable.append((*key, number, "número guardado como texto convertido a número"))
    return number


def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def read_book(path: str | Path, wb=None, allow_fix: bool = False) -> BookData:
    """Lee un libro generado por este programa.

    allow_fix: en lugar de rechazar celdas de ingreso con texto o fórmulas, las
    anota en BookData.fixable para que 'actualizar' las convierta en números.
    """
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"No existe el libro: {path}")
    if path.suffix.lower() != ".xlsx":
        raise WorkbookFormatError(f"'{path.name}' no es un libro .xlsx generado por este programa.")
    try:
        wb = wb or load_workbook(path)
        wb_vals = load_workbook(path, data_only=True)
    except (zipfile.BadZipFile, InvalidFileException, OSError) as exc:
        raise WorkbookFormatError(f"No se pudo abrir '{path.name}' como libro de Excel: {exc}") from exc
    meta = read_meta(wb)
    structure = read_structure(wb)
    horizon = Horizon(meta["horizonte_inicio"], meta["meses"])
    months = horizon.months
    inputs = [s for s in structure if s.is_input]
    fixable: list = []

    def sheet_values(name: str, fix: bool) -> dict[int, list[float | None]]:
        ws, wv = wb[name], wb_vals[name]
        out = {}
        for s in inputs:
            vals = []
            for i in range(months):
                c = FIRST_MONTH_COL + i
                vals.append(_num(ws.cell(row=s.row, column=c).value, wv.cell(row=s.row, column=c).value,
                                 f"{name}!{col(c)}{s.row}", fixable if fix else None, (name, s.row, c)))
            out[s.row] = vals
        return out

    pv_raw = sheet_values(SHEET_PV, allow_fix)
    pv = {r: [v or 0.0 for v in vals] for r, vals in pv_raw.items()}
    ev = sheet_values(SHEET_EV, allow_fix)
    ac = sheet_values(SHEET_AC, allow_fix)
    ws_pv, wv_pv = wb[SHEET_PV], wb_vals[SHEET_PV]
    bac = {}
    method = {}
    for s in inputs:
        bac[s.row] = _num(ws_pv.cell(row=s.row, column=COL_BAC).value,
                          wv_pv.cell(row=s.row, column=COL_BAC).value, f"PV!D{s.row}",
                          fixable if allow_fix else None, (SHEET_PV, s.row, COL_BAC)) or 0.0
        method[s.row] = str(ws_pv.cell(row=s.row, column=COL_METHOD).value or "").strip()
    status = _as_date(wb[SHEET_SUMMARY][STATUS_CELL].value)
    return BookData(path, meta, structure, horizon, pv, ev, ac, bac, method, status, fixable)


# ----------------------------------------------------------------------------
# Actualización mensual
# ----------------------------------------------------------------------------
@dataclass
class RowResult:
    code: str
    name: str
    method: str
    bac: float
    ev_cum: float | None
    ac_cum: float | None
    ev_month: float | None
    ac_month: float | None
    note: str = ""


@dataclass
class UpdateResult:
    status_date: date
    month_index: int
    rows: list[RowResult]
    warnings: list[str]
    backup: Path | None
    book: BookData | None = None
    dry_run: bool = False


def _acwp(t: Task) -> float:
    return t.acwp if t.acwp is not None else t.actual_cost


def _own_ac(task: Task) -> float:
    """Costo real propio de una tarea resumen (sin sus subtareas)."""
    return _acwp(task) - sum(_acwp(c) for c in task.children)


def _own_bcwp(task: Task) -> float | None:
    if task.bcwp is None or any(c.bcwp is None for c in task.children):
        return None
    return task.bcwp - sum(c.bcwp for c in task.children)


def _match_tasks(project: ProjectData, book: BookData, cfg: Config) -> tuple[dict[int, Task], list[str]]:
    """Asocia cada fila de ingreso del libro con la tarea de MS Project.

    Mismo proyecto: por UniqueID (si la tarea ya no existe, la fila se omite).
    Proyecto distinto: por código EDT, sin asignar una tarea a dos filas.
    """
    warnings: list[str] = []
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)
    if not cfg.include_inactive:
        prune_inactive(roots)       # el avance de tareas inactivas no cuenta
    by_uid = {t.uid: t for t in tasks}
    by_code: dict[str, Task] = {}
    for t in tasks:
        code = pad_code(t.wbs, cfg.max_level) if cfg.pad_wbs else (t.wbs or "").strip()
        by_code.setdefault(code, t)
    same = str(book.meta.get("proyecto") or "") == project.name
    if not same:
        warnings.append(f"El libro fue generado para el proyecto '{book.meta.get('proyecto')}' y el origen es "
                        f"'{project.name}': las filas se asocian por código EDT. Verifique los resultados.")
    out: dict[int, Task] = {}
    used: set[tuple[int, bool]] = set()
    for s in book.input_rows:
        residual = s.key.endswith("-D")
        uid_text = s.key[:-2] if residual else s.key
        if same and uid_text.isdigit():
            task = by_uid.get(int(uid_text))
            if task is None:
                warnings.append(f"Fila {s.code} '{s.name}': la tarea (UID {uid_text}) ya no existe en MS Project; "
                                "se omite. Si cambió la EDT, regenere el libro con --migrar-desde.")
                continue
        else:
            code = s.code[:-2] if residual else s.code
            task = by_code.get(code)
            if task is None or (task.uid, residual) in used:
                warnings.append(f"Fila {s.code} '{s.name}': no se encontró una tarea con ese código EDT; se omite.")
                continue
        used.add((task.uid, residual))
        out[s.row] = task
    return out, warnings


def _coverage_warnings(project: ProjectData, book: BookData, matches: dict[int, Task], cfg: Config) -> list[str]:
    """Tareas con costo que no quedaron representadas en ninguna fila del libro."""
    warnings: list[str] = []
    covered: set[int] = set()
    ancestors: set[int] = set()
    residual_uids: set[int] = set()
    rows = {s.row: s for s in book.input_rows}
    for r, task in matches.items():
        if rows[r].kind == "residual":
            residual_uids.add(task.uid)
            continue
        covered.update(n.uid for n in iter_subtree(task))
        p = task.parent
        while p is not None:
            ancestors.add(p.uid)
            p = p.parent
    missing = []
    for t in project.tasks:
        if t.outline_level < 1 or t.uid in covered or (not cfg.include_inactive and not t.active):
            continue
        if t.uid in ancestors:
            own = _own_ac(t) if t.children else _acwp(t)
            if own > tolerance(own) and t.uid not in residual_uids:
                missing.append((t, own, "costo real asignado directamente a la tarea resumen"))
            continue
        parent_covered = t.parent is not None and t.parent.uid in covered
        if not parent_covered and (_acwp(t) > tolerance(_acwp(t)) or t.baseline_cost or t.cost):
            if t.parent is None or t.parent.uid in ancestors:
                missing.append((t, _acwp(t), "tarea sin fila en el libro"))
    for t, amount, why in missing[:15]:
        warnings.append(f"{t.wbs} '{t.name}': {why} (costo real {fmt_money(amount)}); no se incluye en el AC. "
                        "Regenere el libro con --migrar-desde para incorporarla.")
    if len(missing) > 15:
        warnings.append(f"... y {len(missing) - 15} tareas más sin fila en el libro.")
    return warnings


def _month_end(d: date) -> date:
    return date(d.year + (d.month == 12), d.month % 12 + 1, 1) - timedelta(days=1)


def _backup_path(book_path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    path = book_path.with_name(f"{book_path.stem}.respaldo-{stamp}{book_path.suffix}")
    n = 2
    while path.exists():
        path = book_path.with_name(f"{book_path.stem}.respaldo-{stamp}-{n}{book_path.suffix}")
        n += 1
    return path


def run_update(project: ProjectData, cfg: Config, book_path: str | Path, status: date | None = None,
               only: str | None = None, backup: bool = True, force: bool = False,
               dry_run: bool = False) -> UpdateResult:
    book_path = Path(book_path)
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    warnings: list[str] = list(project.warnings)

    if status is None:
        if project.status_date is None:
            raise ValueError("El proyecto no tiene fecha de estado: indique la fecha de corte con --fecha-corte "
                             "AAAA-MM-DD, o defina Proyecto > Fecha de estado en MS Project.")
        status = project.status_date.date()
    idx = book.horizon.index(status)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {status:%d/%m/%Y} está fuera del horizonte del libro "
                         f"({book.horizon.start:%m/%Y} a {book.horizon.end:%m/%Y}).")
    if project.status_date is None and project.backend == "xml":
        warnings.append("El XML no tiene fecha de estado: MS Project calculó el costo real (CRTR) y el CPTR a la "
                        "fecha en que se guardó el archivo. Defina Proyecto > Fecha de estado antes de exportar.")
    if status != _month_end(status):
        warnings.append("La fecha de corte no es el último día del mes: el PV acumulado incluye el mes completo.")
    if project.status_date and project.status_date.date() != status:
        warnings.append(f"La fecha de estado de MS Project ({project.status_date:%d/%m/%Y}) difiere de la fecha de "
                        f"corte ({status:%d/%m/%Y}). El costo real (CRTR) y el CPTR corresponden a la fecha de "
                        "estado del proyecto.")
    if book.status_date and status < book.status_date:
        warnings.append(f"La fecha de corte ({status:%d/%m/%Y}) es anterior a la última cargada en el libro "
                        f"({book.status_date:%d/%m/%Y}).")

    matches, w = _match_tasks(project, book, cfg)
    warnings += w
    if only != "ev":
        warnings += _coverage_warnings(project, book, matches, cfg)

    rows: list[RowResult] = []
    later_filled: list[str] = []
    writes: list[tuple[str, int, int, float]] = []   # (hoja, fila, columna, valor)
    ac_total = 0.0
    for s in book.input_rows:
        task = matches.get(s.row)
        if task is None:
            continue
        bac = book.bac[s.row]
        method_text = book.method[s.row]
        method = parse_method(method_text)
        pv_cum = sum(book.pv[s.row][: idx + 1])
        residual = s.kind == "residual"
        bcwp = _own_bcwp(task) if residual else task.bcwp
        bcwp_base = None
        if task.ev_baseline_cost is not None:       # CPTR calculado con otra línea base
            bcwp_base = task.ev_baseline_cost
            if residual:
                bcwp_base -= sum(c.ev_baseline_cost or 0.0 for c in task.children)
        frac, note = ev_fraction(method, task, (pv_cum / bac) if bac else 0.0, bac=bac, bcwp=bcwp,
                                 bcwp_base=bcwp_base)
        ev_cum = round(bac * frac, 2)
        ac_cum = round(_own_ac(task) if residual else _acwp(task), 2)
        ac_total += ac_cum

        res = RowResult(s.code, s.name, method_text, bac, None, None, None, None, note)
        for kind, cum_value, series in (("ev", ev_cum, book.ev[s.row]), ("ac", ac_cum, book.ac[s.row])):
            if only and only != kind:
                continue
            later = [v for v in series[idx + 1:] if v]
            if later:
                later_filled.append(f"{kind.upper()} {s.code}")
            prev = sum(v or 0.0 for v in series[:idx])
            month_value = round(cum_value - prev, 2)
            if kind == "ev":
                res.ev_cum, res.ev_month = cum_value, month_value
            else:
                res.ac_cum, res.ac_month = cum_value, month_value
            if month_value < -tolerance(cum_value):
                res.note = (res.note + "; " if res.note else "") + f"{kind.upper()} del mes negativo (corrección)"
            sheet = SHEET_EV if kind == "ev" else SHEET_AC
            writes.append((sheet, s.row, FIRST_MONTH_COL + idx, month_value))
        rows.append(res)

    if later_filled and not force:
        raise ValueError(
            "Hay valores cargados en meses posteriores a la fecha de corte (" + ", ".join(later_filled[:10]) +
            ("..." if len(later_filled) > 10 else "") + "). Bórrelos o use --forzar.")

    if only != "ev" and project.summary_acwp is not None:
        diff = project.summary_acwp - ac_total
        if abs(diff) > tolerance(project.summary_acwp):
            warnings.append(f"El AC cargado ({fmt_money(ac_total)}) difiere del costo real del proyecto en MS Project "
                            f"({fmt_money(project.summary_acwp)}) en {fmt_money(diff)}. Revise las advertencias "
                            "anteriores (tareas sin fila, costos en tareas resumen).")

    if book.fixable:
        cells = ", ".join(f"{sh}!{col(c)}{r}" for sh, r, c, _v, _why in book.fixable[:8])
        more = f" y {len(book.fixable) - 8} más" if len(book.fixable) > 8 else ""
        warnings.append(f"Se reemplazaron por su valor numérico celdas de ingreso con texto o fórmulas: {cells}{more}.")

    # Valores en memoria (también en simulación, para mostrar los indicadores resultantes)
    for sheet, r, c, v in writes:
        target = book.ev if sheet == SHEET_EV else book.ac
        target[r][c - FIRST_MONTH_COL] = v

    backup_path = None
    if not dry_run:
        for sheet, r, c, v, _why in book.fixable:
            wb[sheet].cell(row=r, column=c, value=v)
        for sheet, r, c, v in writes:
            wb[sheet].cell(row=r, column=c, value=v)
        wb[SHEET_SUMMARY][STATUS_CELL].value = status
        meta = dict(book.meta)
        chart_months = max(int(meta.get("meses_grafico") or 12), min(book.horizon.months, idx + 1 + 6))
        meta["meses_grafico"] = chart_months
        meta["ultima_actualizacion"] = datetime.now().replace(microsecond=0)
        write_meta(wb[SHEET_META], meta)
        add_s_curve_chart(wb[SHEET_CURVE], chart_months, str(meta.get("proyecto") or ""))
        wb.calculation = CalcProperties(fullCalcOnLoad=True)
        if backup:
            backup_path = _backup_path(book_path)
            shutil.copy2(book_path, backup_path)
        save_workbook(wb, book_path)
    book.status_date = status
    return UpdateResult(status, idx, rows, warnings, backup_path, book, dry_run)


def book_totals(book: BookData) -> tuple[list[float], list[float], list[float], float]:
    """Totales mensuales de PV, EV y AC (sumando las filas de ingreso) y BAC total."""
    n = book.horizon.months
    pv = [0.0] * n
    ev = [0.0] * n
    ac = [0.0] * n
    for s in book.input_rows:
        for i in range(n):
            pv[i] += book.pv[s.row][i]
            ev[i] += book.ev[s.row][i] or 0.0
            ac[i] += book.ac[s.row][i] or 0.0
    bac = sum(book.bac[s.row] for s in book.input_rows)
    return pv, ev, ac, bac


def first_loaded_month(series_by_row: dict[int, list[float | None]]) -> int | None:
    """Primer mes (índice base 0) con algún valor cargado (aunque sea 0)."""
    firsts = [next((i for i, v in enumerate(vals) if v is not None), None) for vals in series_by_row.values()]
    firsts = [f for f in firsts if f is not None]
    return min(firsts) if firsts else None
