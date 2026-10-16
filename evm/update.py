"""Comando 'actualizar' (incorpora EV y AC de un mes) y lectura de libros existentes."""
from __future__ import annotations

import shutil
import zipfile
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook.properties import CalcProperties

from evm.config import Config
from evm.methods import ev_fraction, parse_method
from evm.model import ProjectData, Task
from evm.timeline import Horizon
from evm.wbs import Scope, build_tree, fmt_money, iter_subtree, pad_code, prune_inactive, select_scope, tolerance
from evm.workbook import (COL_BAC, COL_METHOD, FIRST_MONTH_COL, LEGACY_SHEET_CURVE, SHEET_AC, SHEET_EV, SHEET_META,
                          SHEET_PV, SHEET_SUMMARY, STATUS_CELL, StructRow, WorkbookFormatError, add_s_curve_chart,
                          col, company_sheet, read_meta, read_structure, save_workbook, write_meta)


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
    # Multi-empresa: valores de la empresa A y B por fila (pv, ev y ac totales son la suma)
    multi: bool = False
    company_a: str = ""
    company_b: str = ""
    pv_a: dict[int, list[float]] = field(default_factory=dict)
    bac_a: dict[int, float] = field(default_factory=dict)
    ev_a: dict[int, list[float | None]] = field(default_factory=dict)
    ev_b: dict[int, list[float | None]] = field(default_factory=dict)
    ac_a: dict[int, list[float | None]] = field(default_factory=dict)
    ac_b: dict[int, list[float | None]] = field(default_factory=dict)

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
                f"La celda {where} tiene una fórmula sin valor numérico calculado. Se debe abrir y guardar el "
                "libro en Excel, o reemplazar la fórmula por un número.")
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
            "Se debe ingresar como número.")
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

    def bac_values(name: str) -> dict[int, float]:
        ws, wv = wb[name], wb_vals[name]
        return {s.row: _num(ws.cell(row=s.row, column=COL_BAC).value, wv.cell(row=s.row, column=COL_BAC).value,
                            f"{name}!D{s.row}", fixable if allow_fix else None, (name, s.row, COL_BAC)) or 0.0
                for s in inputs}

    def zeros(series: dict[int, list[float | None]]) -> dict[int, list[float]]:
        return {r: [v or 0.0 for v in vals] for r, vals in series.items()}

    def combine(x: dict, y: dict) -> dict:
        return {r: [None if a is None and b is None else (a or 0.0) + (b or 0.0) for a, b in zip(x[r], y[r])]
                for r in x}

    ws_pv = wb[SHEET_PV]
    method = {s.row: str(ws_pv.cell(row=s.row, column=COL_METHOD).value or "").strip() for s in inputs}
    status = _as_date(wb[SHEET_SUMMARY][STATUS_CELL].value)
    if meta.get("multiempresa"):
        a, b = str(meta.get("empresa_a") or ""), str(meta.get("empresa_b") or "")
        names = {k: (company_sheet(k, a), company_sheet(k, b)) for k in (SHEET_PV, SHEET_EV, SHEET_AC)}
        for sheet in [n for pair in names.values() for n in pair]:
            if sheet not in wb.sheetnames:
                raise WorkbookFormatError(f"Falta la hoja '{sheet}' del libro multi-empresa.")
        pv_a = zeros(sheet_values(names[SHEET_PV][0], allow_fix))
        pv_b = zeros(sheet_values(names[SHEET_PV][1], allow_fix))
        bac_a, bac_b = bac_values(names[SHEET_PV][0]), bac_values(names[SHEET_PV][1])
        ev_a, ev_b = sheet_values(names[SHEET_EV][0], allow_fix), sheet_values(names[SHEET_EV][1], allow_fix)
        ac_a, ac_b = sheet_values(names[SHEET_AC][0], allow_fix), sheet_values(names[SHEET_AC][1], allow_fix)
        pv = {r: [x + y for x, y in zip(pv_a[r], pv_b[r])] for r in pv_a}
        bac = {r: bac_a[r] + bac_b[r] for r in bac_a}
        return BookData(path, meta, structure, horizon, pv, combine(ev_a, ev_b), combine(ac_a, ac_b), bac, method,
                        status, fixable, True, a, b, pv_a, bac_a, ev_a, ev_b, ac_a, ac_b)
    pv = zeros(sheet_values(SHEET_PV, allow_fix))
    ev = sheet_values(SHEET_EV, allow_fix)
    ac = sheet_values(SHEET_AC, allow_fix)
    return BookData(path, meta, structure, horizon, pv, ev, ac, bac_values(SHEET_PV), method, status, fixable)


def book_meta(path: str | Path) -> dict:
    """Metadatos de un libro generado por este programa (sin leer sus valores)."""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"No existe el libro: {path}")
    try:
        wb = load_workbook(path, read_only=True)
    except (zipfile.BadZipFile, InvalidFileException, OSError) as exc:
        raise WorkbookFormatError(f"No se pudo abrir '{path.name}' como libro de Excel: {exc}") from exc
    try:
        return read_meta(wb)
    finally:
        wb.close()


def apply_book_settings(cfg: Config, meta: dict) -> Config:
    """Configuración multi-empresa con la que se generó el libro (nombres y campos de MS Project).

    La opción ac_desde_project se toma de la configuración actual.
    """
    cfg = replace(cfg)
    cfg.multi_enabled = bool(meta.get("multiempresa"))
    if cfg.multi_enabled:
        cfg.company_a = str(meta.get("empresa_a") or cfg.company_a)
        cfg.company_b = str(meta.get("empresa_b") or cfg.company_b)
        cfg.company_both = str(meta.get("ambas") or cfg.company_both)
        cfg.company_field = str(meta.get("campo_empresa") or cfg.company_field)
        cfg.split_field = str(meta.get("campo_reparto") or cfg.split_field)
        if meta.get("reparto_por_defecto") is not None:
            cfg.split_default = float(meta["reparto_por_defecto"])
    return cfg


def check_only(cfg: Config, only: str | None) -> None:
    """En multi-empresa, --solo ac requiere ac_desde_project (el AC se carga a mano)."""
    if cfg.multi_enabled and only == "ac" and not cfg.ac_from_project:
        raise ValueError("En modo multi-empresa el AC se carga a mano en las hojas de cada empresa. Para cargarlo "
                         "desde MS Project se debe activar ac_desde_project en [multiempresa].")


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
    by_company: dict[str, float] = field(default_factory=dict)   # valor del mes por hoja de empresa


@dataclass
class UpdateResult:
    status_date: date
    month_index: int
    rows: list[RowResult]
    warnings: list[str]
    backup: Path | None
    book: BookData | None = None
    dry_run: bool = False
    ac_loaded: bool = True              # False en multi-empresa con AC cargado a mano


def _acwp(t: Task) -> float:
    return t.acwp if t.acwp is not None else t.actual_cost


def _own_ac(task: Task, children: list[Task]) -> float:
    """Costo real de una tarea resumen que no está en sus filas hijas."""
    return _acwp(task) - sum(_acwp(c) for c in children)


def _own_bcwp(task: Task, children: list[Task]) -> float | None:
    if task.bcwp is None or any(c.bcwp is None for c in children):
        return None
    return task.bcwp - sum(c.bcwp for c in children)


def _match_tasks(project: ProjectData, book: BookData, cfg: Config) -> tuple[dict[int, Task], list[str], Scope]:
    """Asocia cada fila del libro con la tarea de MS Project.

    Mismo proyecto: por UniqueID (si la tarea ya no existe, la fila se omite).
    Proyecto distinto: por código EDT, sin asignar una tarea a dos filas.
    Se usa la planificación con la que se generó el libro.
    """
    warnings: list[str] = []
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)
    if not cfg.include_inactive:
        roots, _, _ = prune_inactive(roots)     # el avance de tareas inactivas no cuenta
    plan_code = str(book.meta.get("planificacion") or "")
    if cfg.plan and cfg.plan.strip().upper() != plan_code.upper():
        warnings.append(f"El libro se generó para la planificación '{plan_code or 'todas'}': se usa esa "
                        f"planificación y no '{cfg.plan}'.")
    scope = select_scope(roots, plan_code)
    by_uid = {t.uid: t for t in tasks}
    by_code: dict[str, Task] = {}
    for t in scope.tasks():
        raw = scope.raw_code(t)
        if raw:
            by_code.setdefault(pad_code(raw, cfg.max_level) if cfg.pad_wbs else raw, t)
    same = str(book.meta.get("proyecto") or "") == project.name
    if not same:
        warnings.append(f"El libro fue generado para el proyecto '{book.meta.get('proyecto')}' y el origen es "
                        f"'{project.name}': las filas se asocian por código EDT. Se recomienda verificar los "
                        "resultados.")
    out: dict[int, Task] = {}
    used: set[tuple[int, bool]] = set()
    for s in book.structure:
        residual = s.key.endswith("-D")
        uid_text = s.key[:-2] if residual else s.key
        if same and uid_text.isdigit():
            task = by_uid.get(int(uid_text))
            if task is None:
                if s.is_input:
                    warnings.append(f"Fila {s.code} '{s.name}': la tarea (UID {uid_text}) ya no existe en MS "
                                    "Project; se omite. Si cambió la EDT, se debe regenerar el libro con "
                                    "--migrar-desde.")
                continue
        else:
            code = s.code[:-2] if residual else s.code
            task = by_code.get(code)
            if task is None or (task.uid, residual) in used:
                if s.is_input:
                    warnings.append(f"Fila {s.code} '{s.name}': no se encontró una tarea con ese código EDT; "
                                    "se omite.")
                continue
        used.add((task.uid, residual))
        out[s.row] = task
    return out, warnings, scope


def _contributors(book: BookData, row: StructRow, task: Task, matches: dict[int, Task]) -> list[Task]:
    """Tareas cuyo costo forma la fila: la rama del paquete, o (fila -D) lo que cuelga de la
    tarea resumen sin pertenecer a sus filas hijas."""
    if row.kind != "residual":
        return list(iter_subtree(task))
    taken = {n.uid for c in _row_children(book, row, matches) for n in iter_subtree(c)}
    return [n for n in iter_subtree(task) if n.uid not in taken]


def _row_children(book: BookData, row: StructRow, matches: dict[int, Task]) -> list[Task]:
    """Tareas de las filas hijas de la tarea resumen de una fila de otros costos (-D)."""
    return [matches[s.row] for s in book.structure
            if s.parent_row == row.parent_row and s.kind != "residual" and s.row in matches]


def _coverage_warnings(scope: Scope, book: BookData, matches: dict[int, Task], cfg: Config) -> list[str]:
    """Costo real de tareas de la planificación que no queda en ninguna fila del libro."""
    covered: set[int] = set()
    for s in book.input_rows:
        task = matches.get(s.row)
        if task is not None:
            # Una fila -D cubre todo lo que cuelga de su tarea resumen y no tiene fila propia.
            covered.update(n.uid for n in iter_subtree(task))
    missing = []
    for t in scope.tasks():
        if t.uid in covered:
            continue
        own = _acwp(t) - sum(_acwp(c) for c in t.children)
        if own > tolerance(own):
            missing.append((t, own))
    warnings = []
    for t, amount in missing[:15]:
        code = scope.raw_code(t) or "(sin código EDT)"
        warnings.append(f"{code} '{t.name}': costo real {fmt_money(amount)} sin fila en el libro; no se incluye en "
                        "el AC. Para incluirlo, la tarea debe tener código EDT y el libro se debe regenerar con "
                        "--migrar-desde.")
    if len(missing) > 15:
        warnings.append(f"... y {len(missing) - 15} tareas más con costo real sin fila en el libro.")
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
    cfg = apply_book_settings(cfg, book.meta)
    warnings: list[str] = list(project.warnings)
    multi = book.multi
    load_ac = (not multi) or cfg.ac_from_project
    check_only(cfg, only)

    if status is None:
        if project.status_date is None:
            raise ValueError("El proyecto no tiene fecha de estado: se debe indicar la fecha de corte con "
                             "--fecha-corte AAAA-MM-DD o definir Proyecto > Fecha de estado en MS Project.")
        status = project.status_date.date()
    idx = book.horizon.index(status)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {status:%d/%m/%Y} está fuera del horizonte del libro "
                         f"({book.horizon.start:%m/%Y} a {book.horizon.end:%m/%Y}).")
    if project.status_date is None and project.backend == "xml":
        warnings.append("El XML no tiene fecha de estado: MS Project calculó el costo real (CRTR) y el CPTR a la "
                        "fecha en que se guardó el archivo. Se debe definir Proyecto > Fecha de estado antes de "
                        "exportar.")
    if status != _month_end(status):
        warnings.append("La fecha de corte no es el último día del mes: el PV acumulado incluye el mes completo.")
    if project.status_date and project.status_date.date() != status:
        warnings.append(f"La fecha de estado de MS Project ({project.status_date:%d/%m/%Y}) difiere de la fecha de "
                        f"corte ({status:%d/%m/%Y}). El costo real (CRTR) y el CPTR corresponden a la fecha de "
                        "estado del proyecto.")
    if book.status_date and status < book.status_date:
        warnings.append(f"La fecha de corte ({status:%d/%m/%Y}) es anterior a la última cargada en el libro "
                        f"({book.status_date:%d/%m/%Y}).")

    matches, w, scope = _match_tasks(project, book, cfg)
    warnings += w
    if only != "ev" and load_ac:
        warnings += _coverage_warnings(scope, book, matches, cfg)

    shares: dict[int, float] = {}
    if multi:
        from evm.companies import ac_by_company, ev_by_company, resolve_shares
        shares, _ = resolve_shares(scope.roots, cfg)
        sheets = {(k, r): company_sheet(k, book.company_a if r == "A" else book.company_b)
                  for k in (SHEET_EV, SHEET_AC) for r in ("A", "B")}

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
        siblings = _row_children(book, s, matches) if residual else []
        bcwp = _own_bcwp(task, siblings) if residual else task.bcwp
        bcwp_base = None
        if task.ev_baseline_cost is not None:       # CPTR calculado con otra línea base
            bcwp_base = task.ev_baseline_cost
            if residual:
                bcwp_base -= sum(c.ev_baseline_cost or 0.0 for c in siblings)
        frac, note = ev_fraction(method, task, (pv_cum / bac) if bac else 0.0, bac=bac, bcwp=bcwp,
                                 bcwp_base=bcwp_base)

        # (tipo, hoja, valor acumulado, serie del libro)
        targets: list[tuple[str, str, float, list]] = []
        if multi:
            contributors = _contributors(book, s, task, matches)
            bac_a = book.bac_a[s.row]
            pv_cum_a = sum(book.pv_a[s.row][: idx + 1])
            ev_a, ev_b = ev_by_company(method, contributors, shares, (bac_a, bac - bac_a),
                                       (pv_cum_a, pv_cum - pv_cum_a), frac, task.ev_baseline_cost is not None)
            targets += [("ev", sheets[(SHEET_EV, "A")], round(ev_a, 2), book.ev_a[s.row]),
                        ("ev", sheets[(SHEET_EV, "B")], round(ev_b, 2), book.ev_b[s.row])]
            if load_ac:
                ac_a, ac_b = ac_by_company(contributors, shares)
                targets += [("ac", sheets[(SHEET_AC, "A")], round(ac_a, 2), book.ac_a[s.row]),
                            ("ac", sheets[(SHEET_AC, "B")], round(ac_b, 2), book.ac_b[s.row])]
        else:
            targets += [("ev", SHEET_EV, round(bac * frac, 2), book.ev[s.row]),
                        ("ac", SHEET_AC, round(_own_ac(task, siblings) if residual else _acwp(task), 2),
                         book.ac[s.row])]

        res = RowResult(s.code, s.name, method_text, bac, None, None, None, None, note)
        for kind, sheet, cum_value, series in targets:
            if only and only != kind:
                continue
            later = [v for v in series[idx + 1:] if v]
            if later:
                later_filled.append(f"{sheet} {s.code}")
            prev = sum(v or 0.0 for v in series[:idx])
            month_value = round(cum_value - prev, 2)
            if kind == "ev":
                res.ev_cum = (res.ev_cum or 0.0) + cum_value
                res.ev_month = (res.ev_month or 0.0) + month_value
            else:
                res.ac_cum = (res.ac_cum or 0.0) + cum_value
                res.ac_month = (res.ac_month or 0.0) + month_value
                ac_total += cum_value
            if multi:
                res.by_company[sheet] = month_value
            if month_value < -tolerance(cum_value):
                res.note = (res.note + "; " if res.note else "") + f"{sheet} del mes negativo (corrección)"
            writes.append((sheet, s.row, FIRST_MONTH_COL + idx, month_value))
        rows.append(res)

    if later_filled and not force:
        raise ValueError(
            "Hay valores cargados en meses posteriores a la fecha de corte (" + ", ".join(later_filled[:10]) +
            ("..." if len(later_filled) > 10 else "") + "). Se deben borrar esos valores o usar --forzar.")

    reference = _acwp(scope.plan) if scope.plan is not None else project.summary_acwp
    if only != "ev" and load_ac and reference is not None:
        diff = reference - ac_total
        if abs(diff) > tolerance(reference):
            whose = f"de la planificación '{scope.prefix}'" if scope.plan is not None else "del proyecto"
            warnings.append(f"El AC cargado ({fmt_money(ac_total)}) difiere del costo real {whose} en MS Project "
                            f"({fmt_money(reference)}) en {fmt_money(diff)}. Se recomienda revisar las advertencias "
                            "anteriores (tareas sin fila, costos en tareas resumen).")

    if book.fixable:
        cells = ", ".join(f"{sh}!{col(c)}{r}" for sh, r, c, _v, _why in book.fixable[:8])
        more = f" y {len(book.fixable) - 8} más" if len(book.fixable) > 8 else ""
        warnings.append(f"Se reemplazaron por su valor numérico celdas de ingreso con texto o fórmulas: {cells}{more}.")

    # Valores en memoria (también en simulación, para mostrar los indicadores resultantes)
    for sheet, r, c, v in writes:
        i = c - FIRST_MONTH_COL
        if not multi:
            (book.ev if sheet == SHEET_EV else book.ac)[r][i] = v
            continue
        target = {sheets[(SHEET_EV, "A")]: book.ev_a, sheets[(SHEET_EV, "B")]: book.ev_b,
                  sheets[(SHEET_AC, "A")]: book.ac_a, sheets[(SHEET_AC, "B")]: book.ac_b}[sheet]
        target[r][i] = v
        total = book.ev if sheet in (sheets[(SHEET_EV, "A")], sheets[(SHEET_EV, "B")]) else book.ac
        pair = (book.ev_a, book.ev_b) if total is book.ev else (book.ac_a, book.ac_b)
        x, y = pair[0][r][i], pair[1][r][i]
        total[r][i] = None if x is None and y is None else (x or 0.0) + (y or 0.0)

    backup_path = None
    if not dry_run:
        for sheet, r, c, v, _why in book.fixable:
            wb[sheet].cell(row=r, column=c, value=v)
        for sheet, r, c, v in writes:
            wb[sheet].cell(row=r, column=c, value=v)
        wb[SHEET_SUMMARY][STATUS_CELL].value = status
        meta = dict(book.meta)
        meta.pop("meses_grafico", None)
        meta["ultima_actualizacion"] = datetime.now().replace(microsecond=0)
        write_meta(wb[SHEET_META], meta)
        _refresh_charts(wb, book, idx, meta)
        wb.calculation = CalcProperties(fullCalcOnLoad=True)
        if backup:
            backup_path = _backup_path(book_path)
            shutil.copy2(book_path, backup_path)
        save_workbook(wb, book_path)
    book.status_date = status
    return UpdateResult(status, idx, rows, warnings, backup_path, book, dry_run, ac_loaded=load_ac)


def _refresh_charts(wb, book: BookData, idx: int, meta: dict) -> None:
    """Vuelve a crear los gráficos (openpyxl no los conserva) y fija el mes de reporte en el mes de corte."""
    from evm import dashboard
    currency = str(meta.get("moneda") or "")
    if dashboard.SHEET_DASHBOARD in wb.sheetnames:
        wb[dashboard.SHEET_DASHBOARD][dashboard.REPORT_CELL].value = book.horizon.month_start(idx)
        pv, _ev, _ac, bac = book_totals(book)
        loaded = [i for i in range(idx + 1)
                  if any(series[r][i] is not None for series in (book.ev, book.ac) for r in series)]
        packages = dashboard.chart_packages((s.row, s.code, s.level, s.kind) for s in book.structure)
        dashboard.add_dashboard_charts(wb, packages, book.horizon.months, dashboard.plan_end_month(pv), idx + 1,
                                       bac, currency, idx + 1 - loaded[0] if loaded else 0)
    elif LEGACY_SHEET_CURVE in wb.sheetnames:
        add_s_curve_chart(wb[LEGACY_SHEET_CURVE], book.horizon.months, str(meta.get("proyecto") or ""), currency)


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


def indicators_for(book: BookData, status: date):
    """Indicadores del proyecto a la fecha de corte (mismos cálculos que la hoja Resumen)."""
    from evm.indicators import compute
    pv, ev, ac, bac = book_totals(book)
    idx = book.horizon.index(status)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {status:%d/%m/%Y} está fuera del horizonte del libro.")
    return compute(pv, ev, ac, bac, idx)
