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

from evm import analysis
from evm.config import MODE_COST, MODE_PERCENT, Config, ConfigError, parse_mode
from evm.methods import ev_fraction, parse_method
from evm.model import ProjectData, Task
from evm.timeline import Horizon
from evm.wbs import (KIND_RESIDUAL, KIND_SUMMARY, Scope, build_tree, code_level, fmt_money, iter_subtree, pad_code,
                     plural, prune_inactive, select_scope, tolerance)
from evm.workbook import (COL_BAC, COL_CODE, COL_METHOD, FIRST_MONTH_COL, LEGACY_SHEET_CURVE, PCT_PREFIX, SHEET_AC,
                          SHEET_EV, SHEET_META, SHEET_PV, SHEET_SUMMARY, STATUS_CELL, StructRow, WorkbookFormatError,
                          add_s_curve_chart, book_sources, clean_text, col, company_sheet, pct_sheet, read_meta,
                          read_structure, save_workbook, source_meta, split_key, write_meta)


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
    # Celdas que deben tener fórmula (subtotales, fila TOTAL, suma de empresas) y tienen un valor
    formula_gaps: list[str] = field(default_factory=list)
    # Modo porcentaje: BAC ingresado, % planificado y % ganado de cada mes por empresa
    # (clave "" en un libro de una sola empresa; "A" y "B" en multi-empresa)
    mode: str = MODE_COST
    bac_role: dict[str, dict[int, float]] = field(default_factory=dict)
    pct_pv: dict[str, dict[int, list[float]]] = field(default_factory=dict)
    pct_ev: dict[str, dict[int, list[float | None]]] = field(default_factory=dict)

    @property
    def input_rows(self) -> list[StructRow]:
        return [s for s in self.structure if s.is_input]

    @property
    def percent(self) -> bool:
        return self.mode == MODE_PERCENT

    def roles(self) -> list[tuple[str, str]]:
        """(clave, empresa) de cada serie: [("", "")] o [("A", empresa A), ("B", empresa B)]."""
        if self.multi:
            return [("A", self.company_a), ("B", self.company_b)]
        return [("", "")]


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


def _company_names(meta: dict) -> tuple[str, str]:
    return str(meta.get("empresa_a") or ""), str(meta.get("empresa_b") or "")


def book_mode(meta: dict) -> str:
    """Modo del libro: 'porcentaje' o 'costo' (los libros sin modo se generaron en el modo costo)."""
    try:
        return parse_mode(str(meta.get("modo") or MODE_COST))
    except ConfigError:
        raise WorkbookFormatError(f"El dato 'modo' de la hoja oculta _meta no es válido ({meta.get('modo')!r}). Esa "
                                  "hoja no se debe modificar: se debe recuperar el libro desde un respaldo.") from None


def company_sheets(meta: dict) -> list[str]:
    """Hojas de cada empresa en multi-empresa (PV-EmpresaA... y, en el modo porcentaje, '% PV-EmpresaA'...)."""
    if not meta.get("multiempresa"):
        return []
    a, b = _company_names(meta)
    names = [company_sheet(k, c) for k in (SHEET_PV, SHEET_EV, SHEET_AC) for c in (a, b)]
    if book_mode(meta) == MODE_PERCENT:
        names += [pct_sheet(k, c) for k in (SHEET_PV, SHEET_EV) for c in (a, b)]
    return names


def data_sheets(meta: dict) -> list[str]:
    """Hojas mensuales del libro: PV, EV y AC, en multi-empresa las de cada empresa y, en el modo porcentaje, las
    hojas '% PV' y '% EV'."""
    names = [SHEET_PV, SHEET_EV, SHEET_AC]
    if meta.get("multiempresa"):
        return names + company_sheets(meta)
    if book_mode(meta) == MODE_PERCENT:
        names += [pct_sheet(SHEET_PV), pct_sheet(SHEET_EV)]
    return names


def check_sheets(sheetnames: list[str], meta: dict, charts: bool = False) -> None:
    """Verifica que estén las hojas que el programa lee o escribe (no se deben eliminar ni renombrar).

    charts: también la hoja de datos de los gráficos del Cuadro de Mando (la usa 'actualizar').
    """
    from evm.dashboard import SHEET_CHARTDATA, SHEET_DASHBOARD
    names = [SHEET_SUMMARY] + data_sheets(meta)
    if charts and SHEET_DASHBOARD in sheetnames:
        names.append(SHEET_CHARTDATA)
    companies = set(company_sheets(meta))
    for name in names:
        if name in sheetnames:
            continue
        if name in companies:
            raise WorkbookFormatError(f"Falta la hoja '{name}' del libro multi-empresa.")
        raise WorkbookFormatError(
            f"Falta la hoja '{name}' del libro. Las hojas del libro no se deben eliminar ni cambiar de nombre: se "
            "debe restaurar su nombre original o recuperar el libro desde un respaldo.")


def _check_layout(wb, structure: list[StructRow], meta: dict) -> None:
    """Las filas de las hojas mensuales deben estar donde las ubicó 'generar' ('actualizar' escribe por
    posición): se compara el código EDT de la columna A con la hoja _estructura."""
    expected = [(s.row, s.code) for s in structure] + [(meta["fila_total"], "TOTAL")]
    for name in data_sheets(meta):
        ws = wb[name]
        for r, code in expected:
            value = ws.cell(row=r, column=COL_CODE).value
            text = "" if value is None else str(value).strip()
            if text != code.strip():
                raise WorkbookFormatError(
                    f"La hoja '{name}' no tiene las filas con las que se generó el libro: la fila {r} debe tener el "
                    f"código EDT '{code}' y tiene '{text}'. Como 'actualizar' ubica cada fila por su posición, se "
                    "deben deshacer las filas insertadas, eliminadas u ordenadas en esa hoja (o recuperar el libro "
                    "desde un respaldo). Para cambiar la EDT se debe modificar MS Project y regenerar el libro con "
                    "--migrar-desde.")


def _formula_gaps(wb, structure: list[StructRow], meta: dict) -> list[str]:
    """Celdas que deben tener fórmula y tienen un valor o están vacías: subtotales y fila TOTAL de las hojas
    mensuales (meses y, en PV, el BAC) y, en multi-empresa, todas las filas de las hojas PV, EV y AC (suma de
    las dos empresas)."""
    total = meta["fila_total"]
    months = list(range(FIRST_MONTH_COL, FIRST_MONTH_COL + meta["meses"]))
    subtotals = [s.row for s in structure if s.kind == KIND_SUMMARY] + [total]
    every = [s.row for s in structure] + [total]
    multi = bool(meta.get("multiempresa"))
    percent = book_mode(meta) == MODE_PERCENT
    out = []
    for name in data_sheets(meta):
        is_total_sheet = multi and name in (SHEET_PV, SHEET_EV, SHEET_AC)
        is_pv = name == SHEET_PV or name.startswith(SHEET_PV + "-")
        is_ev = name == SHEET_EV or name.startswith(SHEET_EV + "-")
        if is_total_sheet:
            cells = [(r, c) for r in every for c in months + ([COL_BAC] if is_pv else [])]
        elif percent and (is_pv or is_ev):
            # Modo porcentaje: los meses son BAC × % (fórmula) en todas las filas; el BAC se ingresa
            cells = [(r, c) for r in every for c in months] + ([(r, COL_BAC) for r in subtotals] if is_pv else [])
        else:
            cells = [(r, c) for r in subtotals for c in months + ([COL_BAC] if is_pv else [])]
        ws = wb[name]
        for r, c in cells:
            value = ws.cell(row=r, column=c).value
            if not (isinstance(value, str) and value.startswith("=")):
                out.append(f"{name}!{col(c)}{r}")
    return out


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
    check_sheets(wb.sheetnames, meta)
    _check_layout(wb, structure, meta)
    gaps = _formula_gaps(wb, structure, meta)
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
    mode = book_mode(meta)
    if mode == MODE_PERCENT:
        # PV y EV = BAC ingresado × % de las hojas '% PV' y '% EV'
        multi = bool(meta.get("multiempresa"))
        a, b = _company_names(meta) if multi else ("", "")
        roles = [("A", a), ("B", b)] if multi else [("", "")]
        bac_role, pct_pv, pct_ev, pv_r, ev_r, ac_r = {}, {}, {}, {}, {}, {}
        for role, company in roles:
            bac_role[role] = bac_values(company_sheet(SHEET_PV, company) if company else SHEET_PV)
            pct_pv[role] = zeros(sheet_values(pct_sheet(SHEET_PV, company), allow_fix))
            pct_ev[role] = sheet_values(pct_sheet(SHEET_EV, company), allow_fix)
            bacs = bac_role[role]
            pv_r[role] = {r: [bacs[r] * v for v in vals] for r, vals in pct_pv[role].items()}
            ev_r[role] = {r: [None if v is None else bacs[r] * v for v in vals] for r, vals in pct_ev[role].items()}
            ac_r[role] = sheet_values(company_sheet(SHEET_AC, company) if company else SHEET_AC, allow_fix)
        extra = dict(mode=mode, bac_role=bac_role, pct_pv=pct_pv, pct_ev=pct_ev, formula_gaps=gaps)
        if multi:
            pv = {r: [x + y for x, y in zip(pv_r["A"][r], pv_r["B"][r])] for r in pv_r["A"]}
            bac = {r: bac_role["A"][r] + bac_role["B"][r] for r in bac_role["A"]}
            return BookData(path, meta, structure, horizon, pv, combine(ev_r["A"], ev_r["B"]),
                            combine(ac_r["A"], ac_r["B"]), bac, method, status, fixable, True, a, b, pv_r["A"],
                            bac_role["A"], ev_r["A"], ev_r["B"], ac_r["A"], ac_r["B"], **extra)
        return BookData(path, meta, structure, horizon, pv_r[""], ev_r[""], ac_r[""], bac_role[""], method, status,
                        fixable, **extra)
    if meta.get("multiempresa"):
        a, b = _company_names(meta)
        names = {k: (company_sheet(k, a), company_sheet(k, b)) for k in (SHEET_PV, SHEET_EV, SHEET_AC)}
        pv_a = zeros(sheet_values(names[SHEET_PV][0], allow_fix))
        pv_b = zeros(sheet_values(names[SHEET_PV][1], allow_fix))
        bac_a, bac_b = bac_values(names[SHEET_PV][0]), bac_values(names[SHEET_PV][1])
        ev_a, ev_b = sheet_values(names[SHEET_EV][0], allow_fix), sheet_values(names[SHEET_EV][1], allow_fix)
        ac_a, ac_b = sheet_values(names[SHEET_AC][0], allow_fix), sheet_values(names[SHEET_AC][1], allow_fix)
        pv = {r: [x + y for x, y in zip(pv_a[r], pv_b[r])] for r in pv_a}
        bac = {r: bac_a[r] + bac_b[r] for r in bac_a}
        return BookData(path, meta, structure, horizon, pv, combine(ev_a, ev_b), combine(ac_a, ac_b), bac, method,
                        status, fixable, True, a, b, pv_a, bac_a, ev_a, ev_b, ac_a, ac_b, formula_gaps=gaps)
    pv = zeros(sheet_values(SHEET_PV, allow_fix))
    ev = sheet_values(SHEET_EV, allow_fix)
    ac = sheet_values(SHEET_AC, allow_fix)
    return BookData(path, meta, structure, horizon, pv, ev, ac, bac_values(SHEET_PV), method, status, fixable,
                    formula_gaps=gaps)


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
        meta = read_meta(wb)
        check_sheets(wb.sheetnames, meta, charts=True)
        return meta
    finally:
        wb.close()


def apply_book_settings(cfg: Config, meta: dict) -> Config:
    """Configuración multi-empresa con la que se generó el libro (nombres y campos de MS Project).

    La opción ac_desde_project se toma de la configuración actual.
    """
    cfg = replace(cfg)
    cfg.mode = book_mode(meta)
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


def check_sources(meta: dict, names: list[str]) -> None:
    """Los archivos de MS Project indicados deben ser los del libro: todos y con los mismos nombres.

    names: nombres de [[proyectos]]; [''] para un solo archivo (--origen).
    """
    book = book_sources(meta)
    if book == names or sorted(book) == sorted(names):
        return
    if book == [""]:
        raise ConfigError("El libro se generó con un solo archivo de MS Project: se debe indicar ese archivo con "
                          "--origen, con un archivo de configuración sin secciones [[proyectos]].")
    if names == [""]:
        raise ConfigError(f"El libro se generó con varios archivos de MS Project ({', '.join(book)}): se deben "
                          "indicar en secciones [[proyectos]] del archivo de configuración, con los mismos nombres.")
    raise ConfigError(f"Los archivos de [[proyectos]] ({', '.join(names)}) no coinciden con los del libro "
                      f"({', '.join(book)}). Se deben indicar todos los archivos del libro, con los mismos nombres; "
                      "para agregar o quitar un archivo se debe generar un libro nuevo con --migrar-desde.")


def check_only(cfg: Config, only: str | None) -> None:
    """En multi-empresa, --solo ac requiere ac_desde_project (el AC se carga a mano). En el modo porcentaje el AC
    siempre se carga a mano."""
    if cfg.percent_mode and only == "ac":
        raise ValueError("En el modo porcentaje el AC se carga a mano en las hojas AC (o en las de cada empresa): "
                         "'actualizar' solo calcula el % ganado, por lo que no se admite --solo ac.")
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
    deviations: list = field(default_factory=list)   # paquetes fuera de umbral (analysis.Deviation)
    thresholds: object | None = None                 # umbrales del semáforo (workbook.Thresholds)


def _acwp(t: Task) -> float:
    return t.acwp if t.acwp is not None else t.actual_cost


def _own_ac(task: Task, children: list[Task]) -> float:
    """Costo real de una tarea resumen que no está en sus filas hijas."""
    return _acwp(task) - sum(_acwp(c) for c in children)


def _own_bcwp(task: Task, children: list[Task]) -> float | None:
    if task.bcwp is None or any(c.bcwp is None for c in children):
        return None
    return task.bcwp - sum(c.bcwp for c in children)


def book_level(book: BookData, cfg: Config) -> int:
    """Nivel máximo del código EDT con que se generó el libro (libros antiguos: el de la configuración)."""
    return int(book.meta.get("nivel_maximo") or cfg.max_level)


def _match_tasks(project: ProjectData, book: BookData, cfg: Config,
                 source: str = "") -> tuple[dict[int, Task], list[str], Scope]:
    """Asocia cada fila del libro con la tarea de MS Project.

    Mismo proyecto: por UniqueID (si la tarea ya no existe, la fila se omite).
    Proyecto distinto: por código EDT, sin asignar una tarea a dos filas.
    Se usa la planificación con la que se generó el libro.
    source: con varios archivos, nombre del archivo; solo se asocian las filas de ese archivo.
    """
    warnings: list[str] = []
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)
    active: set[int] | None = None
    if not cfg.include_inactive:
        roots, _, _ = prune_inactive(roots)     # el avance de tareas inactivas no cuenta
        active = {t.uid for r in roots for t in iter_subtree(r)}
    plan_code = str(source_meta(book.meta, "planificacion", source) or "")
    if cfg.plan and cfg.plan.strip().upper() != plan_code.upper():
        warnings.append(f"El libro se generó para la planificación '{plan_code or 'todas'}': se usa esa "
                        f"planificación y no '{cfg.plan}'.")
    scope = select_scope(roots, plan_code)
    by_uid = {t.uid: t for t in tasks}
    by_code: dict[str, Task] = {}
    level = book_level(book, cfg)       # los códigos del libro se completaron con su nivel máximo
    for t in scope.tasks():
        raw = scope.raw_code(t)
        if raw:
            by_code.setdefault(clean_text(pad_code(raw, level) if cfg.pad_wbs else raw), t)
    book_name = source_meta(book.meta, "proyecto", source)
    same = str(book_name or "") == clean_text(project.name)
    if not same:
        warnings.append(f"El libro fue generado para el proyecto '{book_name}' y el origen es "
                        f"'{project.name}': las filas se asocian por código EDT. Se recomienda verificar los "
                        "resultados.")
    out: dict[int, Task] = {}
    used: set[tuple[int, bool]] = set()
    for s in book.structure:
        row_source, uid_text, residual = split_key(s.key)
        if row_source != source:
            continue
        if same and uid_text.isdigit():
            task = by_uid.get(int(uid_text))
            if task is None:
                if s.is_input:
                    warnings.append(f"Fila {s.code} '{s.name}': la tarea (UID {uid_text}) ya no existe en MS "
                                    "Project; se omite. Si cambió la EDT, se debe regenerar el libro con "
                                    "--migrar-desde.")
                continue
            if active is not None and task.uid not in active:
                if s.is_input:
                    warnings.append(f"Fila {s.code} '{s.name}': la tarea (UID {uid_text}) está inactiva en MS "
                                    "Project; se omite y se conservan los valores ya cargados. Si quedó fuera del "
                                    "alcance, se recomienda regenerar el libro con --migrar-desde.")
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


def _contributors(book: BookData, row: StructRow, task: Task, matches: dict[int, Task],
                  unlisted: list[Task] = ()) -> list[Task]:
    """Tareas cuyo costo forma la fila: la rama del paquete, o (fila -D) lo que cuelga de la
    tarea resumen sin pertenecer a sus filas hijas ni a paquetes agregados después de generar el libro."""
    if row.kind != KIND_RESIDUAL:
        return list(iter_subtree(task))
    taken = {n.uid for c in _row_children(book, row, matches) + list(unlisted) for n in iter_subtree(c)}
    return [n for n in iter_subtree(task) if n.uid not in taken]


def _row_children(book: BookData, row: StructRow, matches: dict[int, Task]) -> list[Task]:
    """Tareas de las filas hijas de la tarea resumen de una fila de otros costos (-D)."""
    return [matches[s.row] for s in book.structure
            if s.parent_row == row.parent_row and s.kind != "residual" and s.row in matches]


def _unlisted_packages(book: BookData, matches: dict[int, Task], scope: Scope, level: int) -> dict[int, list[Task]]:
    """Tareas con código EDT propio (de nivel no mayor al máximo del libro) que cuelgan de la tarea resumen
    de una fila -D sin tener fila en el libro: se agregaron después de generarlo. Al generar el libro tendrían
    fila propia, por eso no se suman en la fila -D (fila -> tareas)."""
    with_row = {t.uid for t in matches.values()}
    out: dict[int, list[Task]] = {}

    def visit(node: Task, found: list[Task]) -> None:
        for c in node.children:
            if c.uid in with_row:
                continue
            raw = scope.raw_code(c)
            if raw and code_level(raw) <= level:
                found.append(c)
            else:
                visit(c, found)         # tarea sin código o de nivel mayor: forma parte de la fila -D

    for s in book.input_rows:
        task = matches.get(s.row)
        if s.kind == KIND_RESIDUAL and task is not None:
            found: list[Task] = []
            visit(task, found)
            if found:
                out[s.row] = found
    return out


def _coverage_warnings(scope: Scope, book: BookData, matches: dict[int, Task],
                       unlisted: dict[int, list[Task]]) -> list[str]:
    """Costo real de tareas de la planificación que no queda en ninguna fila del libro."""
    covered: set[int] = set()
    for s in book.input_rows:
        task = matches.get(s.row)
        if task is not None:
            # Una fila -D cubre todo lo que cuelga de su tarea resumen y no tiene fila propia.
            covered.update(n.uid for n in iter_subtree(task))
    for packages in unlisted.values():  # salvo los paquetes con código agregados después de generar el libro
        covered.difference_update(n.uid for p in packages for n in iter_subtree(p))
    missing = []
    for t in scope.tasks():
        if t.uid in covered:
            continue
        own = _acwp(t) - sum(_acwp(c) for c in t.children)
        if own > tolerance(own):
            missing.append((t, own))
    warnings = []
    for t, amount in missing[:15]:
        code = scope.raw_code(t)
        how = "el libro" if code else "la tarea debe tener código EDT y el libro"
        warnings.append(f"{code or '(sin código EDT)'} '{t.name}': costo real {fmt_money(amount)} sin fila en el "
                        f"libro; no se incluye en el AC. Para incluirlo, {how} se debe regenerar con --migrar-desde.")
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


def _tag(source: str, text: str) -> str:
    """Con varios archivos, el mensaje indica el archivo: '[P2] ...'."""
    return f"[{source}] {text}" if source else text


def _percent_row(book: BookData, s: StructRow, contributors: list[Task], shares: dict[int, float], calendar,
                 cfg: Config, idx: int) -> tuple[RowResult, list[tuple[str, str, float, float]]]:
    """Modo porcentaje: % ganado acumulado de cada empresa en una fila a la fecha de corte.

    Devuelve la fila del resultado y, por empresa, (clave, empresa, % ganado acumulado, BAC).
    """
    from evm import percent as pct
    method = parse_method(book.method[s.row])
    times = pct.leaf_times([t for t in contributors if not t.children], calendar, cfg.distribution)
    full = pct.full_weights(times)
    out = []
    notes: list[str] = []
    for role, company in book.roles():
        share = None
        if role == "A":
            share = lambda t: shares.get(t.uid, 0.5)            # noqa: E731
        elif role == "B":
            share = lambda t: 1.0 - shares.get(t.uid, 0.5)      # noqa: E731
        rw = pct.company_weights(times, share, company, full)
        pv_cum = sum(book.pct_pv[role][s.row][: idx + 1])
        frac, note = pct.earned_fraction(method, rw, pv_cum)
        bac = book.bac_role[role][s.row]
        if note and note not in notes:
            notes.append(note)
        if rw.fallback and bac > 0:
            notes.append(f"{company} no participa en las tareas de la fila: se usó el avance de la fila completa")
        out.append((role, company, frac, bac))
    res = RowResult(s.code, s.name, book.method[s.row], book.bac[s.row], None, None, None, None, "; ".join(notes))
    return res, out


def run_update(projects: ProjectData | dict[str, ProjectData], cfg: Config, book_path: str | Path,
               status: date | None = None, only: str | None = None, backup: bool = True, force: bool = False,
               dry_run: bool = False) -> UpdateResult:
    """Carga el EV y el AC del mes de la fecha de corte.

    projects: el proyecto leído de MS Project o, en un libro con varios archivos, {nombre: proyecto} con
    todos los archivos de [[proyectos]].
    """
    if isinstance(projects, ProjectData):
        projects = {"": projects}
    book_path = Path(book_path)
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    check_sheets(wb.sheetnames, book.meta, charts=True)
    check_sources(book.meta, list(projects))
    cfg = apply_book_settings(cfg, book.meta)
    warnings: list[str] = [_tag(n, w) for n, p in projects.items() for w in p.warnings]
    multi = book.multi
    percent = book.percent
    load_ac = ((not multi) or cfg.ac_from_project) and not percent
    check_only(cfg, only)
    if percent and multi and cfg.ac_from_project:
        warnings.append("En el modo porcentaje el AC se carga a mano en las hojas de cada empresa: no se usa "
                        "ac_desde_project.")

    if status is None:
        missing = [n for n, p in projects.items() if p.status_date is None]
        if missing == [""]:
            raise ValueError("El proyecto no tiene fecha de estado: se debe indicar la fecha de corte con "
                             "--fecha-corte AAAA-MM-DD o definir Proyecto > Fecha de estado en MS Project.")
        if missing:
            raise ValueError(f"{'El archivo' if len(missing) == 1 else 'Los archivos'} {', '.join(missing)} no "
                             f"{'tiene' if len(missing) == 1 else 'tienen'} fecha de estado: se debe indicar la fecha "
                             "de corte con --fecha-corte AAAA-MM-DD o definir Proyecto > Fecha de estado en MS "
                             "Project.")
        dates = {n: p.status_date.date() for n, p in projects.items()}
        if len(set(dates.values())) > 1:
            raise ValueError("Los archivos tienen fechas de estado distintas ("
                             + ", ".join(f"{n}: {d:%d/%m/%Y}" for n, d in dates.items()) + "): se debe indicar la "
                             "fecha de corte con --fecha-corte AAAA-MM-DD o definir la misma fecha de estado en MS "
                             "Project.")
        status = next(iter(dates.values()))
    idx = book.horizon.index(status)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {status:%d/%m/%Y} está fuera del horizonte del libro "
                         f"({book.horizon.start:%m/%Y} a {book.horizon.end:%m/%Y}).")
    for n, project in projects.items():
        if project.status_date is None and project.backend == "xml":
            warnings.append(_tag(n, "El XML no tiene fecha de estado: MS Project calculó el costo real (CRTR) y el "
                                    "CPTR a la fecha en que se guardó el archivo. Se debe definir Proyecto > Fecha de "
                                    "estado antes de exportar."))
    if status != _month_end(status):
        warnings.append("La fecha de corte no es el último día del mes: el PV acumulado incluye el mes completo.")
    for n, project in projects.items():
        if project.status_date and project.status_date.date() != status:
            warnings.append(_tag(n, f"La fecha de estado de MS Project ({project.status_date:%d/%m/%Y}) difiere de "
                                    f"la fecha de corte ({status:%d/%m/%Y}). El costo real (CRTR) y el CPTR "
                                    "corresponden a la fecha de estado del proyecto."))
    if book.status_date and status < book.status_date:
        warnings.append(f"La fecha de corte ({status:%d/%m/%Y}) es anterior a la última cargada en el libro "
                        f"({book.status_date:%d/%m/%Y}).")

    # Cada archivo asocia sus filas (el UniqueID se repite entre archivos)
    matches: dict[int, Task] = {}
    unlisted: dict[int, list[Task]] = {}
    scopes: dict[str, Scope] = {}
    shares_of: dict[str, dict[int, float]] = {}
    calendars: dict[str, object] = {}                # modo porcentaje: calendario de cada archivo
    level = book_level(book, cfg)
    if multi:
        from evm.companies import ac_by_company, ev_by_company, resolve_shares
        sheets = {(k, r): company_sheet(k, book.company_a if r == "A" else book.company_b)
                  for k in (SHEET_EV, SHEET_AC) for r in ("A", "B")}
    for n, project in projects.items():
        pcfg = cfg.project_config(n)
        found, w, scope = _match_tasks(project, book, pcfg, n)
        warnings += [_tag(n, x) for x in w]
        extra = _unlisted_packages(book, found, scope, level)
        if only != "ev" and load_ac:
            warnings += [_tag(n, x) for x in _coverage_warnings(scope, book, found, extra)]
        matches.update(found)
        unlisted.update(extra)
        scopes[n] = scope
        if multi:
            shares_of[n], _ = resolve_shares(scope.roots, pcfg)
        if percent:
            from evm.percent import calendar_for
            calendars[n] = calendar_for(project, pcfg)

    rows: list[RowResult] = []
    later_filled: list[str] = []
    later_months: set[int] = set()
    writes: list[tuple[str, int, int, float]] = []   # (hoja, fila, columna, valor)
    ac_total: dict[str, float] = {}                  # AC acumulado de cada archivo
    no_bac: list[str] = []                           # modo porcentaje: filas con avance y sin BAC
    for s in book.input_rows:
        task = matches.get(s.row)
        if task is None:
            continue
        src = split_key(s.key)[0]
        shares = shares_of.get(src, {})
        bac = book.bac[s.row]
        method_text = book.method[s.row]
        method = parse_method(method_text)
        pv_cum = sum(book.pv[s.row][: idx + 1])
        residual = s.kind == "residual"
        siblings = _row_children(book, s, matches) + unlisted.get(s.row, []) if residual else []
        bcwp = _own_bcwp(task, siblings) if residual else task.bcwp
        bcwp_base = None
        if task.ev_baseline_cost is not None:       # CPTR calculado con otra línea base
            bcwp_base = task.ev_baseline_cost
            if residual:
                bcwp_base -= sum(c.ev_baseline_cost or 0.0 for c in siblings)
        # Tareas que forman la fila: en una fila -D, el avance se mide solo con las tareas sin fila propia
        contributors = _contributors(book, s, task, matches, unlisted.get(s.row, []))
        if percent:
            res, parts = _percent_row(book, s, contributors, shares, calendars[src], cfg, idx)
            for role, company, frac, bac_c in parts:
                sheet = pct_sheet(SHEET_EV, company)
                series = book.pct_ev[role][s.row]
                later = [i for i in range(idx + 1, len(series)) if series[i]]
                if later:
                    later_filled.append(f"{sheet} {s.code}")
                    later_months.update(later)
                prev = sum(v or 0.0 for v in series[:idx])
                month_value = round(frac - prev, 12)
                res.ev_cum = (res.ev_cum or 0.0) + round(bac_c * frac, 2)
                res.ev_month = (res.ev_month or 0.0) + round(bac_c * month_value, 2)
                if multi:
                    res.by_company[company_sheet(SHEET_EV, company)] = round(bac_c * month_value, 2)
                if month_value < -1e-9:
                    res.note = (res.note + "; " if res.note else "") + f"{sheet} del mes negativo (corrección)"
                if frac > 1e-9 and bac_c <= 0:
                    no_bac.append(f"{s.code} ({company})" if company else s.code)
                writes.append((sheet, s.row, FIRST_MONTH_COL + idx, month_value))
            rows.append(res)
            continue
        frac, note = ev_fraction(method, task, (pv_cum / bac) if bac else 0.0, bac=bac, bcwp=bcwp,
                                 bcwp_base=bcwp_base, tasks=contributors if residual else None)

        # (tipo, hoja, valor acumulado, serie del libro)
        targets: list[tuple[str, str, float, list]] = []
        if multi:
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
            later = [i for i in range(idx + 1, len(series)) if series[i]]
            if later:
                later_filled.append(f"{sheet} {s.code}")
                later_months.update(later)
            prev = sum(v or 0.0 for v in series[:idx])
            month_value = round(cum_value - prev, 2)
            if kind == "ev":
                res.ev_cum = (res.ev_cum or 0.0) + cum_value
                res.ev_month = (res.ev_month or 0.0) + month_value
            else:
                res.ac_cum = (res.ac_cum or 0.0) + cum_value
                res.ac_month = (res.ac_month or 0.0) + month_value
                ac_total[src] = ac_total.get(src, 0.0) + cum_value
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
    if later_filled:
        months = [f"{book.horizon.month_start(i):%m/%Y}" for i in sorted(later_months)]
        warnings.append(
            f"Se cargó la fecha de corte {status:%d/%m/%Y} con --forzar, pero hay valores en meses posteriores ("
            + ", ".join(months[:6]) + ("..." if len(months) > 6 else "") + "; filas " + ", ".join(later_filled[:10])
            + ("..." if len(later_filled) > 10 else "") + "). Esos valores se calcularon con el acumulado anterior: "
            "los acumulados de esos meses ya no coinciden con MS Project. Se deben revisar esos meses o volver a "
            "cargarlos con 'actualizar' (fecha de corte de cada mes, en orden).")

    for n, project in projects.items():
        scope = scopes[n]
        reference = _acwp(scope.plan) if scope.plan is not None else project.summary_acwp
        if only != "ev" and load_ac and reference is not None:
            loaded = ac_total.get(n, 0.0)
            diff = reference - loaded
            if abs(diff) > tolerance(reference):
                whose = f"de la planificación '{scope.prefix}'" if scope.plan is not None else "del proyecto"
                warnings.append(_tag(n, f"El AC cargado ({fmt_money(loaded)}) difiere del costo real {whose} en MS "
                                        f"Project ({fmt_money(reference)}) en {fmt_money(diff)}. Se recomienda revisar "
                                        "las advertencias anteriores (tareas sin fila, costos en tareas resumen)."))

    if percent:
        where = (f"las hojas {company_sheet(SHEET_PV, book.company_a)} y {company_sheet(SHEET_PV, book.company_b)}"
                 if multi else f"la hoja {SHEET_PV}")
        if not any(book.bac[s.row] for s in book.input_rows):
            warnings.append(f"Ninguna fila tiene BAC: el PV, el EV y los indicadores quedan en 0 hasta ingresar el "
                            f"BAC de cada fila en la columna D de {where}.")
        elif no_bac:
            warnings.append(f"{plural(len(no_bac), 'fila tiene', 'filas tienen')} % ganado pero no "
                            f"{'tiene' if len(no_bac) == 1 else 'tienen'} BAC: el EV no se puede valorar hasta "
                            f"ingresar el BAC en la columna D de {where} (" + ", ".join(no_bac[:15]) +
                            (f" y {len(no_bac) - 15} más" if len(no_bac) > 15 else "") + ").")

    if book.formula_gaps:
        cells = ", ".join(book.formula_gaps[:8])
        more = f" y {len(book.formula_gaps) - 8} más" if len(book.formula_gaps) > 8 else ""
        warnings.append(f"Hay celdas sin la fórmula del libro (subtotales, fila TOTAL o suma de las empresas): "
                        f"{cells}{more}. Los totales e indicadores del libro no coinciden con los que se muestran "
                        "aquí: se deben restaurar esas fórmulas (por ejemplo, copiándolas de una celda vecina de la "
                        "misma fila) o recuperar el libro desde un respaldo.")

    if book.fixable:
        cells = ", ".join(f"{sh}!{col(c)}{r}" for sh, r, c, _v, _why in book.fixable[:8])
        more = f" y {len(book.fixable) - 8} más" if len(book.fixable) > 8 else ""
        warnings.append(f"Se reemplazaron por su valor numérico celdas de ingreso con texto o fórmulas: {cells}{more}.")

    # Valores en memoria (también en simulación, para mostrar los indicadores resultantes)
    pct_roles = {pct_sheet(SHEET_EV, company): role for role, company in book.roles()} if percent else {}
    for sheet, r, c, v in writes:
        i = c - FIRST_MONTH_COL
        if percent:
            role = pct_roles[sheet]
            book.pct_ev[role][r][i] = v
            amount = book.bac_role[role][r] * v
            if not multi:
                book.ev[r][i] = amount
                continue
            (book.ev_a if role == "A" else book.ev_b)[r][i] = amount
            x, y = book.ev_a[r][i], book.ev_b[r][i]
            book.ev[r][i] = None if x is None and y is None else (x or 0.0) + (y or 0.0)
            continue
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

    # Paquetes fuera de umbral en el mes de corte (con el EV y el AC que quedan en el libro)
    th = analysis.thresholds(book.meta, cfg)
    found = analysis.deviations(book, idx, th)

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
        _write_analysis(wb, book, idx, found, th)
        _refresh_charts(wb, book, idx, meta)
        backup_path = _save(wb, book_path, backup)
    book.status_date = status
    return UpdateResult(status, idx, rows, warnings, backup_path, book, dry_run, ac_loaded=load_ac,
                        deviations=found, thresholds=th)


def _save(wb, book_path: Path, backup: bool) -> Path | None:
    """Guarda el libro (fórmulas recalculadas al abrirlo) con una copia de respaldo previa, si se pide."""
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    backup_path = None
    if backup:
        backup_path = _backup_path(book_path)
        shutil.copy2(book_path, backup_path)
    try:
        save_workbook(wb, book_path)
    except BaseException:
        if backup_path is not None:     # el libro no cambió: el respaldo no hace falta
            backup_path.unlink(missing_ok=True)
        raise
    return backup_path


def _write_analysis(wb, book: BookData, idx: int, found: list, th) -> None:
    """Agrega a la hoja Análisis de desviaciones los paquetes fuera de umbral del mes idx (la crea en los
    libros de versiones anteriores), conservando el texto ya escrito."""
    entries = analysis.read_entries(wb[analysis.SHEET_ANALYSIS]) if analysis.SHEET_ANALYSIS in wb.sheetnames else []
    month = book.horizon.month_start(idx)
    entries = analysis.refresh(entries, found, month)
    analysis.write_sheet(wb, entries, {s.key: s.row for s in book.input_rows}, book.horizon.months, th,
                         int(book.meta.get("decimales") or 0), month, len(found))


def run_analysis(cfg: Config, book_path: str | Path, backup: bool = True) -> UpdateResult:
    """Comando 'analizar': actualiza la hoja Análisis de desviaciones con los valores del libro, sin leer MS
    Project (por ejemplo, después de cargar a mano el AC del mes)."""
    book_path = Path(book_path)
    wb = load_workbook(book_path)
    book = read_book(book_path, wb, allow_fix=True)
    check_sheets(wb.sheetnames, book.meta, charts=True)
    if book.status_date is None:
        raise ValueError("El libro no tiene fecha de corte (hoja Resumen, celda B5): se debe ejecutar 'actualizar' "
                         "antes de 'analizar'.")
    idx = book.horizon.index(book.status_date)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {book.status_date:%d/%m/%Y} está fuera del horizonte del libro.")
    warnings = []
    if book.fixable:
        cells = ", ".join(f"{sh}!{col(c)}{r}" for sh, r, c, _v, _why in book.fixable[:8])
        more = f" y {len(book.fixable) - 8} más" if len(book.fixable) > 8 else ""
        warnings.append(f"Se reemplazaron por su valor numérico celdas de ingreso con texto o fórmulas: {cells}{more}.")
        for sheet, r, c, v, _why in book.fixable:
            wb[sheet].cell(row=r, column=c, value=v)
    th = analysis.thresholds(book.meta, cfg)
    found = analysis.deviations(book, idx, th)
    _write_analysis(wb, book, idx, found, th)
    _refresh_charts(wb, book, idx, book.meta)
    backup_path = _save(wb, book_path, backup)
    return UpdateResult(book.status_date, idx, [], warnings, backup_path, book, deviations=found, thresholds=th)


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
        dashboard.add_dashboard_charts(wb, packages, book.horizon.months, plan_end(book), idx + 1,
                                       bac, currency, idx + 1 - loaded[0] if loaded else 0)
    elif LEGACY_SHEET_CURVE in wb.sheetnames:
        add_s_curve_chart(wb[LEGACY_SHEET_CURVE], book.horizon.months, str(meta.get("proyecto") or ""), currency)


def plan_end(book: BookData) -> int:
    """Último mes con PV planificado (N° de mes). En el modo porcentaje, el último mes con % planificado, aunque
    falte ingresar algún BAC."""
    from evm import dashboard
    if not book.percent:
        return dashboard.plan_end_month(book_totals(book)[0])
    last = 0
    for series in book.pct_pv.values():
        for vals in series.values():
            for i, v in enumerate(vals):
                if abs(v) > 1e-9 and i + 1 > last:
                    last = i + 1
    return last


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
