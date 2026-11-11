"""Comparación de la planificación del cliente con la planificación interna (comando generar-cliente).

Se compara:

- el % completado del proyecto, calculado como MS Project: la duración real de las tareas de detalle activas
  (duración × % completado) dividida por su duración;
- el libro EVM, fila a fila: cada fila de ingreso (la que tiene % planificado y % ganado) debe tener las mismas tareas
  de detalle, el mismo % planificado de cada mes y el mismo % ganado con el avance actual. Una fila del libro interno
  que en el del cliente tiene otro código se empareja por sus tareas de detalle (por ejemplo, 1.2.3, cuyo código se
  quitó, pasa a la fila 1.2.0-D);
- con actividad_interna y el archivo generado, los datos de cada tarea: fechas, avance, duración y línea base.

La comparación es una previsión con --simular (la planificación del cliente se arma en memoria, client_model) y una
comprobación al generar el archivo (se lee el archivo del cliente en MS Project después de guardarlo).
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field, replace
from datetime import datetime

from evm.client_plan import INTERNAL_NAME, ClientPlan, _active, rollup_percent
from evm.config import Config
from evm.fields import fmt_number, same_name
from evm.messages import EvmError, Issue
from evm.methods import parse_method
from evm.model import MonthKey, ProjectData, Task
from evm.percent import calendar_for, curve, earned_fraction, full_weights, leaf_times
from evm.wbs import build_report, build_tree, plural, select_scope

TOL = 1e-6                  # fracciones del % planificado de un mes y del % ganado
PERCENT_TOL = 0.05          # puntos del % completado del proyecto
MAX_LINES = 500             # filas de cada tabla del informe


@dataclass
class BookRow:
    """Fila de ingreso del libro EVM."""

    code: str
    leaves: frozenset[int]                  # UID de sus tareas de detalle
    pv: dict[MonthKey, float]               # % planificado de cada mes (fracción; fila completa)
    ev: float                               # % ganado con el avance actual (fracción)


def _pv_gap(a: BookRow, b: BookRow) -> float:
    """Mayor diferencia del % planificado de un mes (fracción)."""
    return max((abs(a.pv.get(k, 0.0) - b.pv.get(k, 0.0)) for k in set(a.pv) | set(b.pv)), default=0.0)


def _equal(a: BookRow, b: BookRow) -> bool:
    return a.leaves == b.leaves and _pv_gap(a, b) < TOL and abs(a.ev - b.ev) < TOL


@dataclass
class BookComparison:
    """Libro EVM del cliente frente al de la planificación interna (filas de ingreso)."""

    rows_internal: int = 0
    rows_client: int = 0
    same: int = 0                                                       # mismo código, tareas y valores
    moved: list[tuple[BookRow, BookRow]] = field(default_factory=list)  # otro código, mismas tareas y valores
    different: list[tuple[BookRow, BookRow]] = field(default_factory=list)  # otras tareas o valores
    only_internal: list[BookRow] = field(default_factory=list)
    only_client: list[BookRow] = field(default_factory=list)
    d_rows: list[str] = field(default_factory=list)                     # filas -D del libro del cliente
    # Filas sin pareja: filas del otro libro que tienen sus tareas de detalle (p. ej. 1.1.0 y 1.2.0 -> 1.0.0)
    where: dict[str, list[str]] = field(default_factory=dict)
    leaves_internal: int = 0                                            # tareas de detalle del libro
    leaves_client: int = 0
    error: str = ""                                                     # no se pudo calcular el libro

    @property
    def deviations(self) -> int:
        return len(self.different) + len(self.only_internal) + len(self.only_client)


@dataclass
class TaskDiff:
    """Tarea del archivo del cliente con datos distintos de la planificación interna."""

    task: Task                                  # tarea de la planificación interna
    fields: list[tuple[str, str, str]]          # (dato, planificación interna, archivo del cliente)


@dataclass
class Comparison:
    generated: bool                             # False = previsión (--simular)
    percent_internal: float | None              # % completado del proyecto (como lo calcula MS Project)
    percent_client: float | None
    book: BookComparison
    tasks: list[TaskDiff] = field(default_factory=list)   # actividad_interna y archivo generado
    checked: int = 0                                      # tareas comparadas

    @property
    def percent_differs(self) -> bool:
        a, b = self.percent_internal, self.percent_client
        return (a is None) != (b is None) or (a is not None and abs(a - b) >= PERCENT_TOL)

    def issues(self) -> list[Issue]:
        out: list[Issue] = []
        is_ = "es" if self.generated else "sería"
        if self.percent_differs:
            out.append(Issue("CLI-023", message=f"El % completado del proyecto {is_} {_pct(self.percent_client)} en el "
                                                f"archivo del cliente y es {_pct(self.percent_internal)} en la "
                                                "planificación interna: ver el informe de control."))
        book = self.book
        if book.error:
            out.append(Issue("CLI-029", message=f"No se pudo comparar el libro EVM: {book.error}"))
        elif book.deviations:
            what = plural(book.deviations, "fila del libro EVM", "filas del libro EVM")
            have = "tiene" if book.deviations == 1 else "tienen"
            if not self.generated:
                have = "tendría" if book.deviations == 1 else "tendrían"
            out.append(Issue("CLI-024", message=f"{what} del cliente {have} otras tareas o valores (% planificado o % "
                                                "ganado) que en el libro de la planificación interna: ver el informe "
                                                "de control."))
        if book.d_rows:
            before = {b.code: a.code for a, b in book.moved}
            shown = ", ".join(f"{c} (antes {before[c]})" if c in before else c for c in book.d_rows[:10])
            more = f" y {len(book.d_rows) - 10} más" if len(book.d_rows) > 10 else ""
            have = "tiene" if self.generated else "tendría"
            out.append(Issue("CLI-025", message=f"El libro EVM del cliente {have} "
                                                f"{plural(len(book.d_rows), 'fila', 'filas')} -D («Otras tareas de "
                                                f"…»): {shown}{more}."))
        if self.tasks:
            what = plural(len(self.tasks), "tarea tiene", "tareas tienen")
            out.append(Issue("CLI-026", message=f"{what} en el archivo del cliente datos distintos de la planificación "
                                                "interna (fechas, avance, duración o línea base): ver el informe de "
                                                "control."))
        return out


# ----------------------------------------------------------------------------
# Planificación del cliente en memoria
# ----------------------------------------------------------------------------
def _copy_tasks(tasks: list[Task]) -> list[Task]:
    """Copias de las tareas sin su estructura (la arma build_tree), para no cambiar la de la planificación interna."""
    out = []
    for t in tasks:
        c = copy.copy(t)
        c.parent, c.children, c.predecessors = None, [], list(t.predecessors)
        out.append(c)
    return out


def _copy_project(project: ProjectData, tasks: list[Task]) -> ProjectData:
    return replace(project, tasks=tasks, warnings=[], fields_found=dict(project.fields_found))


def client_model(plan: ClientPlan) -> ProjectData:
    """Planificación del cliente prevista: sin las tareas que se borran; con actividad_interna, las que no se muestran
    como «Otras tareas» (sin nombre ni código EDT); sin actividad_interna, los resúmenes convertidos como tareas."""
    keep = plan.remaining
    hidden = {t.uid for t in plan.hidden}
    collapsed = {t.uid for t in plan.collapsed}
    tasks = _copy_tasks([t for t in plan.project.tasks if t.uid in keep])
    for t in tasks:
        t.client_mark = ""
        t.predecessors = [lk for lk in t.predecessors if lk.uid in keep]
        if t.uid in hidden:
            t.name, t.description = INTERNAL_NAME, ""
            if plan.codes_removed:
                t.wbs = ""
        if t.uid in collapsed:
            t.summary = False
        if plan.root is not None and t.uid == plan.root.uid and plan.new_name:
            t.name = plan.new_name
    return _copy_project(plan.project, tasks)


# ----------------------------------------------------------------------------
# Comparación
# ----------------------------------------------------------------------------
def project_percent(project: ProjectData, cfg: Config) -> float | None:
    """% completado de la planificación como lo calcula MS Project (tareas de detalle activas)."""
    tasks = _copy_tasks(project.tasks)
    scope = select_scope(build_tree(tasks), cfg.plan)
    return rollup_percent([t for t in scope.tasks() if not t.children and _active(t)])


def _book_rows(project: ProjectData, cfg: Config, status: datetime | None) -> list[BookRow]:
    """Filas de ingreso del libro EVM, con su % planificado mensual y su % ganado con el avance actual."""
    report = build_report(_copy_project(project, _copy_tasks(project.tasks)), cfg)
    calendar = calendar_for(project, cfg)
    key = (status.year, status.month) if status is not None else None
    rows = []
    for r in report.rows:
        if not r.is_input:
            continue
        times = leaf_times(r.leaves, calendar, cfg.distribution)
        weights = full_weights(times)
        pv = curve(times, weights)
        pv_cum = sum(v for k, v in pv.items() if key is None or k <= key)
        ev, _note = earned_fraction(parse_method(r.method), weights, pv_cum)
        rows.append(BookRow(r.code, frozenset(t.uid for t in r.leaves), pv, ev))
    return rows


def compare_books(internal: ProjectData, client: ProjectData, cfg: Config,
                  client_cfg: Config | None = None) -> BookComparison:
    """Libro EVM del cliente frente al de la planificación interna."""
    status = internal.status_date
    try:
        a, b = _book_rows(internal, cfg, status), _book_rows(client, client_cfg or cfg, status)
    except EvmError as exc:
        return BookComparison(error=f"{exc.id}: {exc}" if getattr(exc, "id", "") else str(exc))
    out = BookComparison(rows_internal=len(a), rows_client=len(b),
                         leaves_internal=len(set().union(*(r.leaves for r in a))) if a else 0,
                         leaves_client=len(set().union(*(r.leaves for r in b))) if b else 0)
    by_code = {r.code: r for r in b}
    matched: set[str] = set()
    pending: list[BookRow] = []
    for r in a:
        other = by_code.get(r.code)
        if other is None:
            pending.append(r)
            continue
        matched.add(r.code)
        if _equal(r, other):
            out.same += 1
        else:
            out.different.append((r, other))
    by_leaves = {r.leaves: r for r in b if r.code not in matched}
    for r in pending:
        other = by_leaves.pop(r.leaves, None)
        if other is None:
            out.only_internal.append(r)
            continue
        matched.add(other.code)
        (out.moved if _equal(r, other) else out.different).append((r, other))
    out.only_client = [r for r in b if r.code not in matched]
    out.d_rows = [r.code for r in b if r.code.endswith("-D")]
    for rows, others in ((out.only_internal, b), (out.only_client, a)):
        owner = {uid: r.code for r in others for uid in r.leaves}
        for r in rows:
            out.where[r.code] = sorted({owner[uid] for uid in r.leaves if uid in owner})
    return out


def _same_minute(a: datetime | None, b: datetime | None) -> bool:
    if a is None or b is None:
        return a is b
    return a.replace(second=0, microsecond=0) == b.replace(second=0, microsecond=0)


def _dt(value: datetime | None) -> str:
    return f"{value:%d/%m/%Y %H:%M}" if value else "—"


def _days(minutes: float | None, per_day: float) -> str:
    return "—" if minutes is None else fmt_number(round(minutes / per_day, 2), trim=True)


def compare_tasks(plan: ClientPlan, client: ProjectData) -> tuple[list[TaskDiff], int]:
    """actividad_interna: cada tarea del archivo del cliente (también los resúmenes y las «Otras tareas») debe tener
    las fechas, el avance, la duración y la línea base de la planificación interna. Devuelve las diferencias y la
    cantidad de tareas comparadas."""
    per_day = plan.project.minutes_per_day or 480.0
    found = {t.uid: t for t in client.tasks}
    keep = plan.remaining
    targets = [t for t in ([plan.root] if plan.root is not None else []) + plan.tasks if t.uid in keep]
    diffs: list[TaskDiff] = []
    for t in targets:
        c = found.get(t.uid)
        if c is None:
            diffs.append(TaskDiff(t, [("Tarea", "Sí", "No está")]))
            continue
        fields: list[tuple[str, str, str]] = []
        for label, a, b in (("Comienzo", t.start, c.start), ("Fin", t.finish, c.finish),
                            ("Comienzo real", t.actual_start, c.actual_start),
                            ("Fin real", t.actual_finish, c.actual_finish),
                            ("Comienzo de la línea base", t.baseline_start, c.baseline_start),
                            ("Fin de la línea base", t.baseline_finish, c.baseline_finish)):
            if not _same_minute(a, b):
                fields.append((label, _dt(a), _dt(b)))
        for label, a, b in (("% completado", t.percent_complete, c.percent_complete),
                            ("% físico completado", t.physical_percent_complete, c.physical_percent_complete)):
            if abs(a - b) >= 0.5:
                fields.append((label, _pct(a), _pct(b)))
        for label, a, b in (("Duración (días)", t.duration, c.duration),
                            ("Duración de la línea base (días)", t.baseline_duration, c.baseline_duration)):
            if (a is None) != (b is None) or (a is not None and abs(a - b) >= 1):
                fields.append((label, _days(a, per_day), _days(b, per_day)))
        if t.active != c.active:
            fields.append(("Activa", "Sí" if t.active else "No", "Sí" if c.active else "No"))
        if t.milestone != c.milestone:
            fields.append(("Hito", "Sí" if t.milestone else "No", "Sí" if c.milestone else "No"))
        if fields:
            diffs.append(TaskDiff(t, fields))
    return diffs, len(targets)


def client_config(plan: ClientPlan, cfg: Config) -> Config:
    """Configuración para el archivo del cliente: la planificación se busca por su código EDT o, si se eligió por su
    nombre, por el que tiene en el archivo del cliente ([cliente] nombre_planificacion)."""
    if plan.root is None or same_name(plan.root.wbs or "", cfg.plan) or not plan.new_name:
        return cfg
    return replace(cfg, plan=plan.new_name)


def compare(plan: ClientPlan, client: ProjectData, cfg: Config, generated: bool) -> Comparison:
    """Planificación del cliente (prevista o leída del archivo generado) frente a la planificación interna."""
    client_cfg = client_config(plan, cfg)
    out = Comparison(generated, project_percent(plan.project, cfg), project_percent(client, client_cfg),
                     compare_books(plan.project, client, cfg, client_cfg))
    if generated and plan.internal_activity:
        out.tasks, out.checked = compare_tasks(plan, client)
    return out


def predict(plan: ClientPlan, cfg: Config) -> Comparison:
    """Previsión sin MS Project (--simular)."""
    return compare(plan, client_model(plan), cfg, generated=False)


# ----------------------------------------------------------------------------
# Informe de control
# ----------------------------------------------------------------------------
def _pct(value: float | None) -> str:
    return "—" if value is None else f"{fmt_number(round(value, 1), trim=True)} %"


def _cell(value) -> str:
    return " ".join(str(value or "").split()).replace("|", "\\|")


def _more(n: int) -> list[str]:
    return [f"| … y {n - MAX_LINES} más | | | | | |", ""] if n > MAX_LINES else [""]


def comparison_lines(plan: ClientPlan, comp: Comparison, per_day: float) -> list[str]:
    """Sección del informe de control con la comparación."""
    book = comp.book
    how = ("Archivo del cliente leído en MS Project después de guardarlo." if comp.generated else
           "Previsión, sin MS Project: la planificación del cliente se arma a partir de la interna (--simular).")
    lines = ["## Comparación con la planificación interna", "", how, "",
             "| Dato | Planificación interna | Archivo del cliente |", "|---|---:|---:|",
             f"| % completado del proyecto (como lo calcula MS Project) | {_pct(comp.percent_internal)} | "
             f"{_pct(comp.percent_client)} |"]
    if comp.checked:
        lines.append(f"| Tareas comparadas (fechas, avance, duración y línea base) | {comp.checked} | "
                     f"{comp.checked - len(comp.tasks)} iguales |")
    if book.error:
        lines += ["", f"No se pudo comparar el libro EVM: {_cell(book.error)}", ""]
        return lines + _task_lines(plan, comp)
    lines += [f"| Filas de ingreso del libro EVM | {book.rows_internal} | {book.rows_client} |",
              f"| Tareas de detalle en el libro EVM | {book.leaves_internal} | {book.leaves_client} |", "",
              "Cada fila de ingreso del libro debe tener las mismas tareas de detalle, el mismo % planificado de cada "
              "mes y el mismo % ganado (con el avance actual).", "",
              "| Filas del libro EVM | Cantidad |", "|---|---:|",
              f"| Iguales (mismo código, tareas y valores) | {book.same} |",
              f"| Con otro código (mismas tareas y valores) | {len(book.moved)} |",
              f"| Con otras tareas o valores | {len(book.different)} |",
              f"| Solo en el libro interno | {len(book.only_internal)} |",
              f"| Solo en el libro del cliente | {len(book.only_client)} |", ""]
    if book.d_rows:
        before = {b.code: a.code for a, b in book.moved}
        lines += ["### Filas -D en el libro del cliente", "",
                  "Reúnen tareas sin código EDT (en el libro, «Otras tareas de …»): por ejemplo, las de un elemento EDT "
                  "que no se muestra, cuyo código se quitó. El cliente las ve en su libro. Para evitarlas, se debe "
                  "mostrar ese elemento (Sí en el campo del cliente) o reorganizar la EDT.", "",
                  "| Fila del cliente | Fila del libro interno |", "|---|---|"]
        lines += [f"| {code} | {before.get(code, code)} |" for code in book.d_rows]
        lines.append("")
    rows = [(a, b) for a, b in book.different] + [(a, None) for a in book.only_internal] + \
           [(None, b) for b in book.only_client]
    if rows:
        lines += ["### Filas del libro EVM con otras tareas o valores", "",
                  "% planificado: la mayor diferencia de un mes, en puntos. % ganado: con el avance actual.", "",
                  "| Fila interna | Fila del cliente | Tareas de detalle (interno / cliente) | Diferencia del % "
                  "planificado | % ganado interno | % ganado cliente |", "|---|---|---:|---:|---:|---:|"]
        for a, b in rows[:MAX_LINES]:
            gap = _pct(_pv_gap(a, b) * 100) if a is not None and b is not None else "—"
            if a is None or b is None:          # dónde están sus tareas de detalle en el otro libro
                alone = a or b
                found = ", ".join(book.where.get(alone.code, [])) or "ninguna"
                a_text = a.code if a else f"— (sus tareas: {found})"
                b_text = b.code if b else f"— (sus tareas: {found})"
            else:
                a_text, b_text = a.code, b.code
            lines.append(f"| {a_text} | {b_text} | "
                         f"{len(a.leaves) if a else 0} / {len(b.leaves) if b else 0} | {gap} | "
                         f"{_pct(a.ev * 100) if a else '—'} | {_pct(b.ev * 100) if b else '—'} |")
        lines += _more(len(rows))
    return lines + _task_lines(plan, comp)


def _task_lines(plan: ClientPlan, comp: Comparison) -> list[str]:
    if not comp.tasks:
        return []
    rows = [(d.task, f) for d in comp.tasks for f in d.fields]
    lines = ["### Tareas con datos distintos", "",
             f"Tareas del archivo del cliente (también resúmenes y «{INTERNAL_NAME}») cuyos datos no son los de la "
             "planificación interna.", "", "| Tarea | Dato | Planificación interna | Archivo del cliente |",
             "|---|---|---|---|"]
    lines += [f"| {plan.ref(t)} | {label} | {a} | {b} |" for t, (label, a, b) in rows[:MAX_LINES]]
    return lines + (["| … | | | |", ""] if len(rows) > MAX_LINES else [""])
