"""Informe de mejoras de validar (--informe) y datos que validar muestra al final: evaluación DCMA de 14 puntos y
comparación de planificaciones.

El informe es un documento Markdown con las observaciones agrupadas por prioridad (alta: las que cambian el libro;
media: las de la calidad del cronograma, que solo revisa validar; baja: las informativas), cada una con la causa y
qué hacer del catálogo de mensajes. No incluye nombres de tareas, de empresas ni de calendarios, ni el título del
proyecto: las tareas se identifican por su código EDT o, sin código, por su UniqueID. Los textos de las observaciones
generales se incluyen sin esos nombres. Las observaciones informativas que describen la estructura (FACT_CODES) no se
listan como mejoras: sus datos están en el resumen.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from evm.config import Config
from evm.fields import fmt_number
from evm.messages import CATALOG, SEV_ERROR, SEV_INFO, SEV_WARN, W_V, Issue
from evm.methods import is_complete
from evm.model import ProjectData, Task
from evm.wbs import KIND_RESIDUAL, Report, ReportPart, ReportRow, code_key, plural

PRIORITY_HIGH, PRIORITY_MEDIUM, PRIORITY_LOW = "Alta", "Media", "Baja"
PRIORITIES = [PRIORITY_HIGH, PRIORITY_MEDIUM, PRIORITY_LOW]
PRIORITY_TEXT = {
    PRIORITY_HIGH: "errores y advertencias que cambian el libro (filas, % planificado, % ganado u horizonte)",
    PRIORITY_MEDIUM: "advertencias de la calidad del cronograma (línea base, avance y lógica), que solo revisa validar",
    PRIORITY_LOW: "observaciones informativas",
}
# Informativas que describen la estructura del libro: no son mejoras y sus datos están en el resumen (tareas
# inactivas, planificación, filas «Otras tareas», tareas sumadas en su fila superior, Visible = No y líneas base)
FACT_CODES = ("EDT-016", "EDT-017", "EDT-018", "EDT-022", "VIS-012", "LEC-024")


@dataclass
class Unit:
    """Una revisión de validar: un archivo (o los de [[proyectos]]) con una planificación."""

    label: str                          # planificación revisada ("" = la de la configuración)
    parts: list[ReportPart]
    report: Report | None
    issues: list[Issue]                 # observaciones ya silenciadas y con su alcance
    status: date | None = None          # --fecha-corte (None = la fecha de estado de cada archivo)

    @property
    def title(self) -> str:
        return self.label or (self.parts[0].source.plan if len(self.parts) == 1 else "") or "Proyecto"


def priority(issue: Issue) -> str:
    """Alta: errores, y advertencias de los controles que también hace generar (cambian el libro). Media: advertencias
    de los controles que solo hace validar (línea base, avance y lógica del cronograma). Baja: informativas."""
    if issue.severity == SEV_ERROR:
        return PRIORITY_HIGH
    if issue.severity == SEV_INFO:
        return PRIORITY_LOW
    entry = CATALOG.get(issue.id)
    return PRIORITY_MEDIUM if entry is not None and entry.location == W_V else PRIORITY_HIGH


def _num(v: float, decimals: int = 1) -> str:
    return fmt_number(v, decimals, trim=True)


def _day(d: datetime | date | None) -> str:
    return f"{d:%d/%m/%Y}" if d else "-"


def _active_chain(t: Task) -> bool:
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


# ----------------------------------------------------------------------------
# Datos de cada revisión
# ----------------------------------------------------------------------------
@dataclass
class Facts:
    """Cifras de una revisión (un archivo con una planificación), para el resumen y la comparación."""

    label: str
    tasks: int = 0                      # tareas de la planificación (todas)
    active: int = 0
    inactive: int = 0
    hidden: int = 0                     # tareas activas con Visible = No (no se revisan)
    reviewed: int = 0
    rows: int = 0
    inputs: int = 0
    residual: int = 0                   # filas «Otras tareas» (-D)
    aggregated: int = 0                 # tareas sumadas en la fila de su elemento superior
    baselines: dict[int, int] = field(default_factory=dict)
    not_rolled_up: int = 0              # tareas resumen con la línea base sin resumir (FEC-013)
    links_in: int = 0                   # vínculos entre tareas de la planificación
    links_out: int = 0                  # vínculos con tareas de otras planificaciones
    critical: str = "-"                 # % de tareas de detalle sin terminar que son críticas
    baseline_finish: datetime | None = None
    finish: datetime | None = None
    status: date | None = None
    baseline_label: str = ""
    errors: int = 0
    warnings: int = 0
    infos: int = 0


def facts(part: ReportPart, issues: list[Issue], label: str = "", status: date | None = None) -> Facts:
    project: ProjectData = part.source.project
    scope = part.scope
    every, _members = _plan_tasks(part)
    out = Facts(label or part.source.plan or part.source.name or "Proyecto")
    out.tasks = len(every)
    out.active = sum(1 for t in every if _active_chain(t))
    out.inactive = out.tasks - out.active
    reviewed = list(scope.tasks()) if scope is not None else []
    out.reviewed = len(reviewed)
    out.hidden = max(0, out.active - out.reviewed)
    out.rows = len(part.rows)
    out.inputs = sum(1 for r in part.rows if r.is_input)
    for t in reviewed:
        for n in t.baselines:
            out.baselines[n] = out.baselines.get(n, 0) + 1
    out.baselines = dict(sorted(out.baselines.items()))
    out.residual = sum(1 for r in part.rows if r.kind == KIND_RESIDUAL)
    out.aggregated = part.aggregated
    out.not_rolled_up = sum(1 for i in issues if i.id == "FEC-013" and (i.code or i.uid is not None))
    by_uid = {t.uid: t for t in project.tasks}
    inside = {t.uid for t in every}
    for t in project.tasks:
        for link in t.predecessors:
            p = by_uid.get(link.uid)
            if p is None or not _active_chain(t) or not _active_chain(p):
                continue
            a, b = t.uid in inside, p.uid in inside
            if a and b:
                out.links_in += 1
            elif a or b:
                out.links_out += 1
    leaves = [t for t in reviewed if not t.children]
    open_tasks = [t for t in leaves if not is_complete(t)]
    if open_tasks:
        out.critical = f"{_num(100.0 * sum(1 for t in open_tasks if t.critical) / len(open_tasks))} %"
    bl = [t.baseline_finish for t in leaves if t.has_baseline and t.baseline_finish]
    cur = [t.finish for t in leaves if t.finish]
    out.baseline_finish = max(bl) if bl else None
    out.finish = max(cur) if cur else None
    out.status = status or (project.status_date.date() if project.status_date else None)
    out.baseline_label = project.baseline_label
    out.errors = sum(1 for i in issues if i.severity == SEV_ERROR)
    out.warnings = sum(1 for i in issues if i.severity == SEV_WARN)
    out.infos = sum(1 for i in issues if i.severity == SEV_INFO)
    return out


def facts_table(items: list[Facts]) -> list[tuple[str, list[str]]]:
    """Filas (dato, valor de cada revisión) del resumen y de la comparación de planificaciones."""
    def bl(f: Facts) -> str:
        if not f.baselines:
            return "-"
        return " · ".join(f"{'LB' if n == 0 else f'LB{n}'}: {fmt_number(c, 0)}" for n, c in f.baselines.items())

    n = lambda v: fmt_number(v, 0)      # noqa: E731
    return [
        ("Tareas de la planificación (activas)", [f"{n(f.tasks)} ({n(f.active)})" for f in items]),
        ("Tareas inactivas", [n(f.inactive) for f in items]),
        ("Tareas activas con Visible = No", [n(f.hidden) for f in items]),
        ("Tareas revisadas", [n(f.reviewed) for f in items]),
        ("Filas del libro (de ingreso)", [f"{n(f.rows)} ({n(f.inputs)})" for f in items]),
        ("Filas «Otras tareas» (-D)", [n(f.residual) for f in items]),
        ("Tareas sumadas en la fila de su elemento superior", [n(f.aggregated) for f in items]),
        ("Tareas con comienzo en cada línea base", [bl(f) for f in items]),
        ("Tareas resumen con la línea base sin resumir", [n(f.not_rolled_up) for f in items]),
        ("Vínculos entre sus tareas", [n(f.links_in) for f in items]),
        ("Vínculos con otras planificaciones", [n(f.links_out) for f in items]),
        ("Tareas sin terminar que son críticas", [f.critical for f in items]),
        ("Fecha de estado", [_day(f.status) for f in items]),
        ("Fin de la línea base / fin previsto", [f"{_day(f.baseline_finish)} / {_day(f.finish)}" for f in items]),
        ("Errores / advertencias / informativas", [f"{f.errors} / {f.warnings} / {f.infos}" for f in items]),
    ]


def _row_dates(row: ReportRow) -> tuple[date | None, date | None, date | None]:
    finishes = [t.finish for t in row.leaves if t.finish] if row.is_input else []
    if not row.is_input:
        stack = list(row.children)
        while stack:
            c = stack.pop()
            finishes += [t.finish for t in c.leaves if t.finish]
            stack += c.children
    return (row.start.date() if row.start else None, row.finish.date() if row.finish else None,
            max(finishes).date() if finishes else None)


def _short(codes: list[str], limit: int = 10) -> str:
    return ", ".join(codes[:limit]) + (f" y {len(codes) - limit} más" if len(codes) > limit else "")


def compare(units: list[Unit]) -> list[str]:
    """Diferencias entre las planificaciones revisadas (una parte cada una): códigos EDT de las filas del libro que
    están solo en una, y filas con el mismo código con otra línea base (fechas de plan) u otro fin previsto."""
    rows = [{code_key(r.code): r for r in u.parts[0].rows if r.kind != KIND_RESIDUAL} for u in units]
    out: list[str] = []
    for k, (u, mine) in enumerate(zip(units, rows)):
        others = set().union(*(set(r) for j, r in enumerate(rows) if j != k))
        only = [mine[c].code for c in mine if c not in others]
        out.append(f"Filas solo en {u.title}: {len(only)}" + (f" ({_short(only)})" if only else ""))
    first, *rest = rows
    for j, other in enumerate(rest, start=1):
        common = [c for c in first if c in other]
        baseline = [first[c].code for c in common if _row_dates(first[c])[:2] != _row_dates(other[c])[:2]]
        forecast = [first[c].code for c in common if _row_dates(first[c])[2] != _row_dates(other[c])[2]]
        pair = f"{units[0].title} y {units[j].title}"
        out.append(f"Filas comunes a {pair}: {len(common)}; con otras fechas de plan (línea base): {len(baseline)}"
                   + (f" ({_short(baseline)})" if baseline else ""))
        out.append(f"Filas comunes a {pair} con otro fin previsto: {len(forecast)}"
                   + (f" ({_short(forecast)})" if forecast else ""))
    return out


# ----------------------------------------------------------------------------
# Líneas base
# ----------------------------------------------------------------------------
# Una línea base que no usa el % planificado y está en menos de esta fracción de las tareas que tienen la usada
# parece guardada por error para unas pocas tareas seleccionadas
FEW_TASKS = 0.10


def _baseline_name(n: int) -> str:
    return "Línea base" if n == 0 else f"Línea base {n}"


@dataclass
class BaselineUse:
    """Una línea base guardada, en las tareas de una planificación (o del proyecto, sin planificación)."""

    number: int
    tasks: list[Task]                   # tareas de la planificación con fechas en esta línea base
    used: bool                          # la usa el % planificado
    saved: datetime | None = None       # fecha de guardado (solo con .mpp)
    start: datetime | None = None       # comienzo más temprano
    finish: datetime | None = None      # fin más tardío
    same: int | None = None             # tareas con las mismas fechas que la línea base usada (None = la usada)
    missing: int = 0                    # la usada: tareas de detalle de la planificación sin ella
    outside: int = 0                    # tareas del archivo, fuera de la planificación, con esta línea base
    signals: list[str] = field(default_factory=list)

    @property
    def summaries(self) -> int:
        return sum(1 for t in self.tasks if t.summary or t.children)


def _plan_tasks(part: ReportPart) -> tuple[list[Task], set[int] | None]:
    """Tareas de la planificación revisada en el árbol completo de MS Project (también las inactivas y las de
    Visible = No) y sus UniqueID (None sin planificación: todas las tareas del proyecto)."""
    from evm.validation import plan_members
    project: ProjectData = part.source.project
    members = plan_members(project, part.scope) if part.scope is not None else None
    return [t for t in project.tasks if t.outline_level >= 1 and (members is None or t.uid in members)], members


def baseline_uses(part: ReportPart) -> list[BaselineUse]:
    """Líneas base guardadas en las tareas de la planificación (y, con .mpp, las que tienen fecha de guardado), con
    sus fechas e indicios de una línea base guardada por error: en pocas tareas, solo en tareas resumen o copia de
    la que usa el % planificado."""
    project: ProjectData = part.source.project
    every, members = _plan_tasks(part)
    n_used = project.baseline_number
    used = _baseline_name(n_used)
    numbers = {k for t in every for k in t.baseline_dates} | set(project.baseline_saved)
    used_dates = {t.uid: t.baseline_dates.get(n_used) for t in every}
    with_used = sum(1 for v in used_dates.values() if v is not None)
    out: list[BaselineUse] = []
    for k in sorted(numbers):
        tasks = [t for t in every if k in t.baseline_dates]
        use = BaselineUse(k, tasks, k == n_used, project.baseline_saved.get(k))
        starts = [t.baseline_dates[k][0] for t in tasks if t.baseline_dates[k][0]]
        finishes = [t.baseline_dates[k][1] for t in tasks if t.baseline_dates[k][1]]
        use.start = min(starts) if starts else None
        use.finish = max(finishes) if finishes else None
        if members is not None:
            use.outside = sum(1 for t in project.tasks if t.uid not in members and t.outline_level >= 1
                              and k in t.baseline_dates)
        if use.used:
            use.missing = sum(1 for t in every if not t.children and k not in t.baseline_dates)
        else:
            use.same = sum(1 for t in tasks if used_dates.get(t.uid) == t.baseline_dates[k])
            if not tasks:
                use.signals.append("guardada, pero ninguna tarea de la planificación tiene sus fechas")
            else:
                base = with_used or len(every)
                if len(tasks) < FEW_TASKS * base:
                    use.signals.append(f"solo en {_tasks(len(tasks))} "
                                       f"({_num(100.0 * len(tasks) / base)} % de las que tienen la {used}): "
                                       "parece guardada por error para unas tareas seleccionadas")
                if use.summaries == len(tasks):
                    use.signals.append("solo en tareas resumen")
                if use.same == len(tasks):
                    use.signals.append(f"mismas fechas que la {used} en todas sus tareas: es una copia")
        out.append(use)
    return out


def _tasks(n: int) -> str:
    """'1 tarea', '1.351 tareas'."""
    return f"{fmt_number(n, 0)} {'tarea' if n == 1 else 'tareas'}"


def _task_label(t: Task, code_of) -> str:
    return code_of(t) or f"UID {t.uid}"


RANGE = 3          # tareas de detalle hermanas seguidas, en la misma situación, que se indican como un rango


class TaskTree:
    """Tareas de una lista (las que tienen una línea base, o las que no tienen la usada) en el esquema de la
    planificación, por su código EDT (sin código, por su UniqueID), para ubicarlas en MS Project:

    - starts(): el comienzo de cada rama, la tarea más alta de cada grupo (su tarea superior no está en la lista),
      con dónde está y qué abarca;
    - detail(): el árbol con sangría, con las tareas de la lista en negrita y, en texto normal, las tareas resumen
      de encima que no están (para ver la jerarquía).

    Una tarea resumen que está con toda su rama en la misma situación va en una línea (rama completa); tres o más
    tareas de detalle hermanas seguidas en la misma situación, como un rango. selected: UniqueID -> situación (texto
    que se agrega a la tarea; vacío si no hace falta). verb: lo que hacen las subtareas de la lista («la tienen»).
    context: texto de una tarea resumen de encima que no está en la lista."""

    def __init__(self, every: list[Task], code_of, selected: dict[int, str], verb: str,
                 context: Callable[[Task], str]):
        self.code_of, self.selected, self.verb, self.context = code_of, selected, verb, context
        inside = {t.uid for t in every}
        self.children: dict[int, list[Task]] = {}
        for t in every:
            if t.parent is not None and t.parent.uid in inside:
                self.children.setdefault(t.parent.uid, []).append(t)
        self.roots = [t for t in every if t.parent is None or t.parent.uid not in inside]
        self.size: dict[int, int] = {}      # tareas de la rama (con la propia)
        self.count: dict[int, int] = {}     # tareas de la rama que están en la lista
        self.full: dict[int, bool] = {}     # toda la rama está en la lista, en la misma situación
        for t in reversed(every):           # las subtareas van después de su tarea resumen
            kids = self.children.get(t.uid, [])
            self.size[t.uid] = 1 + sum(self.size[c.uid] for c in kids)
            self.count[t.uid] = (t.uid in selected) + sum(self.count[c.uid] for c in kids)
            self.full[t.uid] = t.uid in selected and all(
                self.full[c.uid] and selected[c.uid] == selected[t.uid] for c in kids)

    def code(self, t: Task) -> str:
        return f"`{_task_label(t, self.code_of)}`"

    def _share(self, t: Task, k: int) -> str:
        """«de sus 12 subtareas, 9 la tienen» (verb en plural; en singular, sin la n final)."""
        n = self.size[t.uid] - 1
        one = self.verb[:-1] if self.verb.endswith("n") else self.verb
        if k == 0:              # solo en la lista de una línea base: una tarea resumen que la tiene sin subtareas
            return f"su única subtarea no {one}" if n == 1 else f"ninguna de sus {fmt_number(n, 0)} subtareas {one}"
        of = "de su única subtarea" if n == 1 else f"de sus {fmt_number(n, 0)} subtareas"
        return f"{of}, {fmt_number(k, 0)} {one if k == 1 else self.verb}"

    def _what(self, t: Task) -> str:
        """Qué abarca una tarea de la lista."""
        if t.uid not in self.children:
            return ""
        if self.full[t.uid]:
            return f"rama completa ({_tasks(self.size[t.uid])})"
        return f"tarea resumen; {self._share(t, self.count[t.uid] - 1)}"

    def _tail(self, *parts: str) -> str:
        parts = tuple(p for p in parts if p)
        return f" — {'; '.join(parts)}" if parts else ""

    def _segments(self, kids: list[Task], pick: Callable[[Task], bool]) -> list[list[Task]]:
        """Tareas hermanas en grupos: RANGE o más tareas de detalle seguidas que cumplen pick, en la misma situación,
        forman un rango; las demás van solas (también las que no cumplen pick, que cortan el rango)."""
        out: list[list[Task]] = []
        run: list[Task] = []

        def flush() -> None:
            out.extend([list(run)] if len(run) >= RANGE else [[t] for t in run])
            run.clear()

        for t in kids:
            leaf = pick(t) and t.uid not in self.children
            if leaf and run and self.selected[run[0].uid] != self.selected[t.uid]:
                flush()
            if leaf:
                run.append(t)
                continue
            flush()
            out.append([t])
        flush()
        return out

    def starts(self) -> list[str]:
        lines: list[str] = []

        def walk(kids: list[Task], parent: Task | None) -> None:
            inside = parent is not None and parent.uid in self.selected
            for seg in self._segments(kids, lambda t: not inside and t.uid in self.selected):
                t = seg[0]
                if not inside and t.uid in self.selected:
                    # Una tarea superior sin código EDT (p. ej. la tarea resumen del programa) no ayuda a ubicarla
                    where = (f"dentro de {self.code(parent)}" if parent is not None and self.code_of(parent) else
                             f"nivel {t.outline_level}")
                    note = self.selected[t.uid]
                    if len(seg) > 1:
                        lines.append(f"- {self.code(t)} a {self.code(seg[-1])}"
                                     f"{self._tail(where, _tasks(len(seg)) + ' de detalle', note)}")
                    else:
                        lines.append(f"- {self.code(t)}"
                                     f"{self._tail(where, self._what(t) or 'tarea de detalle', note)}")
                if len(seg) == 1 and self.count[t.uid] and t.uid in self.children and not self.full[t.uid]:
                    walk(self.children[t.uid], t)

        walk(self.roots, None)
        return lines

    def detail(self) -> list[str]:
        lines: list[str] = []

        def walk(kids: list[Task], depth: int) -> None:
            pad = "  " * depth
            for seg in self._segments(kids, lambda t: t.uid in self.selected):
                t = seg[0]
                if not self.count[t.uid]:
                    continue
                if len(seg) > 1:
                    lines.append(f"{pad}- **{self.code(t)}** a **{self.code(seg[-1])}** ({_tasks(len(seg))})"
                                 f"{self._tail(self.selected[t.uid])}")
                    continue
                if t.uid in self.selected:
                    lines.append(f"{pad}- **{self.code(t)}**{self._tail(self._what(t), self.selected[t.uid])}")
                    if self.full[t.uid]:
                        continue
                else:
                    lines.append(f"{pad}- {self.code(t)} — {self.context(t)}; {self._share(t, self.count[t.uid])}")
                walk(self.children.get(t.uid, []), depth + 1)

        walk(self.roots, 0)
        return lines


def _baseline_section(part: ReportPart, level: str) -> list[str]:
    """Análisis de las líneas base de una revisión: tabla con sus datos e indicios, y las tareas de cada una."""
    project: ProjectData = part.source.project
    uses = baseline_uses(part)
    if not uses:
        return ["Ninguna tarea tiene línea base guardada: el % planificado usa las fechas actuales.", ""]
    every, members = _plan_tasks(part)
    code_of = part.code_of or (lambda t: t.wbs or "")
    lines = ["| Línea base | Uso | Guardada | Tareas (resumen) | Comienzo más temprano | Fin más tardío | "
             "Mismas fechas que la usada | Indicios |", "|---|---|---|---:|---|---|---:|---|"]
    for u in uses:
        saved = _day(u.saved) if u.saved else ("-" if project.backend == "com" else "no consta en el XML")
        tasks = f"{fmt_number(len(u.tasks), 0)} ({fmt_number(u.summaries, 0)})"
        same = "-" if u.same is None else fmt_number(u.same, 0)
        if u.used:
            use = "% planificado"
            signals = f"{fmt_number(u.missing, 0)} {'tarea' if u.missing == 1 else 'tareas'} de detalle sin ella" if u.missing else "-"
        else:
            use = "-"
            signals = "; ".join(u.signals) if u.signals else "-"
        lines.append(f"| {_baseline_name(u.number)} | {use} | {saved} | {tasks} | {_day(u.start)} | "
                     f"{_day(u.finish)} | {same} | {_md(signals)} |")
    lines.append("")
    outside = [u for u in uses if u.outside]
    if members is not None and outside:
        lines += ["Fuera de la planificación, en otras tareas del archivo: "
                  + ", ".join(f"{_baseline_name(u.number)}: {_tasks(u.outside)}" for u in outside)
                  + ".", ""]
    for u in uses:
        name = _baseline_name(u.number)
        lines += [f"{level} {name}", ""]
        have = {t.uid for t in u.tasks}
        if u.used:                      # la que usa el % planificado: solo cuántas tareas la tienen
            lines += [f"{_tasks(len(have))} de {fmt_number(len(every), 0)} (la usa el % "
                      f"planificado). Las que no la tienen están al final, en «Tareas sin la {name}».", ""]
        elif not have:
            lines += ["Ninguna tarea de la planificación tiene sus fechas.", ""]
        else:
            tree = TaskTree(every, code_of, {uid: "" for uid in have}, "la tienen", lambda t: "no la tiene")
            summaries = "ninguna resumen" if not u.summaries else f"{fmt_number(u.summaries, 0)} resumen"
            lines += [f"{_tasks(len(have))} {'tiene' if len(have) == 1 else 'tienen'} la {name} ({summaries}).", "",
                      "**Comienzo de cada rama** (la tarea más alta de cada grupo: su tarea superior no la tiene):",
                      ""] + tree.starts() + [
                      "", "**Detalle** (en negrita, las tareas que la tienen; en texto normal, las tareas resumen de "
                      "encima que no la tienen):", ""] + tree.detail() + [""]
    return lines + _without_baseline(project, every, code_of, level)


def _without_baseline(project: ProjectData, every: list[Task], code_of, level: str) -> list[str]:
    """Tareas de la planificación sin la línea base que usa el % planificado, con las otras líneas base que tienen
    (ninguna, o solo otras): el comienzo de cada rama y el detalle (TaskTree). Están las de detalle y las tareas
    resumen sin ella en toda su rama; una tarea resumen sin ella con alguna subtarea que la tiene no está (sus
    fechas no se usan): se ve en el detalle como tarea de encima."""
    n = project.baseline_number
    used = _baseline_name(n)
    inside = {t.uid for t in every}
    children: dict[int, list[Task]] = {}
    for t in every:
        if t.parent is not None and t.parent.uid in inside:
            children.setdefault(t.parent.uid, []).append(t)
    lacking: dict[int, bool] = {}       # UniqueID -> ni la tarea ni ninguna de sus subtareas tiene la línea base
    for t in reversed(every):           # las subtareas van después de su tarea resumen
        lacking[t.uid] = n not in t.baseline_dates and all(lacking[c.uid] for c in children.get(t.uid, []))
    leaves = [t for t in every if t.uid not in children and lacking[t.uid]]
    lines = [f"{level} Tareas sin la {used}", ""]
    if not any(lacking.values()):
        return lines + [f"Todas las tareas de detalle de la planificación tienen la {used}.", ""]

    def situation(t: Task) -> str:
        others = sorted(t.baseline_dates)
        if not others:
            return "sin ninguna línea base"
        if len(others) == 1:
            return f"solo con la {_baseline_name(others[0])}"
        return f"solo con las líneas base {', '.join(str(k) for k in others[:-1])} y {others[-1]}"

    # La situación más frecuente se indica una vez, antes de las listas; en cada tarea, solo las demás
    selected = {t.uid: situation(t) for t in every if lacking[t.uid]}
    common = max(sorted(set(selected.values())), key=lambda s: sum(1 for t in leaves if selected[t.uid] == s))
    selected = {uid: ("" if s == common else s) for uid, s in selected.items()}
    none = sum(1 for t in leaves if not t.baseline_dates)
    tree = TaskTree(every, code_of, selected, "no la tienen",
                    lambda t: "la tiene" if n in t.baseline_dates else
                    "no la tiene, pero alguna subtarea sí (sus fechas no se usan)")
    lines += [f"{_tasks(len(leaves))} de detalle sin la {used}, que usa el % planificado: {fmt_number(none, 0)} sin "
              f"ninguna línea base y {fmt_number(len(leaves) - none, 0)} solo con otras. No forman parte del PV "
              "(cuentan en el EV como trabajo no previsto). Si su trabajo está aprobado y falta solo guardarlo, la "
              f"línea base se establece con `corregir-linea-base --tarea CODIGO --establecer {n}` (con `--rama` "
              "para una rama completa) o en MS Project (*Establecer línea base*, *Tareas seleccionadas*).", "",
              f"Salvo que se indique otra cosa, las tareas de las listas están {common}.", "",
              f"**Comienzo de cada rama** (la tarea más alta de cada grupo: su tarea superior tiene la {used}, o "
              "alguna subtarea con ella):", ""]
    lines += tree.starts()
    lines += ["", f"**Detalle** (en negrita, las tareas sin la {used}; en texto normal, las tareas resumen de encima "
              "que la tienen o que tienen alguna subtarea con ella):", ""]
    return lines + tree.detail() + [""]


# ----------------------------------------------------------------------------
# Sin nombres
# ----------------------------------------------------------------------------
def sanitizer(projects: list[ProjectData], cfg: Config, keep: set[str]) -> Callable[[str], str]:
    """Quita de un texto los nombres de las tareas, del proyecto, de los calendarios de tarea y de las empresas: lo
    que va entre «» y cada nombre que aparece como palabra completa. keep: textos que se conservan (p. ej. los
    códigos de las planificaciones, que pueden ser también el nombre de su tarea)."""
    names = {t.name.strip() for p in projects for t in p.tasks if t.name}
    names |= {p.name.strip() for p in projects if p.name}
    names |= {t.calendar for p in projects for t in p.tasks if t.calendar}
    companies = {}
    if cfg.multi_enabled:
        companies = {cfg.company_a: "empresa A", cfg.company_b: "empresa B", cfg.company_both: "ambas empresas"}
    names = {n for n in names if len(n) >= 4 and n not in keep and n not in companies}
    pattern = (re.compile(r"(?<!\w)(" + "|".join(re.escape(n) for n in sorted(names, key=len, reverse=True))
                          + r")(?!\w)") if names else None)

    def clean(text: str) -> str:
        text = re.sub(r"«[^»]*»", "«…»", text)
        for name, generic in companies.items():
            if name:
                text = re.sub(r"(?<!\w)" + re.escape(name) + r"(?!\w)", generic, text)
        return pattern.sub("…", text) if pattern is not None else text

    return clean


# ----------------------------------------------------------------------------
# Informe Markdown
# ----------------------------------------------------------------------------
def _where(found: list[Issue], limit: int = 20) -> str:
    """Tareas de las observaciones: su código EDT o, sin código, su UniqueID."""
    names = list(dict.fromkeys((i.code or (f"UID {i.uid}" if i.uid is not None else "")) for i in found))
    names = [n for n in names if n]
    return ", ".join(f"`{n}`" for n in names[:limit]) + (f" y {len(names) - limit} más" if len(names) > limit else "")


def _md(text: str) -> str:
    """Texto del catálogo o de una observación para una tabla o una lista de Markdown."""
    return text.replace("|", "\\|").replace("\n", " ")


def _where_lines(tasks: list[Issue]) -> list[str]:
    """Líneas «Dónde» de un código: una por caso (variant) si el código tiene varios casos (p. ej. LOG-003: sin
    predecesoras, sin sucesoras o sin ninguna)."""
    variants = list(dict.fromkeys(i.variant for i in tasks))
    if variants == [""]:
        return [f"- **Dónde:** {_where(tasks)}"]
    out = []
    for v in variants:
        group = [i for i in tasks if i.variant == v]
        out.append(f"- **Dónde ({v or 'otros casos'}, {plural(len(group), 'tarea', 'tareas')}):** {_where(group)}")
    return out


def _improvements(issues: list[Issue], clean: Callable[[str], str], level: str) -> list[str]:
    out: list[str] = []
    by_priority: dict[str, dict[str, list[Issue]]] = {p: {} for p in PRIORITIES}
    for i in issues:
        if i.id not in FACT_CODES:
            by_priority[priority(i)].setdefault(i.id, []).append(i)
    for p in PRIORITIES:
        codes = by_priority[p]
        out += [f"{level} Prioridad {p.lower()}", ""]
        if not codes:
            out += ["Sin observaciones.", ""]
            continue
        out += [f"*{PRIORITY_TEXT[p][0].upper() + PRIORITY_TEXT[p][1:]}.*", ""]
        # Los errores primero; después, por código (cada código reúne su resumen y sus tareas)
        order = sorted(codes.items(), key=lambda kv: (kv[1][0].severity != SEV_ERROR, kv[0]))
        for code, found in order:
            entry = CATALOG.get(code)
            title = entry.title if entry else code
            tasks = [i for i in found if i.code or i.uid is not None]
            general = [i for i in found if not (i.code or i.uid is not None)]
            out.append(f"{level}# {code} · {title}")
            out.append("")
            severity = found[0].severity
            count = f"{len(tasks)} {'tarea' if len(tasks) == 1 else 'tareas'}" if tasks else "general"
            out.append(f"- **{severity}** · {count}")
            if tasks:
                out += _where_lines(tasks)
            for i in general[:3]:
                out.append(f"- **Observación:** {_md(clean(i.message))}")
            if entry is not None:
                out.append(f"- **Causa:** {_md(entry.cause)}")
                out.append(f"- **Qué hacer:** {_md(entry.action)}")
            out.append("")
    return out


def _dcma_cell(value: tuple[str, str, bool | None] | None) -> str:
    if value is None:
        return "-"
    text, _threshold, ok = value
    mark = "" if ok is None else (" · cumple" if ok else " · **no cumple**")
    return _md(text) + mark


def dcma_rows(parts: list[ReportPart]) -> list[tuple[int, str, list[tuple[str, str, bool | None] | None], str]]:
    """Filas de la evaluación DCMA: (punto, nombre, valor de cada revisión, umbral)."""
    from evm.validation import DCMA_POINTS, DCMA_THRESHOLDS
    return [(k, DCMA_POINTS[k], [p.dcma.get(k) for p in parts], DCMA_THRESHOLDS[k]) for k in sorted(DCMA_POINTS)]


def write_report(path: str | Path, units: list[Unit], cfg: Config, version: str, engine: str) -> None:
    """Escribe el informe de mejoras (Markdown, UTF-8)."""
    projects = [p.source.project for u in units for p in u.parts]
    keep = {u.title for u in units} | {p.source.plan for u in units for p in u.parts if p.source.plan}
    clean = sanitizer(projects, cfg, keep)
    columns = [(u, p) for u in units for p in u.parts]
    labels = [u.title if len(u.parts) == 1 else (p.source.name or u.title) for u, p in columns]
    all_facts = [facts(p, [i for i in u.issues if not i.file or i.file == p.source.name], label)
                 for (u, p), label in zip(columns, labels)]

    lines = ["# Informe de mejoras del cronograma", "",
             "| | |", "|---|---|",
             f"| **Fecha** | {datetime.now():%d/%m/%Y %H:%M} |",
             f"| **Versión** | Exportador EVM {version} · `validar --informe` |",
             f"| **Archivo** | {engine} |",
             f"| **Revisado** | {', '.join(labels)} |",
             f"| **Línea base** | {projects[0].baseline_label} |", "",
             "> El informe no incluye nombres de tareas, de empresas ni de calendarios, ni el título del proyecto: las "
             "tareas se identifican por su código EDT o, sin código, por su UniqueID (UID). Los textos entre «» se "
             "omiten (`«…»`).", "",
             "---", "", "## 1. Resumen", "",
             "| Dato | " + " | ".join(labels) + " |", "|---|" + "---:|" * len(labels)]
    for name, values in facts_table(all_facts):
        lines.append(f"| {name} | " + " | ".join(values) + " |")
    lines.append("")
    # Por código: tareas (las observaciones de una tarea) y observaciones generales
    counts: dict[tuple[str, str], list[int]] = {}
    for u in units:
        for i in u.issues:
            if priority(i) == PRIORITY_HIGH and i.id not in FACT_CODES:
                c = counts.setdefault((u.title, i.id), [0, 0])
                c[0 if (i.code or i.uid is not None) else 1] += 1
    if counts:
        lines += ["**Mejoras de prioridad alta** (las que más tareas u observaciones tienen):", ""]
        for (title, code), (tasks, general) in sorted(counts.items(), key=lambda kv: -(kv[1][0] or kv[1][1]))[:8]:
            entry = CATALOG.get(code)
            where = f" en {title}" if len(units) > 1 else ""
            what = plural(tasks, "tarea", "tareas") if tasks else plural(general, "observación", "observaciones")
            lines.append(f"- `{code}` {entry.title if entry else ''}: {what}{where}.")
        lines.append("")

    lines += ["---", "", "## 2. Mejoras por prioridad", "",
              "Cada código se explica en [Códigos de errores y advertencias]"
              "(https://github.com/MarceloSalazar/evm-automation/blob/main/docs/codigos.md). Las observaciones "
              f"informativas de la estructura ({', '.join(f'`{c}`' for c in FACT_CODES)}) no se repiten aquí: sus "
              "datos están en el resumen. Los índices DCMA (% y umbrales) están en la evaluación DCMA (sección 3).",
              ""]
    for u in units:
        level = "###"
        if len(units) > 1:
            lines += [f"### {u.title}", ""]
            level = "####"
        lines += _improvements(u.issues, clean, level)

    lines += ["---", "", "## 3. Evaluación DCMA de 14 puntos", "",
              "| Punto | " + " | ".join(labels) + " | Umbral |", "|---|" + "---|" * len(labels) + "---|"]
    for k, name, values, threshold in dcma_rows([p for _u, p in columns]):
        lines.append(f"| {k}. {name} | " + " | ".join(_dcma_cell(v) for v in values) + f" | {threshold} |")
    lines.append("")

    lines += ["---", "", "## 4. Líneas base", "",
              "Líneas base guardadas en las tareas de la planificación (también las inactivas y las de Visible = No) "
              "y, de cada una salvo la que usa el % planificado, las tareas que la tienen, por su código EDT. Sirve "
              "para detectar las guardadas por error: en pocas tareas (p. ej. con *Tareas seleccionadas* y otra "
              "línea base), solo en tareas resumen o copia de la que usa el % planificado. Al final, las tareas sin "
              "la línea base que usa el % planificado, con las otras líneas base que tienen.\n\n"
              "Las tareas de cada lista se indican en dos partes. **Comienzo de cada rama:** la tarea más alta de "
              "cada grupo (la que se busca en MS Project), con dónde está y qué abarca. **Detalle:** el esquema con "
              "sangría, con las tareas de la lista en negrita y, en texto normal, las tareas resumen de encima que no "
              "están en ella, para ver la jerarquía. Una tarea resumen con toda su rama en la lista va en una línea "
              "(*rama completa*); tres o más tareas de detalle hermanas seguidas, como un rango (*`1.1.3` a "
              "`1.1.11`*). Una línea base que sobra "
              "se borra en MS Project con *Proyecto > Establecer línea base > Borrar línea base* (*Clear Baseline*), "
              "para todo el proyecto o para las tareas seleccionadas.", ""]
    for (_u, p), label in zip(columns, labels):
        level = "###"
        if len(columns) > 1:
            lines += [f"### {label}", ""]
            level = "####"
        lines += _baseline_section(p, level)

    if len(units) > 1:
        lines += ["---", "", "## 5. Comparación de planificaciones", ""]
        lines += [f"- {_md(clean(text))}" for text in compare(units)]
        lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")
