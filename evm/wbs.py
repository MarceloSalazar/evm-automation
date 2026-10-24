"""Construcción de la estructura EDT del informe (hasta el nivel máximo) y
cálculo del PV mensual de cada fila."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

from evm.config import MODE_COST, MODE_PERCENT, Config, ConfigError
from evm.model import MonthKey, ProjectData, Task
from evm.timeline import Horizon, add_into, distribute_linear, first_of_month, lump

ACCRUE_START, ACCRUE_PRORATED, ACCRUE_END = "inicio", "prorrateado", "fin"

SEV_ERROR = "Error"
SEV_WARN = "Advertencia"
SEV_INFO = "Info"

KIND_LEAF = "hoja"          # fila de ingreso de datos (paquete de trabajo / cuenta de control)
KIND_SUMMARY = "resumen"    # fila subtotal (fórmulas)
KIND_RESIDUAL = "residual"  # costo asignado directamente a una tarea resumen


@dataclass
class Issue:
    severity: str
    code: str
    name: str
    message: str


@dataclass
class ReportRow:
    code: str
    name: str
    level: int
    kind: str
    task: Task
    method: str = ""
    bac: float = 0.0
    start: datetime | None = None
    finish: datetime | None = None
    pv: dict[MonthKey, float] = field(default_factory=dict)
    parent: "ReportRow | None" = field(default=None, repr=False)
    children: list["ReportRow"] = field(default_factory=list, repr=False)
    vector: list[float] = field(default_factory=list, repr=False)
    pv_before: float = 0.0
    pv_after: float = 0.0
    hidden_tasks: int = 0       # tareas de nivel inferior agregadas en esta fila
    # Multi-empresa: parte de la empresa A (la de B es el total menos A)
    pv_a: dict[MonthKey, float] = field(default_factory=dict)
    vector_a: list[float] = field(default_factory=list, repr=False)
    bac_a: float = 0.0
    source: str = ""            # varios archivos: nombre del archivo de la tarea ([[proyectos]])
    # Modo porcentaje: tareas de detalle que forman la fila y % planificado de cada empresa
    # (clave "" en un libro de una sola empresa; "A" y "B" en multi-empresa)
    leaves: list[Task] = field(default_factory=list, repr=False)
    pct: dict[str, dict[MonthKey, float]] = field(default_factory=dict, repr=False)
    pct_vec: dict[str, list[float]] = field(default_factory=dict, repr=False)

    @property
    def is_input(self) -> bool:
        return self.kind in (KIND_LEAF, KIND_RESIDUAL)


@dataclass
class Source:
    """Archivo de MS Project incluido en el informe."""

    name: str                   # nombre del archivo en [[proyectos]]; "" = informe de un solo archivo
    project: ProjectData
    plan: str = ""              # código EDT de la planificación exportada ("" = todo el proyecto)
    pv_source: str = ""         # descripción de la fuente del PV mensual


@dataclass
class Report:
    project: ProjectData        # varios archivos: el primero (los datos de cada uno están en sources)
    horizon: Horizon
    rows: list[ReportRow]
    roots: list[ReportRow]
    issues: list[Issue]
    total_vector: list[float]
    total_bac: float
    pv_before: float
    pv_after: float
    pv_source: str = ""
    plan: str = ""              # código EDT de la planificación exportada ("" = todo el proyecto)
    multi: bool = False         # multi-empresa activado
    company_a: str = ""
    company_b: str = ""
    sources: list[Source] = field(default_factory=list)
    title: str = ""             # título del informe
    currency: str = ""          # moneda de MS Project (se usa si la configuración no indica 'moneda')
    mode: str = MODE_COST       # "porcentaje": % de MS Project × BAC del libro; "costo": costo de línea base
    # Depuración de costes: hoja PV que lleva las filas «(coste)» ("A" o "B" en multi-empresa, "" en un libro de una
    # sola empresa); None = sin depuración
    cost_review: str | None = None

    @property
    def percent(self) -> bool:
        return self.mode == MODE_PERCENT

    @property
    def multi_file(self) -> bool:
        """True si el informe combina archivos de [[proyectos]] (las claves de las filas llevan su nombre)."""
        return any(s.name for s in self.sources)


@dataclass
class ReportPart:
    """Filas de un archivo de MS Project, antes de combinar los archivos y fijar el horizonte."""

    source: Source
    rows: list[ReportRow]
    roots: list[ReportRow]
    issues: list[Issue]         # observaciones de la lectura y de la estructura EDT
    checks: list[Issue]         # controles finales (tareas omitidas, BAC, códigos duplicados...)
    first_date: date | None     # inicio del horizonte si ninguna fila tiene PV


# ----------------------------------------------------------------------------
# Fuentes de PV
# ----------------------------------------------------------------------------
class PVSource:
    """Obtiene la distribución mensual del costo de línea base de una tarea."""

    description = ""

    def leaf(self, task: Task) -> dict[MonthKey, float]:
        raise NotImplementedError

    def summary_residual(self, task: Task, amount: float,
                         child_dists: list[dict[MonthKey, float]] | None = None) -> dict[MonthKey, float]:
        raise NotImplementedError

    def subtree(self, task: Task) -> dict[MonthKey, float]:
        """Distribución de toda la rama (se usa para tareas excluidas)."""
        return distribute_linear(task.baseline_cost, task.baseline_start or task.start,
                                 task.baseline_finish or task.finish)


class LinearPVSource(PVSource):
    """Distribución lineal entre inicio y fin de la línea base."""

    def __init__(self, mode: str = "habiles", holidays: set[date] | None = None, calendar=None):
        from evm.workcalendar import WorkCalendar
        self.mode = mode
        self.calendar = calendar or WorkCalendar()
        self.calendar.add_holidays(holidays or ())
        self.issues: list[Issue] = []
        self.linear_count = 0          # tareas repartidas linealmente
        self.timephased_count = 0      # tareas con datos escalonados de MS Project
        self.description = (
            f"Distribución lineal por tiempo laborable (calendario {self.calendar.name}) entre inicio y fin "
            "de la línea base" if mode == "habiles"
            else "Distribución lineal por días corridos entre inicio y fin de la línea base"
        )

    def _dates(self, task: Task) -> tuple[datetime | None, datetime | None]:
        if task.baseline_start or task.baseline_finish:
            return task.baseline_start, task.baseline_finish
        return task.start, task.finish

    def linear_leaf(self, task: Task) -> dict[MonthKey, float]:
        amount = task.baseline_cost
        if not amount:
            return {}
        self.linear_count += 1
        start, finish = self._dates(task)
        if task.milestone:
            return lump(amount, finish or start)
        # Costo fijo con acumulación al comienzo / al final (sin costo de recursos)
        only_fixed = bool(task.fixed_cost) and abs(task.cost - task.fixed_cost) <= 0.005 * max(1.0, abs(task.cost))
        if only_fixed and task.fixed_cost_accrual == ACCRUE_START:
            return lump(amount, start)
        if only_fixed and task.fixed_cost_accrual == ACCRUE_END:
            return lump(amount, finish)
        return distribute_linear(amount, start, finish, self.mode, self.calendar)

    def leaf(self, task: Task) -> dict[MonthKey, float]:
        return self.linear_leaf(task)

    def summary_residual(self, task: Task, amount: float,
                         child_dists: list[dict[MonthKey, float]] | None = None) -> dict[MonthKey, float]:
        start, finish = self._dates(task)
        return distribute_linear(amount, start, finish, self.mode, self.calendar)

    def subtree(self, task: Task) -> dict[MonthKey, float]:
        start, finish = self._dates(task)
        return distribute_linear(task.baseline_cost, start, finish, self.mode, self.calendar)


class TimephasedPVSource(LinearPVSource):
    """Usa la distribución mensual del costo de línea base calculada por MS Project
    (datos escalonados en el tiempo). Si una tarea no la tiene, reparte en forma lineal.

    `fetch(task)` devuelve {(año, mes): monto} o None.
    """

    def __init__(self, fetch, description: str, mode: str = "habiles", holidays: set[date] | None = None,
                 calendar=None):
        self._linear_description = ""
        super().__init__(mode, holidays, calendar)
        self.fetch = fetch
        self.base_description = description

    @property
    def description(self) -> str:  # type: ignore[override]
        if not self.timephased_count:
            return self._linear_description
        text = self.base_description
        if self.linear_count:
            text += f" ({plural(self.linear_count, 'tarea', 'tareas')} sin datos mensuales: reparto lineal)"
        return text

    @description.setter
    def description(self, value: str) -> None:
        self._linear_description = value

    def leaf(self, task: Task) -> dict[MonthKey, float]:
        amount = task.baseline_cost
        if not amount:
            return {}
        tp = self.fetch(task) or {}
        total = sum(tp.values())
        if abs(total) < 1e-9:
            return self.linear_leaf(task)
        if total * amount < 0:
            # No se puede ajustar en forma proporcional: el ajuste invertiría el signo de cada mes
            self.issues.append(Issue(
                SEV_WARN, task.wbs, task.name,
                f"La distribución mensual de la línea base en MS Project suma {fmt_money(total)} y el costo de "
                f"línea base es {fmt_money(amount)} (signo opuesto). El costo de línea base se repartió en forma "
                "lineal entre el inicio y el fin de la línea base."))
            return self.linear_leaf(task)
        self.timephased_count += 1
        if abs(total - amount) > tolerance(amount):
            self.issues.append(Issue(
                SEV_WARN, task.wbs, task.name,
                f"La distribución mensual de la línea base en MS Project suma {fmt_money(total)} y el costo de "
                f"línea base es {fmt_money(amount)} (¿costo de línea base editado a mano?). Se ajustó en forma "
                "proporcional al costo de línea base."))
            factor = amount / total
            tp = {k: v * factor for k, v in tp.items()}
        return dict(tp)

    def summary_residual(self, task: Task, amount: float,
                         child_dists: list[dict[MonthKey, float]] | None = None) -> dict[MonthKey, float]:
        """Costo propio de la tarea resumen = su distribución menos la de sus hijos."""
        own = self.fetch(task)
        if own:
            diff = dict(own)
            for d in (child_dists if child_dists is not None else [self.fetch(c) or {} for c in task.children]):
                add_into(diff, d, -1.0)
            diff = {k: v for k, v in diff.items() if abs(v) > 0.005}
            if diff and abs(sum(diff.values()) - amount) <= tolerance(amount) and min(diff.values()) > -0.005:
                return diff
        return super().summary_residual(task, amount)

    def subtree(self, task: Task) -> dict[MonthKey, float]:
        tp = self.fetch(task)
        if tp and abs(sum(tp.values())) > 1e-9:
            return dict(tp)
        return super().subtree(task)


# ----------------------------------------------------------------------------
# Árbol de tareas
# ----------------------------------------------------------------------------
def build_tree(tasks: list[Task]) -> list[Task]:
    """Arma la jerarquía a partir del orden de las tareas y su nivel de esquema."""
    roots: list[Task] = []
    stack: list[Task] = []
    for t in tasks:
        t.parent = None
        t.children = []
        while stack and stack[-1].outline_level >= t.outline_level:
            stack.pop()
        if stack:
            t.parent = stack[-1]
            stack[-1].children.append(t)
        else:
            roots.append(t)
        stack.append(t)
    return roots


def iter_subtree(task: Task):
    yield task
    for c in task.children:
        yield from iter_subtree(c)


def prune_inactive(roots: list[Task]) -> tuple[list[Task], dict[int, list[Task]], int]:
    """Quita las tareas inactivas (con toda su rama).

    Devuelve (raíces activas, {uid del padre: hijos excluidos}, cantidad de tareas excluidas).
    """
    removed_children: dict[int, list[Task]] = {}
    count = 0

    def prune(nodes: list[Task], parent: Task | None) -> list[Task]:
        nonlocal count
        kept = []
        for n in nodes:
            if not n.active:
                count += sum(1 for _ in iter_subtree(n))
                if parent is not None:
                    removed_children.setdefault(parent.uid, []).append(n)
                continue
            n.children = prune(n.children, n)
            kept.append(n)
        return kept

    return prune(roots, None), removed_children, count


def pad_code(code: str, levels: int) -> str:
    """Completa un código EDT numérico con ceros: '1' -> '1.0.0', '1.2' -> '1.2.0'.

    Los códigos no numéricos (p. ej. 'INTERNO') no se modifican.
    """
    code = (code or "").strip()
    if not code:
        return code
    parts = code.split(".")
    if not all(p.strip().isdigit() for p in parts) or len(parts) >= levels:
        return code
    return ".".join(parts + ["0"] * (levels - len(parts)))


def plural(n: int, singular: str, plural_form: str) -> str:
    """'1 tarea', '3 tareas'."""
    return f"{n} {singular if n == 1 else plural_form}"


def tolerance(amount: float) -> float:
    return max(0.5, abs(amount) * 1e-7)


def fmt_money(v: float) -> str:
    return f"{v:,.2f}"


# ----------------------------------------------------------------------------
# Códigos EDT y planificación
# ----------------------------------------------------------------------------
def code_level(code: str) -> int:
    """Nivel de un código EDT: cantidad de segmentos sin contar los ceros finales.

    '2.0.0' -> 1, '2.1.0' -> 2, '2.1.1' -> 3, '2.1.1.4' -> 4, 'INTERNO' -> 1.
    """
    parts = [p.strip() for p in (code or "").strip().split(".")]
    while len(parts) > 1 and not parts[-1].strip("0"):
        parts.pop()
    return len(parts)


@dataclass
class Scope:
    """Tareas a exportar: todo el proyecto o la planificación elegida (p. ej. INTERNO)."""

    roots: list[Task]
    plan: Task | None = None
    prefix: str = ""            # código EDT de la planificación

    def raw_code(self, task: Task) -> str:
        """Código EDT de la tarea, sin el prefijo de la planificación si lo tuviera."""
        code = (task.wbs or "").strip()
        p = self.prefix
        if p and len(code) > len(p) + 1 and code[:len(p)].upper() == p.upper() and code[len(p)] in ".-_ /":
            code = code[len(p) + 1:].strip()
        return code

    def tasks(self):
        for r in self.roots:
            yield from iter_subtree(r)


def select_scope(roots: list[Task], plan_code: str) -> Scope:
    """Elige la planificación por su código EDT (o, si no lo hay, por el nombre de la tarea)."""
    plan_code = (plan_code or "").strip()
    if not plan_code:
        return Scope(list(roots))
    from evm.fields import same_name
    every = [t for r in roots for t in iter_subtree(r)]
    found = sorted((t for t in every if same_name(t.wbs or "", plan_code)), key=lambda t: t.outline_level)
    if not found:
        found = sorted((t for t in every if same_name(t.name or "", plan_code)), key=lambda t: t.outline_level)
    if not found:
        tops = [f"{(t.wbs or '').strip()} ({t.name})" if (t.wbs or "").strip() else t.name for t in roots]
        raise ConfigError(
            f"No se encontró la planificación '{plan_code}': ninguna tarea tiene ese código EDT. Tareas de nivel "
            f"superior del proyecto: {', '.join(tops[:15]) or 'ninguna'}.")
    plan = found[0]
    return Scope(list(plan.children), plan, (plan.wbs or "").strip() or plan_code)


@dataclass
class Visibility:
    """Filas del libro según el código EDT, el nivel máximo y el campo Visible."""

    rows: dict[int, bool]               # UniqueID -> tiene fila propia
    issues: list[Issue]
    shown: int = 0                      # tareas con Visible = Sí y fila propia por debajo del nivel máximo
    # Visible = No: ramas que no forman parte del PV, el EV ni el AC
    excluded: list[Task] = field(default_factory=list)
    # Tareas cuyas subtareas tienen todas Visible = No: la tarea las representa (con toda su rama)
    represented: list[Task] = field(default_factory=list)

    def has_row(self, task: Task) -> bool:
        return self.rows.get(task.uid, False)

    def excluded_uids(self) -> set[int]:
        """Tareas de las ramas excluidas (con sus subtareas)."""
        return {n.uid for t in self.excluded for n in iter_subtree(t)}


def row_visibility(scope: Scope, cfg: Config, code_of=None) -> Visibility:
    """Decide qué tareas tienen fila propia en el libro y cuáles quedan fuera del cálculo.

    - Visible vacío: fila propia si la tarea tiene código EDT de nivel no mayor que el nivel máximo.
    - Visible = Sí: fila propia aunque su código sea más detallado que el nivel máximo (necesita código EDT).
    - Visible = No: la tarea y su rama no forman parte del PV, el EV ni el AC. Si todas las subtareas de una tarea
      tienen Visible = No, esa tarea las representa: no tienen fila propia, pero su rama se suma en la fila de la
      tarea (que, si tiene fila, pasa a ser el paquete de trabajo). Prevalece sobre un Sí en la rama.
    """
    from evm.fields import parse_visible
    code_of = code_of or scope.raw_code
    out = Visibility({}, [])
    values: dict[int, bool | None] = {}
    if cfg.visible_field:               # visible = "" en la configuración: el campo no se usa
        for t in scope.tasks():
            value, valid = parse_visible(t.visible)
            values[t.uid] = value
            if not valid:
                out.issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                        f"Visible = '{t.visible}' no se reconoce (se admite Sí o No): se usa el "
                                        "nivel máximo."))

    def hidden(t: Task) -> bool:
        return values.get(t.uid) is False

    def label(t: Task) -> str:
        return f"{code_of(t) or '(sin código)'} «{t.name}»"

    def drop(t: Task) -> None:
        """Rama excluida: ninguna tarea tiene fila; se avisa de los Sí que quedan fuera."""
        out.excluded.append(t)
        for n in iter_subtree(t):
            out.rows[n.uid] = False
            if n is not t and values.get(n.uid) is True:
                out.issues.append(Issue(SEV_WARN, code_of(n), n.name,
                                        f"Visible = Sí, pero depende de {label(t)}, con Visible = No: queda fuera "
                                        "del cálculo con esa rama."))

    def walk(group: list[Task], parent: Task | None, hidden_by: Task | None) -> None:
        every = bool(group) and all(hidden(t) for t in group)
        if every and parent is not None:
            out.represented.append(parent)
        for t in group:
            raw = scope.raw_code(t)
            value = values.get(t.uid)
            if value is False and not every:
                drop(t)
                continue
            inner = hidden_by or (t if value is False else None)
            if value is True and hidden_by is not None:
                out.issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                        f"Visible = Sí, pero depende de {label(hidden_by)}, con Visible = No: se "
                                        "suma con su rama."))
            elif value is True and not raw:
                out.issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                        "Visible = Sí en una tarea sin código EDT: no puede tener fila propia; se "
                                        "suma en su fila superior."))
            has_row = bool(raw) and inner is None and (value is True or code_level(raw) <= cfg.max_level)
            if has_row and code_level(raw) > cfg.max_level:
                out.shown += 1
            out.rows[t.uid] = has_row
            walk(t.children, t, inner)

    roots = scope.roots
    if roots and all(hidden(t) for t in roots):
        out.issues.append(Issue(SEV_WARN, "", "",
                                "Todos los elementos de nivel superior tienen Visible = No: se ignora en ellos (el "
                                "libro no tendría filas)."))
        for t in roots:
            values[t.uid] = None
    walk(roots, None, None)
    return out


def prune_hidden(scope: Scope, visibility: Visibility) -> dict[int, list[Task]]:
    """Quita del árbol las ramas excluidas por Visible = No (como las tareas inactivas). Devuelve
    {UniqueID de la tarea superior: ramas quitadas}."""
    cut: dict[int, list[Task]] = {}
    out_ids = {t.uid for t in visibility.excluded}
    if not out_ids:
        return cut
    scope.roots = [r for r in scope.roots if r.uid not in out_ids]
    stack = list(scope.roots)
    while stack:
        t = stack.pop()
        removed = [c for c in t.children if c.uid in out_ids]
        if removed:
            cut[t.uid] = removed
            t.children = [c for c in t.children if c.uid not in out_ids]
        stack.extend(t.children)
    return cut


# ----------------------------------------------------------------------------
# Modo porcentaje
# ----------------------------------------------------------------------------
PERCENT_PV_SOURCE = ("% planificado de MS Project (duración de la línea base de las tareas y peso de los hitos) "
                     "× BAC ingresado en el libro")


def _has_weight(t: Task) -> bool:
    """Tarea que puede aportar % planificado: tarea con fechas o hito con peso."""
    if t.milestone:
        return (t.milestone_weight or 0) > 0
    return bool((t.baseline_start or t.start) and (t.baseline_finish or t.finish))


def _percent_curves(project: ProjectData, cfg: Config, rows: list[ReportRow], code_of, share) -> list[Issue]:
    """Modo porcentaje: % planificado de cada fila de ingreso (y de cada empresa) y sus observaciones."""
    from evm import percent as pct
    issues: list[Issue] = []
    calendar = pct.calendar_for(project, cfg)
    cache: dict[int, pct.TaskTime] = {}
    alone: dict[str, list[str]] = {}            # empresa -> filas en las que no participa
    if cfg.multi_enabled:
        roles = [(pct.ROLE_A, cfg.company_a, share), (pct.ROLE_B, cfg.company_b, lambda t: 1.0 - share(t))]
    else:
        roles = [(pct.ROLE_ALL, "", None)]
    for row in rows:
        if not row.is_input:
            continue
        times = [cache.setdefault(t.uid, pct.task_time(t, calendar, cfg.distribution)) for t in row.leaves]
        full = pct.full_weights(times)
        for role, company, fn in roles:
            rw = pct.company_weights(times, fn, company, full)
            if rw.fallback:
                alone.setdefault(company, []).append(row.code)
            row.pct[role] = pct.curve(times, rw)
            issues += [Issue(SEV_WARN, row.code, row.name, note) for note in rw.notes]
        if not any(row.pct.values()):
            issues.append(Issue(SEV_WARN, row.code, row.name,
                                "Ninguna tarea de la fila tiene fechas ni es un hito con peso: no tiene % "
                                "planificado y su BAC no se puede repartir en el tiempo."))

    leaves = list(cache.values())
    undated = [x for x in leaves if not x.dist]
    current = [x for x in leaves if x.dist and not x.from_baseline] if project.baseline_found else []
    if current:
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(current), 'tarea no tiene', 'tareas no tienen')} fechas de línea base "
                            f"(p. ej. se {'agregó' if len(current) == 1 else 'agregaron'} después de guardarla): "
                            "su % planificado se calculó con sus fechas actuales. Si corresponde, se debe "
                            "actualizar la línea base de esas tareas."))
        for x in current[:200]:
            issues.append(Issue(SEV_WARN, code_of(x.task), x.task.name,
                                "Sin fechas de línea base: se usaron sus fechas actuales."))
    for x in undated[:200]:
        issues.append(Issue(SEV_WARN, code_of(x.task), x.task.name,
                            "Sin fechas de línea base ni fechas actuales: no aporta % planificado."))
    with_cost = [x for x in leaves if x.task.baseline_cost or x.task.cost]
    if with_cost:
        issues.append(Issue(SEV_INFO, "", "",
                            f"{plural(len(with_cost), 'tarea tiene', 'tareas tienen')} costo en MS Project: en el "
                            "modo porcentaje los costos no se usan (el importe es el BAC que se ingresa en el libro). "
                            "Para calcular el PV con los costos de MS Project se debe indicar modo = \"costo\" en "
                            "[exportacion]."))
    for company, codes in alone.items():
        shown = ", ".join(codes[:10]) + (f" y {len(codes) - 10} más" if len(codes) > 10 else "")
        issues.append(Issue(SEV_INFO, "", "",
                            f"{company} no participa en las tareas de {plural(len(codes), 'fila', 'filas')} "
                            f"({shown}): si se le asigna BAC, su PV se reparte con el % planificado de la fila "
                            "completa."))
    return issues


# ----------------------------------------------------------------------------
# Informe
# ----------------------------------------------------------------------------
def build_report(project: ProjectData, cfg: Config, source: PVSource) -> Report:
    """Informe de un solo archivo de MS Project."""
    return assemble_report([build_part(project, cfg, source)], cfg)


def build_part(project: ProjectData, cfg: Config, source: PVSource, name: str = "") -> ReportPart:
    """Filas EDT y PV (por mes calendario) de un archivo de MS Project.

    name: nombre del archivo en [[proyectos]] ("" = informe de un solo archivo).
    """
    issues: list[Issue] = [Issue(SEV_WARN, "", "", w) for w in project.warnings]
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)

    # Tareas inactivas (MS Project 2010+): se excluyen con toda su rama
    percent = cfg.percent_mode
    excluded: dict[int, list[Task]] = {}
    if not cfg.include_inactive:
        inactive_roots = [r for r in roots if not r.active]
        roots, excluded, removed = prune_inactive(roots)
        if removed and percent:
            issues.append(Issue(SEV_INFO, "", "",
                                f"Se {'excluyó' if removed == 1 else 'excluyeron'} "
                                f"{plural(removed, 'tarea inactiva', 'tareas inactivas')}: no forman parte del % "
                                "planificado."))
        elif removed:
            excl_cost = (sum(c.baseline_cost for lst in excluded.values() for c in lst)
                         + sum(r.baseline_cost for r in inactive_roots))
            issues.append(Issue(SEV_INFO, "", "",
                                f"Se {'excluyó' if removed == 1 else 'excluyeron'} "
                                f"{plural(removed, 'tarea inactiva', 'tareas inactivas')} (costo de línea base "
                                f"{fmt_money(excl_cost)}, que no forma parte del BAC)."))

    # Planificación (p. ej. INTERNO o CLIENTE): solo las tareas que dependen de ella
    scope = select_scope(roots, cfg.plan)
    if scope.plan is not None:
        issues.append(Issue(SEV_INFO, scope.prefix, scope.plan.name,
                            f"Planificación '{scope.prefix}': se exportan las {sum(1 for _ in scope.tasks())} "
                            f"tareas que dependen de '{scope.plan.name}'."))

    if not project.baseline_found and percent:
        issues.append(Issue(
            SEV_WARN, "", "",
            "El proyecto no tiene guardada la línea base seleccionada: el % planificado se calculó con las "
            "fechas ACTUALES de las tareas. Se debe guardar la línea base en MS Project (Proyecto > Establecer "
            "línea base) y volver a exportar."))
    elif not project.baseline_found:
        issues.append(Issue(
            SEV_WARN, "", "",
            "El proyecto no tiene guardada la línea base seleccionada: el PV y el BAC se calcularon con "
            "el programa ACTUAL (costo y fechas actuales). Se debe guardar la línea base en MS Project "
            "(Proyecto > Establecer línea base) y volver a exportar."))

    def code_of(t: Task) -> str:
        raw = scope.raw_code(t)
        return pad_code(raw, cfg.max_level) if cfg.pad_wbs else raw

    # Fila propia: código EDT de nivel no mayor que el nivel máximo, salvo lo que indique el campo Visible. Las ramas
    # con Visible = No se excluyen del cálculo, como las tareas inactivas
    visibility = row_visibility(scope, cfg, code_of)
    issues += visibility.issues
    included = visibility.has_row
    hidden_cut = prune_hidden(scope, visibility)
    for uid, lst in hidden_cut.items():
        excluded.setdefault(uid, []).extend(lst)

    def cut_cost(t: Task) -> float:
        """Costo de línea base de las ramas excluidas por Visible = No dentro de la rama de t."""
        return sum(c.baseline_cost for n in iter_subtree(t) for c in hidden_cut.get(n.uid, []))

    # Multi-empresa: proporción de la empresa A en cada tarea
    shares: dict[int, float] = {}
    company_notes: dict[str, list[Task]] = {}
    if cfg.multi_enabled:
        from evm.companies import resolve_shares
        shares, company_notes = resolve_shares(scope.roots, cfg)

    def share(t: Task) -> float:
        return shares.get(t.uid, 0.5)

    # --- Costo propio (distribución mensual) de cada tarea -------------------
    # own: costo propio de cada tarea; eff / eff_dist: total efectivo de su rama
    # (para una tarea resumen = max(su línea base, suma efectiva de sus hijos)).
    own: dict[int, dict[MonthKey, float]] = {}
    eff: dict[int, float] = {}
    eff_dist: dict[int, dict[MonthKey, float]] = {}
    no_baseline: list[Task] = []

    def compute_own(t: Task) -> None:
        for c in t.children:
            compute_own(c)
        pruned = excluded.get(t.uid, [])
        if t.is_leaf and not pruned:
            dist = source.leaf(t)
            if t.baseline_cost and not dist:
                issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                    f"Tiene costo de línea base {fmt_money(t.baseline_cost)} pero no tiene "
                                    "fechas: no se pudo distribuir en el tiempo."))
            if (project.baseline_found and not t.has_baseline and not t.baseline_cost
                    and t.cost > tolerance(t.cost)):
                no_baseline.append(t)
            own[t.uid] = dist
            eff[t.uid] = t.baseline_cost
            eff_dist[t.uid] = dict(dist)
            return
        # El costo de línea base de la tarea incluye el de las ramas quitadas (inactivas o con Visible = No),
        # también las que dependen de sus subtareas: no es costo propio
        children_cost = (sum(eff[c.uid] + cut_cost(c) for c in t.children)
                         + sum(c.baseline_cost for c in pruned))
        residual = t.baseline_cost - children_cost
        own[t.uid] = {}
        if residual > tolerance(t.baseline_cost):
            child_dists = [eff_dist[c.uid] for c in t.children] + [source.subtree(c) for c in pruned]
            own[t.uid] = source.summary_residual(t, residual, child_dists)
            issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                f"La tarea resumen tiene costo propio de {fmt_money(residual)} (p. ej. costo fijo "
                                "en la tarea resumen). Se incluye en el PV. Se recomienda asignar costos solo a "
                                "tareas de detalle."))
        elif residual < -tolerance(t.baseline_cost):
            issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                f"El costo de línea base de la tarea resumen ({fmt_money(t.baseline_cost)}) es "
                                f"menor que la suma de sus subtareas ({fmt_money(children_cost)}). Probablemente "
                                "la línea base no se acumuló en las tareas resumen. Se usa la suma de las subtareas."))
        eff[t.uid] = sum(eff[c.uid] for c in t.children) + max(0.0, sum(own[t.uid].values()))
        acc: dict[MonthKey, float] = dict(own[t.uid])
        for c in t.children:
            add_into(acc, eff_dist[c.uid])
        eff_dist[t.uid] = acc

    for r in scope.roots:
        if not percent:                 # modo porcentaje: los costos de MS Project no se usan
            compute_own(r)
    by_wbs = {t.wbs: code_of(t) for t in scope.tasks()}
    for i in getattr(source, "issues", []):
        issues.append(Issue(i.severity, by_wbs.get(i.code, i.code), i.name, i.message))

    if no_baseline:
        total_nb = sum(t.cost for t in no_baseline)
        issues.append(Issue(SEV_WARN, "", "",
                            f"{len(no_baseline)} tareas tienen costo ({fmt_money(total_nb)} en total) pero no "
                            "tienen línea base: no aportan PV. Si corresponde, se debe actualizar la línea base."))
        for t in no_baseline[:200]:
            issues.append(Issue(SEV_WARN, code_of(t), t.name, f"Sin línea base. Costo actual {fmt_money(t.cost)}."))

    # --- Filas del informe -----------------------------------------------------
    rows: list[ReportRow] = []
    report_roots: list[ReportRow] = []

    has_included: dict[int, bool] = {}

    def mark(t: Task) -> bool:
        found = False
        for c in t.children:
            found = mark(c) or included(c) or found
        has_included[t.uid] = found
        return found

    for r in scope.roots:
        mark(r)

    def subtree_pv(t: Task) -> dict[MonthKey, float]:
        acc: dict[MonthKey, float] = {}
        for n in iter_subtree(t):
            add_into(acc, own.get(n.uid, {}))
        return acc

    def subtree_pv_a(t: Task) -> dict[MonthKey, float]:
        acc: dict[MonthKey, float] = {}
        for n in iter_subtree(t):
            add_into(acc, own.get(n.uid, {}), share(n))
        return acc

    def subtree_dates(t: Task) -> tuple[datetime | None, datetime | None]:
        if t.baseline_start or t.baseline_finish:
            return t.baseline_start, t.baseline_finish
        starts = [n.baseline_start for n in iter_subtree(t) if n.baseline_start]
        finishes = [n.baseline_finish for n in iter_subtree(t) if n.baseline_finish]
        return (min(starts) if starts else None, max(finishes) if finishes else None)

    def gather(node: Task) -> tuple[list[Task], list[Task], list[Task]]:
        """Recorre los hijos hasta encontrar tareas con fila propia.

        Devuelve (tareas con fila, ramas completas sin fila, tareas intermedias sin fila).
        Las intermedias aportan solo su costo propio; sus hijos se recorren.
        """
        with_row: list[Task] = []
        loose_full: list[Task] = []
        loose_own: list[Task] = []
        for c in node.children:
            if included(c):
                with_row.append(c)
            elif has_included[c.uid]:
                loose_own.append(c)
                a, b, d = gather(c)
                with_row += a
                loose_full += b
                loose_own += d
            else:
                loose_full.append(c)
        return with_row, loose_full, loose_own

    aggregated = 0      # tareas sin fila propia sumadas en una fila

    def visit(t: Task, parent: ReportRow | None) -> None:
        nonlocal aggregated
        s, f = subtree_dates(t)
        row_tasks, loose_full, loose_own = gather(t)
        # En el nivel máximo la fila suma toda su rama, salvo que una tarea de la rama tenga Visible = Sí
        is_leaf_row = not row_tasks or (code_level(scope.raw_code(t)) >= cfg.max_level
                                        and not any(code_level(scope.raw_code(c)) > cfg.max_level
                                                    for c in row_tasks))
        row = ReportRow(code=code_of(t), name=t.name, level=(parent.level + 1) if parent else 1,
                        kind=KIND_LEAF if is_leaf_row else KIND_SUMMARY, task=t,
                        start=s, finish=f, parent=parent)
        rows.append(row)
        if parent:
            parent.children.append(row)
        else:
            report_roots.append(row)

        if is_leaf_row and percent:
            row.leaves = [n for n in iter_subtree(t) if not n.children]
            row.method = (t.method or "").strip() or cfg.default_method
            row.hidden_tasks = sum(1 for _ in iter_subtree(t)) - 1
            aggregated += row.hidden_tasks
            return
        if is_leaf_row:
            row.pv = subtree_pv(t)
            row.bac = sum(row.pv.values())
            if cfg.multi_enabled:
                row.pv_a = subtree_pv_a(t)
            row.method = (t.method or "").strip() or cfg.default_method
            row.hidden_tasks = sum(1 for _ in iter_subtree(t)) - 1
            aggregated += row.hidden_tasks
            expected = t.baseline_cost - cut_cost(t)          # sin las ramas con Visible = No
            if (t.children and t.baseline_cost
                    and abs(row.bac - expected) > tolerance(t.baseline_cost)):
                issues.append(Issue(SEV_WARN, row.code, t.name,
                                    f"BAC calculado como suma de subtareas ({fmt_money(row.bac)}) difiere del "
                                    f"costo de línea base de la tarea en MS Project ({fmt_money(expected)})."))
            return

        for c in row_tasks:
            visit(c, row)
        if percent:
            # Tareas sin código EDT (o de nivel mayor al máximo) que cuelgan de la tarea resumen: forman la
            # fila -D, con su propio BAC, si tienen duración o son hitos con peso.
            loose = [n for b in loose_full for n in iter_subtree(b) if not n.children]
            aggregated += sum(1 for b in loose_full for _ in iter_subtree(b)) + len(loose_own)
            if any(_has_weight(n) for n in loose):
                res = ReportRow(code=f"{row.code}-D", name=f"Otras tareas de «{t.name}»",
                                level=row.level + 1, kind=KIND_RESIDUAL, task=t,
                                method=cfg.default_method, start=s, finish=f, parent=row)
                res.leaves = loose
                rows.append(res)
                row.children.append(res)
                issues.append(Issue(SEV_INFO, row.code, t.name,
                                    f"Tareas sin código EDT o de nivel mayor a {cfg.max_level}: forman la fila "
                                    f"{row.code}-D, con su propio BAC."))
            return
        # Costos de la tarea resumen que no quedan en ninguna fila hija: su costo propio y el de
        # las tareas sin código EDT (o de nivel mayor al máximo) que cuelgan de ella.
        others = dict(own[t.uid])
        others_a: dict[MonthKey, float] = {}
        add_into(others_a, own[t.uid], share(t))
        for n in loose_full:
            add_into(others, subtree_pv(n))
            add_into(others_a, subtree_pv_a(n))
            aggregated += sum(1 for _ in iter_subtree(n))
        for n in loose_own:
            add_into(others, own.get(n.uid, {}))
            add_into(others_a, own.get(n.uid, {}), share(n))
            aggregated += 1
        uncoded_cost = sum(others.values()) - sum(own[t.uid].values())
        if uncoded_cost > 0.005:
            issues.append(Issue(SEV_INFO, row.code, t.name,
                                f"Tareas sin código EDT o de nivel mayor a {cfg.max_level} con costo de línea base "
                                f"{fmt_money(uncoded_cost)}: se suman en la fila {row.code}-D."))
        own_actual = t.actual_cost - sum(c.actual_cost for c in row_tasks)
        has_own_now = own_actual > tolerance(t.actual_cost)
        if has_own_now and not others:
            issues.append(Issue(SEV_WARN, row.code, t.name,
                                f"La tarea resumen tiene costo real asignado directamente ({fmt_money(own_actual)}) "
                                "que no está en la línea base. Se agrega una fila de otros costos (BAC 0) para "
                                "registrar su costo real."))
        if others or has_own_now:
            res = ReportRow(code=f"{row.code}-D", name=f"Otros costos de «{t.name}»",
                            level=row.level + 1, kind=KIND_RESIDUAL, task=t,
                            method=cfg.default_method, start=s, finish=f, parent=row)
            res.pv = others
            res.bac = sum(res.pv.values())
            res.pv_a = others_a
            rows.append(res)
            row.children.append(res)

    omitted: list[Task] = []

    def top(nodes: list[Task]) -> None:
        for n in nodes:
            if included(n):
                visit(n, None)
            elif has_included[n.uid]:
                omitted.append(n)           # contenedor sin código: se omite, sus hijos se recorren
                top(n.children)
            else:
                omitted.extend(iter_subtree(n))

    top(scope.roots)

    if percent:
        issues += _percent_curves(project, cfg, rows, code_of, share)

    if cfg.multi_enabled:
        # Avisos de empresa, solo de las tareas que forman parte del libro
        omitted_ids = {n.uid for n in omitted}
        notes = {k: [t for t in v if t.uid not in omitted_ids] for k, v in company_notes.items()}
        missing = notes["sin_empresa"]
        missing_ids = {t.uid for t in missing}
        top_missing = [t for t in missing if t.parent is None or t.parent.uid not in missing_ids]
        if top_missing:
            one = len(missing) == 1
            issues.append(Issue(SEV_WARN, "", "",
                                f"{plural(len(missing), 'tarea', 'tareas')} sin empresa en la columna "
                                f"'{cfg.company_field}' (ni en sus tareas superiores): se "
                                f"{'trata' if one else 'tratan'} como '{cfg.company_both}' "
                                f"({cfg.split_default:g} % para {cfg.company_a})."))
            for t in top_missing[:50]:
                issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                    f"Sin empresa: se trata como '{cfg.company_both}'"
                                    + (" (incluye sus subtareas sin empresa)." if t.children else ".")))
        for t in notes["desconocida"][:50]:
            issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                f"Empresa '{t.company}' no reconocida (se esperaba '{cfg.company_a}', "
                                f"'{cfg.company_b}' o '{cfg.company_both}'): se trata como '{cfg.company_both}'."))
        for t in notes["reparto"][:50]:
            issues.append(Issue(SEV_WARN, code_of(t), t.name,
                                f"'{cfg.split_field}' = {t.split_a:g} fuera de rango (debe ser mayor que 0 y como "
                                f"máximo 100): se usa {cfg.split_default:g} %."))

    # --- Controles del archivo ------------------------------------------------
    checks: list[Issue] = []
    if omitted and percent:
        one = len(omitted) == 1
        checks.append(Issue(SEV_WARN if any(_has_weight(n) for n in omitted if not n.children) else SEV_INFO,
                            "", "",
                            f"Se {'omitió' if one else 'omitieron'} {plural(len(omitted), 'tarea', 'tareas')} sin "
                            f"código EDT que no depende{'' if one else 'n'} de ningún elemento EDT: no "
                            f"forma{'' if one else 'n'} parte de ninguna fila del libro."))
    elif omitted:
        omitted_cost = sum(sum(own.get(n.uid, {}).values()) for n in omitted)
        checks.append(Issue(SEV_WARN if omitted_cost > 0.005 else SEV_INFO, "", "",
                            f"Se {'omitió' if len(omitted) == 1 else 'omitieron'} "
                            f"{plural(len(omitted), 'tarea', 'tareas')} sin código EDT que no depende"
                            f"{'' if len(omitted) == 1 else 'n'} de ningún elemento EDT (costo de línea base "
                            f"{fmt_money(omitted_cost)}, que no forma parte del BAC)."))
    if not rows:
        checks.append(Issue(SEV_WARN, "", "",
                            "Ninguna tarea tiene código EDT: el libro no tiene filas. Se debe revisar el campo EDT "
                            "configurado ([campos] edt) y la planificación elegida."))

    part_bac = sum(r.bac for r in rows if r.is_input)
    reference = scope.plan.baseline_cost if scope.plan is not None else project.summary_baseline_cost
    if reference is not None:
        reference -= sum(t.baseline_cost for t in visibility.excluded)     # ramas con Visible = No
    if reference is not None and project.baseline_found and not percent:
        if abs(reference - part_bac) > tolerance(reference):
            whose = f"de la planificación '{scope.prefix}'" if scope.plan is not None else "del proyecto"
            checks.append(Issue(SEV_WARN, "", "",
                                f"El BAC total exportado ({fmt_money(part_bac)}) difiere del costo de línea base "
                                f"{whose} en MS Project ({fmt_money(reference)})."))

    seen: set[str] = set()
    for row in rows:
        if row.code in seen:
            checks.append(Issue(SEV_WARN, row.code, row.name,
                                "Código EDT duplicado." + ("" if scope.plan is not None else
                                                           " Si el proyecto tiene varias planificaciones, se "
                                                           "debe elegir una con 'planificacion'.")))
        else:
            seen.add(row.code)

    if aggregated:
        checks.append(Issue(SEV_INFO, "", "",
                            f"{plural(aggregated, 'tarea', 'tareas')} sin código EDT o de nivel EDT mayor a "
                            f"{cfg.max_level} se {'agregó' if aggregated == 1 else 'agregaron'} en su elemento EDT "
                            "superior."))
    if visibility.excluded:
        n = len(visibility.excluded)
        names = ", ".join(code_of(t) or t.name for t in visibility.excluded[:10])
        more = f" y {n - 10} más" if n > 10 else ""
        if percent:
            detail = ("no forman parte del % planificado ni del % ganado; el BAC y el AC que se ingresan en el libro "
                      "no las deben incluir")
        else:
            cost = sum(t.baseline_cost for t in visibility.excluded)
            detail = f"su costo de línea base ({fmt_money(cost)}) no forma parte del BAC y su costo real no se carga"
        checks.append(Issue(SEV_INFO, "", "",
                            f"{plural(n, 'tarea', 'tareas')} con Visible = No (con sus subtareas) "
                            f"{'queda' if n == 1 else 'quedan'} fuera del PV, el EV y el AC ({names}{more}): "
                            f"{detail}."))
        for t in visibility.excluded[:200]:
            checks.append(Issue(SEV_INFO, code_of(t), t.name, "Visible = No: fuera del PV, el EV y el AC."))
    for t in visibility.represented:
        checks.append(Issue(SEV_INFO, code_of(t), t.name,
                            "Todas sus subtareas tienen Visible = No: "
                            + ("la fila es el paquete de trabajo, con toda su rama." if included(t)
                               else "se suman con la tarea en su fila superior.")))
    if visibility.shown:
        checks.append(Issue(SEV_INFO, "", "",
                            f"{plural(visibility.shown, 'tarea tiene', 'tareas tienen')} Visible = Sí y fila propia "
                            f"por debajo del nivel máximo ({cfg.max_level})."))
    linear = getattr(source, "linear_count", 0)
    if linear and isinstance(source, TimephasedPVSource) and not percent:
        if source.timephased_count:
            checks.append(Issue(SEV_INFO, "", "",
                                f"{plural(linear, 'tarea no tiene', 'tareas no tienen')} distribución mensual de la "
                                "línea base en MS Project (p. ej. costo de línea base escrito a mano): su BAC se "
                                "repartió en forma lineal entre el inicio y el fin de la línea base."))
        elif project.baseline_found:
            checks.append(Issue(SEV_WARN, "", "",
                                "MS Project no entregó la distribución mensual de la línea base: el BAC de cada tarea "
                                "se repartió en forma lineal entre el inicio y el fin de la línea base."))

    for row in rows:
        row.source = name
    dates = [t.baseline_start or t.start for t in scope.tasks() if (t.baseline_start or t.start)]
    ref = min(dates) if dates else project.project_start
    first_date = (ref.date() if isinstance(ref, datetime) else ref) if ref else None
    description = PERCENT_PV_SOURCE if percent else source.description
    return ReportPart(Source(name, project, scope.prefix, description), rows, report_roots, issues, checks,
                      first_date)


def _code_order(row: ReportRow) -> tuple:
    """Orden de los elementos de nivel superior de varios archivos: por código EDT numérico (1.0.0 < 9.0.0 <
    10.0.0); los códigos no numéricos van al final, en su orden."""
    segments = row.code.split(".")
    if all(x.strip().isdigit() for x in segments):
        return (0, tuple(int(x) for x in segments))
    return (1, ())


def _block(row: ReportRow):
    """La fila y sus descendientes, en el orden del libro."""
    yield row
    for c in row.children:
        yield from _block(c)


def _tagged(part: ReportPart, issues: list[Issue]) -> list[Issue]:
    """Con varios archivos, cada observación indica el archivo: '[P2] ...'."""
    name = part.source.name
    if not name:
        return list(issues)
    return [Issue(i.severity, i.code, i.name, f"[{name}] {i.message}") for i in issues]


def assemble_report(parts: list[ReportPart], cfg: Config) -> Report:
    """Combina las filas de uno o varios archivos, fija el horizonte y calcula los vectores mensuales."""
    issues = [i for part in parts for i in _tagged(part, part.issues)]
    if len(parts) > 1:
        # Las hojas identifican cada fila por su código EDT: un código no puede estar en dos archivos
        owner: dict[str, str] = {}
        for part in parts:
            for row in part.rows:
                other = owner.setdefault(row.code, part.source.name)
                if other != part.source.name:
                    raise ConfigError(
                        f"El código EDT '{row.code}' está en los archivos '{other}' y '{part.source.name}'. Con "
                        "varios archivos de MS Project, cada código EDT debe estar en un solo archivo (p. ej. 1.0.0 a "
                        "8.0.0 en uno y 9.0.0 en otro): se deben revisar los códigos EDT o la planificación de cada "
                        "archivo en [[proyectos]].")
        roots = sorted((r for part in parts for r in part.roots), key=_code_order)
        rows = [x for root in roots for x in _block(root)]
    else:
        roots, rows = list(parts[0].roots), list(parts[0].rows)

    # --- Horizonte -------------------------------------------------------------
    # Sin mes de inicio: el primer mes con PV en las filas del libro (las tareas omitidas no cuentan)
    percent = cfg.percent_mode
    if cfg.start_month:
        start = first_of_month(cfg.start_month)
    else:
        if percent:
            keys = [k for row in rows if row.is_input for c in row.pct.values() for k, v in c.items() if v]
        else:
            keys = [k for row in rows if row.is_input for k, v in row.pv.items() if v]
        if keys:
            y, m = min(keys)
            start = date(y, m, 1)
        else:
            firsts = [p.first_date for p in parts if p.first_date]
            start = first_of_month(min(firsts) if firsts else datetime.now().date())
    horizon = Horizon(start, cfg.months)

    # --- Vectores mensuales ---------------------------------------------------
    total_before = total_after = 0.0
    pct_before: set[str] = set()       # modo porcentaje: filas con % planificado fuera del horizonte
    pct_after: set[str] = set()
    for row in rows:
        if row.is_input and percent:
            for role, values in row.pct.items():
                vec, before, after = horizon.to_vector(values)
                vec[0] += before            # se acumula en el mes 1 para que las curvas acumuladas sean correctas
                row.pct_vec[role] = vec
                if before > 1e-9:
                    pct_before.add(row.code)
                if after > 1e-9:
                    pct_after.add(row.code)
            row.vector = [sum(v[i] for v in row.pct_vec.values()) for i in range(horizon.months)]
            row.vector_a = [0.0] * horizon.months
        elif row.is_input:
            row.vector, row.pv_before, row.pv_after = horizon.to_vector(row.pv)
            total_before += row.pv_before
            total_after += row.pv_after
            if row.pv_before:
                # El PV anterior al horizonte se acumula en el mes 1 para que las
                # curvas acumuladas sean correctas.
                row.vector[0] += row.pv_before
            if cfg.multi_enabled:
                row.vector_a, before_a, _after_a = horizon.to_vector(row.pv_a)
                row.vector_a[0] += before_a
                row.bac_a = sum(row.pv_a.values())

    def fill_summary(row: ReportRow) -> None:
        if row.is_input:
            return
        for c in row.children:
            fill_summary(c)
        row.vector = [sum(c.vector[i] for c in row.children) for i in range(horizon.months)]
        row.bac = sum(c.bac for c in row.children)
        if cfg.multi_enabled:
            row.vector_a = [sum(c.vector_a[i] for c in row.children) for i in range(horizon.months)]
            row.bac_a = sum(c.bac_a for c in row.children)
        row.pv_before = sum(c.pv_before for c in row.children)
        row.pv_after = sum(c.pv_after for c in row.children)

    for r in roots:
        fill_summary(r)

    total_vector = [sum(r.vector[i] for r in roots) for i in range(horizon.months)]
    total_bac = sum(r.bac for r in roots)

    if total_before > 0.005:
        issues.append(Issue(SEV_WARN, "", "",
                            f"Hay PV por {fmt_money(total_before)} anterior al mes 1 ({horizon.start:%m/%Y}); "
                            "se acumuló en la columna del mes 1."))
    if total_after > 0.005:
        issues.append(Issue(SEV_WARN, "", "",
                            f"Hay PV por {fmt_money(total_after)} posterior al último mes del horizonte "
                            f"({horizon.end:%m/%Y}); no aparece en las columnas mensuales. Se debe aumentar "
                            "'meses'."))
    if pct_before:
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(pct_before), 'fila tiene', 'filas tienen')} % planificado anterior al mes "
                            f"1 ({horizon.start:%m/%Y}): se acumuló en la columna del mes 1."))
    if pct_after:
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(pct_after), 'fila tiene', 'filas tienen')} % planificado posterior al "
                            f"último mes del horizonte ({horizon.end:%m/%Y}): no aparece en las columnas mensuales "
                            "y su PV no llega al BAC. Se debe aumentar 'meses'."))

    # --- Controles -------------------------------------------------------------
    for part in parts:
        issues += _tagged(part, part.checks)
    if percent:
        sheets = (f"PV-{cfg.company_a} y PV-{cfg.company_b}" if cfg.multi_enabled else "PV")
        issues.insert(0, Issue(SEV_INFO, "", "",
                               "Modo porcentaje: el PV y el EV de cada fila son el BAC × el % planificado y el % "
                               "ganado (hojas % PV y % EV), que se calculan con la duración de la línea base de las "
                               "tareas y el peso de los hitos. Se debe ingresar el BAC de cada fila en las celdas "
                               f"amarillas de la columna D de {'las hojas' if cfg.multi_enabled else 'la hoja'} "
                               f"{sheets}."))
    elif cfg.multi_enabled:
        bac_a = sum(r.bac_a for r in roots)
        issues.append(Issue(SEV_INFO, "", "",
                            f"Multi-empresa: BAC de {cfg.company_a} {fmt_money(bac_a)}; BAC de {cfg.company_b} "
                            f"{fmt_money(total_bac - bac_a)}."))
    if cfg.cost_review and not percent:
        issues.append(Issue(SEV_WARN, "", "", "depuracion_costes solo se aplica en el modo porcentaje: se ignora."))
    undefined = [r for r in rows if r.is_input and r.method.strip().lower() in ("", "por definir", "a definir")]
    if undefined:
        issues.append(Issue(SEV_INFO, "", "",
                            f"{plural(len(undefined), 'fila', 'filas')} sin método de medición definido "
                            "(columna C)."))

    sources = [p.source for p in parts]
    symbols = {s.name: s.project.currency_symbol for s in sources if s.project.currency_symbol}
    currency = next(iter(symbols.values()), "")
    if not cfg.currency and len(set(symbols.values())) > 1:
        issues.append(Issue(SEV_WARN, "", "",
                            "Los archivos tienen monedas distintas en MS Project ("
                            + ", ".join(f"{k}: {v}" for k, v in symbols.items()) + f"): el libro muestra {currency}. "
                            "Se debe indicar la moneda en [exportacion] moneda."))
    if len(sources) > 1 or sources[0].name:
        pv_source = "  |  ".join(f"{s.name}: {s.pv_source}" for s in sources)
        title = cfg.title or " + ".join(dict.fromkeys(s.project.name for s in sources))
        plan = ""
    else:
        pv_source, plan = sources[0].pv_source, sources[0].plan
        title = cfg.title or sources[0].project.name

    report = Report(project=sources[0].project, horizon=horizon, rows=rows, roots=roots, issues=issues,
                    total_vector=total_vector, total_bac=total_bac, pv_before=total_before,
                    pv_after=total_after, pv_source=pv_source, plan=plan,
                    multi=cfg.multi_enabled, company_a=cfg.company_a if cfg.multi_enabled else "",
                    company_b=cfg.company_b if cfg.multi_enabled else "", sources=sources, title=title,
                    currency=currency, mode=cfg.mode, cost_review=cfg.cost_review_role)
    if report.cost_review is not None:
        from evm.cost_review import review_issues
        report.issues[1:1] = review_issues(report)      # después del aviso del modo porcentaje
    return report
