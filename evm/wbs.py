"""Construcción de la estructura EDT del informe (hasta el nivel máximo) y
cálculo del PV mensual de cada fila."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime

from evm.config import Config
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

    @property
    def is_input(self) -> bool:
        return self.kind in (KIND_LEAF, KIND_RESIDUAL)


@dataclass
class Report:
    project: ProjectData
    horizon: Horizon
    rows: list[ReportRow]
    roots: list[ReportRow]
    issues: list[Issue]
    total_vector: list[float]
    total_bac: float
    pv_before: float
    pv_after: float
    pv_source: str = ""


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
            text += f" ({self.linear_count} tareas sin datos mensuales: reparto lineal)"
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
    """Completa un código EDT con ceros: '1' -> '1.0.0', '1.2' -> '1.2.0'."""
    code = (code or "").strip()
    if not code:
        return code
    parts = code.split(".")
    if any(not p for p in parts) or len(parts) >= levels:
        return code
    return ".".join(parts + ["0"] * (levels - len(parts)))


def tolerance(amount: float) -> float:
    return max(0.5, abs(amount) * 1e-7)


def fmt_money(v: float) -> str:
    return f"{v:,.2f}"


# ----------------------------------------------------------------------------
# Informe
# ----------------------------------------------------------------------------
def build_report(project: ProjectData, cfg: Config, source: PVSource) -> Report:
    issues: list[Issue] = [Issue(SEV_WARN, "", "", w) for w in project.warnings]
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)

    # Tareas inactivas (MS Project 2010+): se excluyen con toda su rama
    excluded: dict[int, list[Task]] = {}
    if not cfg.include_inactive:
        roots, excluded, removed = prune_inactive(roots)
        if removed:
            excl_cost = sum(c.baseline_cost for lst in excluded.values() for c in lst)
            issues.append(Issue(SEV_INFO, "", "",
                                f"Se excluyeron {removed} tareas inactivas (costo de línea base "
                                f"{fmt_money(excl_cost)}, que no forma parte del BAC)."))

    if not project.baseline_found:
        issues.append(Issue(
            SEV_WARN, "", "",
            "El proyecto no tiene guardada la línea base seleccionada: el PV y el BAC se calcularon con "
            "el programa ACTUAL (costo y fechas actuales). Se debe guardar la línea base en MS Project "
            "(Proyecto > Establecer línea base) y volver a exportar."))

    # --- Costo propio (distribución mensual) de cada tarea -------------------
    # own: costo propio de cada tarea; eff / eff_dist: total efectivo de su rama
    # (para una tarea resumen = max(su línea base, suma efectiva de sus hijos)).
    own: dict[int, dict[MonthKey, float]] = {}
    eff: dict[int, float] = {}
    eff_dist: dict[int, dict[MonthKey, float]] = {}
    no_baseline: list[Task] = []

    def code_of(t: Task) -> str:
        return pad_code(t.wbs, cfg.max_level) if cfg.pad_wbs else (t.wbs or "").strip()

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
        children_cost = sum(eff[c.uid] for c in t.children) + sum(c.baseline_cost for c in pruned)
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

    for r in roots:
        compute_own(r)
    by_wbs = {t.wbs: code_of(t) for t in tasks}
    for i in getattr(source, "issues", []):
        issues.append(Issue(i.severity, by_wbs.get(i.code, i.code), i.name, i.message))

    if no_baseline:
        total_nb = sum(t.cost for t in no_baseline)
        issues.append(Issue(SEV_WARN, "", "",
                            f"{len(no_baseline)} tareas tienen costo ({fmt_money(total_nb)} en total) pero no "
                            "tienen línea base: no aportan PV. Si corresponde, se debe actualizar la línea base."))
        for t in no_baseline[:200]:
            issues.append(Issue(SEV_WARN, code_of(t), t.name, f"Sin línea base. Costo actual {fmt_money(t.cost)}."))

    # --- Horizonte -------------------------------------------------------------
    if cfg.start_month:
        start = first_of_month(cfg.start_month)
    else:
        keys = [k for d in own.values() for k, v in d.items() if v]
        if keys:
            y, m = min(keys)
            start = date(y, m, 1)
        else:
            dates = [t.baseline_start or t.start for t in tasks if (t.baseline_start or t.start)]
            ref = min(dates) if dates else (project.project_start or datetime.now())
            start = first_of_month(ref.date() if isinstance(ref, datetime) else ref)
    horizon = Horizon(start, cfg.months)

    # --- Filas del informe -----------------------------------------------------
    rows: list[ReportRow] = []
    report_roots: list[ReportRow] = []

    def subtree_pv(t: Task) -> dict[MonthKey, float]:
        acc: dict[MonthKey, float] = {}
        for n in iter_subtree(t):
            add_into(acc, own.get(n.uid, {}))
        return acc

    def subtree_dates(t: Task) -> tuple[datetime | None, datetime | None]:
        if t.baseline_start or t.baseline_finish:
            return t.baseline_start, t.baseline_finish
        starts = [n.baseline_start for n in iter_subtree(t) if n.baseline_start]
        finishes = [n.baseline_finish for n in iter_subtree(t) if n.baseline_finish]
        return (min(starts) if starts else None, max(finishes) if finishes else None)

    def visit(t: Task, parent: ReportRow | None) -> None:
        s, f = subtree_dates(t)
        is_leaf_row = t.is_leaf or t.outline_level >= cfg.max_level
        row = ReportRow(code=code_of(t), name=t.name, level=t.outline_level,
                        kind=KIND_LEAF if is_leaf_row else KIND_SUMMARY, task=t,
                        start=s, finish=f, parent=parent)
        rows.append(row)
        if parent:
            parent.children.append(row)
        else:
            report_roots.append(row)

        if is_leaf_row:
            row.pv = subtree_pv(t)
            row.bac = sum(row.pv.values())
            row.method = (t.method or "").strip() or cfg.default_method
            row.hidden_tasks = sum(1 for _ in iter_subtree(t)) - 1
            if (t.children and t.baseline_cost
                    and abs(row.bac - t.baseline_cost) > tolerance(t.baseline_cost)):
                issues.append(Issue(SEV_WARN, row.code, t.name,
                                    f"BAC calculado como suma de subtareas ({fmt_money(row.bac)}) difiere del "
                                    f"costo de línea base de la tarea en MS Project ({fmt_money(t.baseline_cost)})."))
            return

        for c in t.children:
            visit(c, row)
        # Costo real cargado directamente en la tarea resumen (sin línea base propia):
        # se agrega una fila de costos directos con BAC 0 para poder registrar su AC.
        own_actual = t.actual_cost - sum(c.actual_cost for c in t.children)
        has_own_now = own_actual > tolerance(t.actual_cost)
        if has_own_now and not own.get(t.uid):
            issues.append(Issue(SEV_WARN, row.code, t.name,
                                f"La tarea resumen tiene costo real asignado directamente ({fmt_money(own_actual)}) "
                                "que no está en la línea base. Se agrega una fila de costos directos (BAC 0) para "
                                "registrar su costo real."))
        if own.get(t.uid) or has_own_now:
            res = ReportRow(code=f"{row.code}-D", name=f"Costos directos de «{t.name}»",
                            level=t.outline_level + 1, kind=KIND_RESIDUAL, task=t,
                            method=cfg.default_method, start=s, finish=f, parent=row)
            res.pv = dict(own[t.uid])
            res.bac = sum(res.pv.values())
            rows.append(res)
            row.children.append(res)

    for t in roots:
        visit(t, None)

    # --- Vectores mensuales ---------------------------------------------------
    total_before = total_after = 0.0
    for row in rows:
        if row.is_input:
            row.vector, row.pv_before, row.pv_after = horizon.to_vector(row.pv)
            total_before += row.pv_before
            total_after += row.pv_after
            if row.pv_before:
                # El PV anterior al horizonte se acumula en el mes 1 para que las
                # curvas acumuladas sean correctas.
                row.vector[0] += row.pv_before

    def fill_summary(row: ReportRow) -> None:
        if row.is_input:
            return
        for c in row.children:
            fill_summary(c)
        row.vector = [sum(c.vector[i] for c in row.children) for i in range(horizon.months)]
        row.bac = sum(c.bac for c in row.children)
        row.pv_before = sum(c.pv_before for c in row.children)
        row.pv_after = sum(c.pv_after for c in row.children)

    for r in report_roots:
        fill_summary(r)

    total_vector = [sum(r.vector[i] for r in report_roots) for i in range(horizon.months)]
    total_bac = sum(r.bac for r in report_roots)

    if total_before > 0.005:
        issues.append(Issue(SEV_WARN, "", "",
                            f"Hay PV por {fmt_money(total_before)} anterior al mes 1 ({horizon.start:%m/%Y}); "
                            "se acumuló en la columna del mes 1."))
    if total_after > 0.005:
        issues.append(Issue(SEV_WARN, "", "",
                            f"Hay PV por {fmt_money(total_after)} posterior al último mes del horizonte "
                            f"({horizon.end:%m/%Y}); no aparece en las columnas mensuales. Se debe aumentar "
                            "'meses'."))

    # --- Controles -------------------------------------------------------------
    if project.summary_baseline_cost is not None and project.baseline_found:
        ref = project.summary_baseline_cost
        if abs(ref - total_bac) > tolerance(ref):
            issues.append(Issue(SEV_WARN, "", "",
                                f"El BAC total exportado ({fmt_money(total_bac)}) difiere del costo de línea base "
                                f"del proyecto en MS Project ({fmt_money(ref)})."))

    seen: set[str] = set()
    for row in rows:
        if not row.code:
            issues.append(Issue(SEV_WARN, "", row.name, "Fila sin código EDT."))
        elif row.code in seen:
            issues.append(Issue(SEV_WARN, row.code, row.name, "Código EDT duplicado."))
        else:
            seen.add(row.code)

    hidden = sum(r.hidden_tasks for r in rows if r.kind == KIND_LEAF)
    if hidden:
        issues.append(Issue(SEV_INFO, "", "",
                            f"{hidden} tareas de nivel mayor a {cfg.max_level} se agregaron en su tarea de "
                            f"nivel {cfg.max_level}."))
    linear = getattr(source, "linear_count", 0)
    if linear and isinstance(source, TimephasedPVSource):
        if source.timephased_count:
            issues.append(Issue(SEV_INFO, "", "",
                                f"{linear} tareas no tienen distribución mensual de la línea base en MS Project (p. ej. "
                                "costo de línea base escrito a mano): su BAC se repartió en forma lineal entre el "
                                "inicio y el fin de la línea base."))
        elif project.baseline_found:
            issues.append(Issue(SEV_WARN, "", "",
                                "MS Project no entregó la distribución mensual de la línea base: el BAC de cada tarea "
                                "se repartió en forma lineal entre el inicio y el fin de la línea base."))
    undefined = [r for r in rows if r.is_input and r.method.strip().lower() in ("", "por definir", "a definir")]
    if undefined:
        issues.append(Issue(SEV_INFO, "", "",
                            f"{len(undefined)} filas sin método de medición definido (columna C)."))

    return Report(project=project, horizon=horizon, rows=rows, roots=report_roots, issues=issues,
                  total_vector=total_vector, total_bac=total_bac, pv_before=total_before,
                  pv_after=total_after, pv_source=source.description)
