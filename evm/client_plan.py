"""Planificación para el cliente (comando generar-cliente): qué tareas de la planificación interna recibe el cliente y
cómo queda la lógica del cronograma sin las tareas que no se muestran.

Reglas (docs/generar-cliente.md):

- Se muestra una tarea si el campo de [cliente] campo vale Sí (Si o S, en mayúsculas o minúsculas) en ella y en todas
  sus tareas superiores de la planificación. Con No, vacío u otro valor, la tarea no se muestra, con toda su rama.
- La tarea de la planificación (por ejemplo INTERNO) se conserva y se renombra ([cliente] nombre_planificacion). Las
  tareas de otras planificaciones se quitan, con sus vínculos.
- Un resumen que se muestra sin ninguna subtarea mostrada pasa a ser una tarea (resumen convertido): conserva sus
  fechas, su duración, su avance y sus líneas base, y representa a sus subtareas en la lógica.
- Los vínculos entre tareas mostradas se conservan. Los de las subtareas de un resumen convertido pasan al resumen.
  Si entre dos tareas mostradas hay tareas ocultas, se crea un vínculo puente equivalente: su posición es la suma de
  las posiciones de los vínculos del camino más la duración de las tareas ocultas que el camino recorre de comienzo a
  fin (menos la de las que recorre de fin a comienzo).
- Un vínculo nuevo no debe retrasar a su sucesora respecto de la planificación interna: si lo haría, su posición se
  ajusta a la que dan las fechas actuales. Entre dos tareas solo puede haber un vínculo: de los posibles, se elige el
  que más restringe a la sucesora.
- No se crean vínculos entre tareas que ya están vinculadas, entre un resumen y sus subtareas ni los que cerrarían un
  ciclo.

Esta parte no usa MS Project: calcula los cambios y escribe el informe de control. Los cambios en el .mpp los hace
evm.client_writer con MS Project, en una copia del archivo.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from evm.config import Config, ConfigError
from evm.fields import fmt_number, plain_text
from evm.messages import EvmError, Issue
from evm.model import LINK_FF, LINK_FS, LINK_SF, LINK_SS, Link, ProjectData, Task
from evm.percent import calendar_for
from evm.wbs import Scope, build_tree, plural, select_scope
from evm.workcalendar import WorkCalendar

START, FINISH = "C", "F"
# Puntos que une cada tipo de vínculo: (punto de la predecesora, punto de la sucesora)
POINTS = {LINK_FF: (FINISH, FINISH), LINK_FS: (FINISH, START), LINK_SF: (START, FINISH), LINK_SS: (START, START)}
KIND_OF = {points: kind for kind, points in POINTS.items()}
LINK_NAMES = {LINK_FF: "FF", LINK_FS: "FC", LINK_SF: "CF", LINK_SS: "CC"}

# Formatos de la posición en tiempo laborable (<LagFormat> del XML y LagType de COM; +32 = estimada): minutos, horas,
# días, semanas y meses. Una posición transcurrida o en % no se puede sumar a una duración: el vínculo que la tiene en
# su camino toma la posición de las fechas actuales.
_WORKING_FORMATS = {3, 5, 7, 9, 11}

ORIGIN_SUMMARY = "resumen"      # vínculo de una subtarea que pasa a su resumen convertido
ORIGIN_BRIDGE = "puente"        # vínculo que reemplaza un camino que pasa por tareas ocultas

SKIP_LINKED = "ya están vinculadas"
SKIP_SUMMARY = "resumen y subtarea"
SKIP_CYCLE = "cerraría un ciclo"

MARK_YES, MARK_NO, MARK_EMPTY, MARK_OTHER = "Sí", "No", "vacío", "otro valor"
# Valores del campo del cliente (sin tildes ni mayúsculas): solo Sí, Si y S muestran la tarea. Un valor que no es Sí
# ni No (p. ej. un error de escritura) tampoco la muestra, pero se avisa (CLI-003)
_YES = {"si", "s"}
_NO = {"no", "n"}

# Pasos de búsqueda de los caminos por tareas ocultas desde cada tarea (protección ante cronogramas con ciclos)
_MAX_STEPS = 200_000


@dataclass
class NewLink:
    """Vínculo que se crea en la planificación del cliente."""

    pred: Task
    succ: Task
    kind: int
    lag: float                      # minutos laborables (negativa = adelanto)
    origin: str                     # ORIGIN_SUMMARY o ORIGIN_BRIDGE
    hidden: int = 0                 # puentes: tareas ocultas del camino
    adjusted: bool = False          # la posición se ajustó para no retrasar a la sucesora
    from_dates: bool = False        # el camino tiene una posición transcurrida o en %: posición de las fechas
    slack: float | None = None      # minutos entre la fecha que impone el vínculo y la actual de la sucesora


@dataclass
class SkippedLink:
    pred: Task
    succ: Task
    reason: str                     # SKIP_LINKED, SKIP_SUMMARY o SKIP_CYCLE


@dataclass
class SpanChange:
    """Resumen que se muestra con fechas distintas: las fija una subtarea que no se muestra."""

    task: Task
    start: datetime | None          # fechas de sus subtareas que se muestran
    finish: datetime | None


@dataclass
class DateChange:
    """Tarea del cliente cuyas fechas, después de recalcular en MS Project, no eran las de la planificación interna."""

    task: Task                      # tarea de la planificación interna (fechas originales)
    start: datetime | None          # fechas en el archivo del cliente
    finish: datetime | None
    note: str = ""                  # por qué no se fijó (tarea comenzada, otra delimitación...)


@dataclass
class WriteResult:
    """Resultado de la escritura del archivo del cliente con MS Project (evm.client_writer)."""

    path: str = ""                                              # archivo generado
    source_open: bool = False                                   # el origen estaba abierto en MS Project
    fixed: list[Task] = field(default_factory=list)             # fijadas con «No comenzar antes del»
    differences: list[DateChange] = field(default_factory=list)  # fechas que siguen distintas
    rejected: list[tuple[NewLink, str]] = field(default_factory=list)   # vínculos que MS Project no aceptó
    removed_links: list[tuple[int, int]] = field(default_factory=list)  # creados por MS Project al borrar tareas
    restore_problems: list[tuple[Task, str]] = field(default_factory=list)  # resúmenes convertidos sin restaurar
    field_name_removed: bool = True                             # se quitó el nombre (alias) del campo del cliente
    rounds: int = 0                                             # cálculos de MS Project hasta fijar las fechas

    def issues(self, plan: "ClientPlan") -> list[Issue]:
        """Observaciones de la escritura del archivo del cliente."""
        out: list[Issue] = []
        if self.source_open:
            out.append(Issue("CLI-012", message="El archivo de origen está abierto en MS Project: se usó la versión "
                                                "guardada en el disco (los cambios sin guardar no están en el archivo "
                                                "del cliente)."))
        if self.restore_problems:
            what = plural(len(self.restore_problems), "resumen convertido", "resúmenes convertidos")
            out.append(Issue("CLI-019", message=f"{what} con datos que MS Project no aceptó (avance o línea base): "
                                                "ver el informe de control."))
        if self.removed_links:
            what = plural(len(self.removed_links), "par", "pares")
            out.append(Issue("CLI-016", message=f"MS Project vinculó {what} de tareas al borrar las tareas ocultas: "
                                                "esos vínculos se quitaron."))
        if self.rejected:
            what = plural(len(self.rejected), "vínculo nuevo", "vínculos nuevos")
            out.append(Issue("CLI-015", message=f"MS Project no aceptó {what}: ver el informe de control."))
        if not self.field_name_removed:
            out.append(Issue("CLI-017", message=f"No se pudo quitar el nombre del campo '{plan.field_name}' en el "
                                                "archivo del cliente (sus valores sí se vaciaron)."))
        if self.fixed:
            what = plural(len(self.fixed), "tarea", "tareas")
            out.append(Issue("CLI-013", message=f"{what} con el comienzo fijado con «No comenzar antes del», para que "
                                                "tengan la fecha de la planificación interna."))
        if self.differences:
            what = plural(len(self.differences), "tarea tiene", "tareas tienen")
            out.append(Issue("CLI-014", message=f"{what} fechas distintas de la planificación interna: ver el informe "
                                                "de control."))
        return out


@dataclass
class ClientPlan:
    """Cambios para obtener la planificación del cliente a partir de la planificación interna."""

    project: ProjectData
    scope: Scope
    field_name: str
    new_name: str                               # nombre de la tarea de la planificación en el archivo del cliente
    tasks: list[Task]                           # tareas de la planificación (sin su tarea raíz), en orden
    shown: set[int]                             # UID de las tareas que se muestran
    marks: Counter                              # valor del campo de cada tarea: MARK_YES, MARK_NO...
    unknown: list[Task] = field(default_factory=list)           # valor no reconocido (no se muestran)
    orphans: list[Task] = field(default_factory=list)           # Sí debajo de una tarea sin Sí (las de más arriba)
    orphan_count: int = 0                                       # tareas con Sí que no se muestran
    outside: list[Task] = field(default_factory=list)           # tareas fuera de la planificación (se quitan)
    collapsed: list[Task] = field(default_factory=list)         # resúmenes que pasan a ser tareas
    kept_links: int = 0                         # vínculos entre tareas que se muestran (se conservan)
    kept_pairs: set[tuple[int, int]] = field(default_factory=set)  # (UID predecesora, UID sucesora) de esos vínculos
    removed_links: int = 0                      # vínculos con tareas ocultas o dentro de un resumen convertido
    outside_links: int = 0                      # vínculos con tareas fuera de la planificación
    new_links: list[NewLink] = field(default_factory=list)
    skipped: list[SkippedLink] = field(default_factory=list)
    span_changes: list[SpanChange] = field(default_factory=list)

    @property
    def root(self) -> Task | None:
        return self.scope.plan

    @property
    def hidden(self) -> list[Task]:
        return [t for t in self.tasks if t.uid not in self.shown]

    @property
    def delete_roots(self) -> list[Task]:
        """Tareas que se borran en MS Project (al borrar una tarea se borra su rama): las de otras planificaciones de
        primer nivel y la tarea de más arriba de cada rama oculta."""
        top = [t for t in self.outside if t.parent is None]
        return top + [t for t in self.hidden if t.parent is None or t.parent is self.root or t.parent.uid in self.shown]

    def ref(self, t: Task) -> str:
        """Tarea en el informe, sin su nombre: código EDT y UID (o número de esquema y UID, si no tiene código)."""
        code = self.scope.raw_code(t)
        return f"{code} (UID {t.uid})" if code else f"UID {t.uid} (esquema {t.outline_number or '?'})"

    def issues(self) -> list[Issue]:
        """Observaciones de la consola y del informe de control."""
        out: list[Issue] = []
        if self.unknown:
            values = Counter(t.client_mark.strip() for t in self.unknown)
            shown = ", ".join(f"«{v}» ({n})" for v, n in values.most_common(5))
            out.append(Issue("CLI-003", message=f"{plural(len(self.unknown), 'tarea tiene', 'tareas tienen')} un valor "
                                                f"no reconocido en '{self.field_name}': {shown}. No se muestran."))
        if self.orphans:
            what = plural(self.orphan_count, "tarea con Sí no se muestra", "tareas con Sí no se muestran")
            where = ", ".join(self.ref(t) for t in self.orphans[:10]) + (" y más" if len(self.orphans) > 10 else "")
            out.append(Issue("CLI-004", message=f"{what} porque una tarea superior no tiene Sí: {where}."))
        if self.collapsed:
            out.append(Issue("CLI-005", message=f"{plural(len(self.collapsed), 'resumen pasa', 'resúmenes pasan')} a "
                                                "ser tareas: no se muestra ninguna de sus subtareas."))
        if self.new_links:
            bridges = sum(1 for x in self.new_links if x.origin == ORIGIN_BRIDGE)
            adjusted = sum(1 for x in self.new_links if x.adjusted or x.from_dates)
            out.append(Issue("CLI-006", message=f"Vínculos nuevos: {len(self.new_links)} ({bridges} puentes por tareas "
                                                f"ocultas y {len(self.new_links) - bridges} pasados a un resumen "
                                                f"convertido); {adjusted} con la posición de las fechas actuales."))
        if any(s.reason != SKIP_LINKED for s in self.skipped):
            reasons = Counter(s.reason for s in self.skipped if s.reason != SKIP_LINKED)
            out.append(Issue("CLI-007", message="Vínculos que no se crean: " + ", ".join(
                f"{n} porque {r}" if r == SKIP_CYCLE else f"{n} entre {r}" for r, n in reasons.items()) + "."))
        if self.span_changes:
            out.append(Issue("CLI-008", message=f"{plural(len(self.span_changes), 'resumen cambia', 'resúmenes cambian')} "
                                                "de fechas: las fija una subtarea que no se muestra."))
        return out


# ----------------------------------------------------------------------------
# Cálculo
# ----------------------------------------------------------------------------
def _mark(value: str) -> tuple[bool, str]:
    """(se muestra por su propio valor, categoría del valor)."""
    text = plain_text(value or "")
    if not text:
        return False, MARK_EMPTY
    if text in _YES:
        return True, MARK_YES
    return False, MARK_NO if text in _NO else MARK_OTHER


def _active(t: Task) -> bool:
    """La tarea y sus tareas superiores están activas (una tarea inactiva no participa en el cálculo de fechas)."""
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


def _is_ancestor(a: Task, b: Task) -> bool:
    n = b.parent
    while n is not None:
        if n is a:
            return True
        n = n.parent
    return False


def _working(link: Link) -> bool:
    return link.lag == 0 or link.lag_format % 32 in _WORKING_FORMATS


def working_minutes(calendar: WorkCalendar, a: datetime, b: datetime) -> float:
    """Minutos laborables de a a b (negativos si b es anterior)."""
    if b >= a:
        return sum(calendar.minutes_by_month(a, b).values())
    return -sum(calendar.minutes_by_month(b, a).values())


def _point(t: Task, point: str) -> datetime | None:
    return t.start if point == START else t.finish


def _from_start(link: NewLink, calendar: WorkCalendar) -> None:
    """Un vínculo con la posición de las fechas actuales que queda con adelanto (fin a comienzo con -40 días, por
    ejemplo, desde un resumen convertido que termina después de la subtarea que fijaba la sucesora) pasa a partir del
    comienzo de la predecesora, con la posición positiva que da la misma fecha. Los adelantos se consideran un defecto
    de la lógica del cronograma (evaluación DCMA)."""
    pred_point, succ_point = POINTS[link.kind]
    target = _point(link.succ, succ_point)
    if pred_point != FINISH or link.pred.start is None or target is None:
        return
    lag = working_minutes(calendar, link.pred.start, target)
    if lag >= 0:
        link.kind, link.lag = KIND_OF[(START, succ_point)], lag


@dataclass
class _Candidate:
    pred: Task
    succ: Task
    kind: int
    lag: float
    origin: str
    hidden: int
    working: bool


def build_client_plan(project: ProjectData, cfg: Config) -> ClientPlan:
    """Tareas que se muestran al cliente, resúmenes convertidos y vínculos nuevos de la planificación del cliente."""
    roots = build_tree(project.tasks)
    scope = select_scope(roots, cfg.plan)
    if scope.plan is not None and scope.plan.parent is not None:
        raise EvmError("CLI-001", f"La planificación '{cfg.plan}' no es una tarea de primer nivel del esquema: la "
                                  "planificación del cliente se genera a partir de una tarea de primer nivel (por "
                                  "ejemplo INTERNO) y sus subtareas.")
    tasks = list(scope.tasks())
    in_scope = {t.uid for t in tasks}
    root = scope.plan
    plan = ClientPlan(project=project, scope=scope, field_name=cfg.client_field,
                      new_name=cfg.client_plan_name if root is not None else "", tasks=tasks, shown=set(),
                      marks=Counter())
    plan.outside = [t for t in project.tasks if t.uid not in in_scope and t is not root]

    # Tareas que se muestran: Sí en la tarea y en todas sus tareas superiores
    own: dict[int, bool] = {}
    for t in tasks:
        own[t.uid], category = _mark(t.client_mark)
        plan.marks[category] += 1
        if category == MARK_OTHER:
            plan.unknown.append(t)

    def walk(nodes: list[Task], parent_shown: bool) -> None:
        for t in nodes:
            shown = parent_shown and own[t.uid]
            if shown:
                plan.shown.add(t.uid)
            elif own[t.uid]:
                plan.orphan_count += 1
                if t.parent is not None and t.parent is not root and not own.get(t.parent.uid, False):
                    plan.orphans.append(t)
            walk(t.children, shown)

    walk(scope.roots, True)
    if not plan.shown:
        raise EvmError("CLI-002", f"Ninguna tarea de la planificación tiene Sí en '{cfg.client_field}' (también en sus "
                                  "tareas superiores): no hay nada que mostrar al cliente.")
    plan.collapsed = [t for t in tasks if t.uid in plan.shown and t.children
                      and not any(c.uid in plan.shown for c in t.children)]
    collapsed = {t.uid for t in plan.collapsed}

    def rep(t: Task) -> Task | None:
        """Tarea del cliente que representa a t en la lógica: ella misma, el resumen convertido que la contiene o
        None (tarea oculta entre tareas que se muestran, o fuera de la planificación)."""
        if t.uid not in in_scope:
            return None
        if t.uid in plan.shown:
            return t
        a = t.parent
        while a is not None and a.uid in in_scope:
            if a.uid in plan.shown:
                return a if a.uid in collapsed else None
            a = a.parent
        return None

    by_uid = {t.uid: t for t in project.tasks}
    succs: dict[int, list[tuple[Task, Link]]] = defaultdict(list)
    links: list[tuple[Task, Task, Link]] = []
    for t in project.tasks:
        for lk in t.predecessors:
            p = by_uid.get(lk.uid)
            if p is None:
                continue
            if t.uid in in_scope and p.uid in in_scope:
                links.append((p, t, lk))
                succs[p.uid].append((t, lk))
            elif t.uid in in_scope or p.uid in in_scope:
                plan.outside_links += 1

    reps = {t.uid: rep(t) for t in tasks}
    candidates: dict[tuple[int, int], list[_Candidate]] = defaultdict(list)
    existing: set[tuple[int, int]] = set()
    starts: list[tuple[Task, str, Task, str, float, bool]] = []
    for p, s, lk in links:
        rp, rs = reps[p.uid], reps[s.uid]
        if rp is p and rs is s:
            plan.kept_links += 1
            plan.kept_pairs.add((p.uid, s.uid))
            existing.add((p.uid, s.uid))
            continue
        plan.removed_links += 1
        if not (_active(p) and _active(s)) or rp is None:
            continue
        pp, ps = POINTS.get(lk.kind, POINTS[LINK_FS])
        if rs is not None:
            if rs is not rp:
                candidates[(rp.uid, rs.uid)].append(_Candidate(rp, rs, KIND_OF[(pp, ps)], lk.lag, ORIGIN_SUMMARY, 0,
                                                               _working(lk)))
        else:
            starts.append((rp, pp, s, ps, lk.lag, _working(lk)))

    # Puentes: caminos por tareas ocultas (sin representante) desde una tarea que se muestra hasta otra
    by_start: dict[tuple[int, str], list[tuple[Task, str, float, bool]]] = defaultdict(list)
    for a, pa, s, ps, lag, ok in starts:
        by_start[(a.uid, pa)].append((s, ps, lag, ok))
    for (a_uid, pa), firsts in by_start.items():
        a = by_uid[a_uid]
        best: dict[tuple[int, str], float] = {}
        stack = [(s, ps, lag, 1, ok) for s, ps, lag, ok in firsts]
        steps = 0
        while stack and steps < _MAX_STEPS:
            steps += 1
            h, entry, acc, n, ok = stack.pop()
            if best.get((h.uid, entry), float("-inf")) >= acc:
                continue
            best[(h.uid, entry)] = acc
            duration = 0.0 if h.milestone else (h.duration or 0.0)
            for y, lk in succs[h.uid]:
                if not _active(y):
                    continue
                exit_point, entry2 = POINTS.get(lk.kind, POINTS[LINK_FS])
                step = duration if (entry, exit_point) == (START, FINISH) else \
                    -duration if (entry, exit_point) == (FINISH, START) else 0.0
                acc2 = acc + step + lk.lag
                ok2 = ok and _working(lk)
                ry = reps.get(y.uid)
                if ry is not None:
                    if ry is not a:
                        candidates[(a.uid, ry.uid)].append(_Candidate(a, ry, KIND_OF[(pa, entry2)], acc2,
                                                                      ORIGIN_BRIDGE, n, ok2))
                elif y.uid in in_scope:
                    stack.append((y, entry2, acc2, n + 1, ok2))

    # Un vínculo por par de tareas: el que más restringe a la sucesora, sin retrasarla respecto de las fechas actuales
    calendar = calendar_for(project, cfg)
    graph: dict[int, set[int]] = defaultdict(set)
    for p_uid, s_uid in existing:
        graph[p_uid].add(s_uid)

    def reaches(src: int, dst: int) -> bool:
        seen, stack = set(), [src]
        while stack:
            x = stack.pop()
            if x == dst:
                return True
            if x in seen:
                continue
            seen.add(x)
            stack.extend(graph[x])
        return False

    order = {t.uid: k for k, t in enumerate(project.tasks)}
    for (p_uid, s_uid) in sorted(candidates, key=lambda k: (order[k[0]], order[k[1]])):
        cands = candidates[(p_uid, s_uid)]
        p, s = by_uid[p_uid], by_uid[s_uid]
        if (p_uid, s_uid) in existing or (s_uid, p_uid) in existing:
            plan.skipped.append(SkippedLink(p, s, SKIP_LINKED))
            continue
        if _is_ancestor(p, s) or _is_ancestor(s, p):
            plan.skipped.append(SkippedLink(p, s, SKIP_SUMMARY))
            continue
        chosen: NewLink | None = None
        for c in cands:
            pp, ps = POINTS[c.kind]
            a, b = _point(p, pp), _point(s, ps)
            gap = working_minutes(calendar, a, b) if a is not None and b is not None else None
            link = NewLink(p, s, c.kind, c.lag, c.origin, c.hidden)
            if not c.working:
                if gap is None:
                    continue
                link.lag, link.from_dates = gap, True
            if gap is not None:
                link.slack = gap - link.lag
            if chosen is None or (link.slack is not None and (chosen.slack is None or link.slack < chosen.slack)):
                chosen = link
        if chosen is None:
            continue
        if chosen.slack is not None and chosen.slack < 0:
            chosen.lag += chosen.slack          # = minutos laborables entre las dos fechas actuales
            chosen.slack, chosen.adjusted = 0.0, True
        if (chosen.adjusted or chosen.from_dates) and chosen.lag < 0:
            _from_start(chosen, calendar)
        if reaches(s_uid, p_uid):
            plan.skipped.append(SkippedLink(p, s, SKIP_CYCLE))
            continue
        graph[p_uid].add(s_uid)
        existing.add((p_uid, s_uid))
        plan.new_links.append(chosen)

    # Resúmenes que se muestran con fechas distintas: las fija una subtarea oculta
    def leaves(t: Task) -> list[Task]:
        out = []
        for c in t.children:
            if c.uid not in plan.shown or not _active(c):
                continue
            out.extend(leaves(c) if c.uid not in collapsed and c.children else [c])
        return out

    for t in tasks:
        if t.uid not in plan.shown or t.uid in collapsed or not t.children or not _active(t):
            continue
        found = leaves(t)
        first = [x.start for x in found if x.start is not None]
        last = [x.finish for x in found if x.finish is not None]
        if not first or not last:
            continue
        start, finish = min(first), max(last)
        if (start, finish) != (t.start, t.finish):
            plan.span_changes.append(SpanChange(t, start, finish))
    return plan


# ----------------------------------------------------------------------------
# Informe de control (Markdown)
# ----------------------------------------------------------------------------
def _cell(value) -> str:
    return " ".join(str(value or "").split()).replace("|", "\\|")


def _d(value: datetime | None) -> str:
    return f"{value:%d/%m/%Y}" if value else "—"


def _days(minutes: float, per_day: float) -> str:
    return fmt_number(round(minutes / per_day, 2), trim=True)


def render_control(plan: ClientPlan, source: str, today: date, result: str,
                   issues: list[Issue] | None = None, outcome: WriteResult | None = None) -> str:
    """Informe de control interno de generar-cliente: lo que se quita, se convierte y se vincula, y (outcome) lo que
    hizo MS Project al escribir el archivo. issues: las observaciones que se muestran (sin las silenciadas); por
    defecto, las del cálculo y las de la escritura."""
    p = plan.project
    per_day = p.minutes_per_day or 480.0
    hidden = len(plan.hidden)
    marks = plan.marks
    hidden_detail = (f"No: {marks[MARK_NO]}, sin valor: {marks[MARK_EMPTY]}, otro valor: {marks[MARK_OTHER]}, con Sí "
                     f"debajo de una tarea que no se muestra: {plan.orphan_count}")
    bridges = [x for x in plan.new_links if x.origin == ORIGIN_BRIDGE]
    by_kind = Counter(LINK_NAMES[x.kind] for x in plan.new_links)
    plan_text = "todo el proyecto"
    if plan.root is not None:
        plan_text = plan.scope.prefix + (f" (en el archivo del cliente: {plan.new_name})"
                                         if plan.new_name and plan.new_name != plan.root.name else "")
    lines = [
        "# Planificación para el cliente: informe de control",
        "",
        "| | |",
        "|---|---|",
        f"| **Origen** | {_cell(Path(source).name)} |",
        f"| **Planificación** | {_cell(plan_text)} |",
        f"| **Campo** | {_cell(plan.field_name)}: solo se muestran las tareas con Sí (y con Sí en sus tareas superiores) |",
        f"| **Fecha de estado** | {_d(p.status_date)} |",
        f"| **Resultado** | {_cell(result)} |",
        f"| **Fecha de emisión** | {today:%d/%m/%Y} |",
        "",
        "> Documento interno: no se entrega al cliente. Las tareas se identifican por su código EDT y su UID (Id. único "
        "de MS Project).",
        "",
        "## Resumen",
        "",
        "| Dato | Cantidad |",
        "|---|---:|",
        f"| Tareas de la planificación | {len(plan.tasks)} |",
        f"| Se muestran | {len(plan.shown)} |",
        f"| No se muestran | {hidden} ({hidden_detail}; incluye sus subtareas) |" if hidden else "| No se muestran | 0 |",
        f"| Tareas de otras planificaciones (se quitan) | {len(plan.outside)} |",
        f"| Resúmenes que pasan a ser tareas | {len(plan.collapsed)} |",
        f"| Vínculos que se conservan | {plan.kept_links} |",
        f"| Vínculos que desaparecen con las tareas ocultas o dentro de un resumen convertido | {plan.removed_links} |",
        f"| Vínculos con otras planificaciones (desaparecen) | {plan.outside_links} |",
        f"| Vínculos nuevos | {len(plan.new_links)} ({len(bridges)} puentes y {len(plan.new_links) - len(bridges)} "
        f"pasados a un resumen convertido; {', '.join(f'{k} {n}' for k, n in sorted(by_kind.items())) or 'ninguno'}) |",
        f"| Vínculos nuevos con la posición ajustada a las fechas actuales | "
        f"{sum(1 for x in plan.new_links if x.adjusted or x.from_dates)} |",
        f"| Vínculos nuevos que no se crean porque las dos tareas ya están vinculadas | "
        f"{sum(1 for x in plan.skipped if x.reason == SKIP_LINKED)} |",
        f"| Vínculos nuevos que no se crean (entre un resumen y su subtarea, o cerrarían un ciclo) | "
        f"{sum(1 for x in plan.skipped if x.reason != SKIP_LINKED)} |",
        f"| Resúmenes que cambian de fechas | {len(plan.span_changes)} |",
        "",
    ]
    if issues is None:
        issues = plan.issues() + (outcome.issues(plan) if outcome is not None else [])
    if issues:
        lines += ["## Observaciones", ""]
        lines += [f"- **{i.id}** ({i.severity}): {_cell(i.message)}" for i in issues]
        lines.append("")
    if plan.orphans:
        lines += ["## Tareas con Sí que no se muestran", "",
                  "Tienen Sí, pero su tarea superior no: no se muestran, con sus subtareas. Si se deben mostrar, hay "
                  "que poner Sí también en la tarea superior.", "",
                  "| Tarea | Tarea superior sin Sí |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {plan.ref(t.parent)} |" for t in plan.orphans]
        lines.append("")
    if plan.unknown:
        lines += ["## Valores no reconocidos", "", "Se tratan como No.", "", "| Tarea | Valor |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {_cell(t.client_mark)} |" for t in plan.unknown]
        lines.append("")
    if plan.collapsed:
        lines += ["## Resúmenes que pasan a ser tareas", "",
                  "No se muestra ninguna de sus subtareas: el resumen pasa a ser una tarea con sus fechas, su duración, "
                  "su avance y sus líneas base, y recibe los vínculos de sus subtareas.", "",
                  "| Tarea | Comienzo | Fin | % completado |", "|---|---|---|---:|"]
        lines += [f"| {plan.ref(t)} | {_d(t.start)} | {_d(t.finish)} | {fmt_number(t.percent_complete, trim=True)} |"
                  for t in plan.collapsed]
        lines.append("")
    if plan.new_links:
        lines += ["## Vínculos nuevos", "",
                  "Posición en días laborables (negativa = adelanto). *Puente*: reemplaza un camino que pasa por "
                  "tareas ocultas. *Resumen*: vínculo de una subtarea que pasa a su resumen convertido. *Ajustada*: "
                  "con la posición calculada, el vínculo retrasaría a la sucesora; se usa la de las fechas actuales.",
                  "", "| Predecesora | Sucesora | Tipo | Posición | Origen | Observación |",
                  "|---|---|---|---:|---|---|"]
        for x in plan.new_links:
            remarks = []
            if x.origin == ORIGIN_BRIDGE:
                remarks.append(plural(x.hidden, "tarea oculta", "tareas ocultas"))
            if x.adjusted:
                remarks.append("ajustada")
            if x.from_dates:
                remarks.append("posición de las fechas (camino con posición transcurrida o en %)")
            lines.append(f"| {plan.ref(x.pred)} | {plan.ref(x.succ)} | {LINK_NAMES[x.kind]} | "
                         f"{_days(x.lag, per_day)} | {x.origin.capitalize()} | {', '.join(remarks)} |")
        lines.append("")
    if plan.skipped:
        lines += ["## Vínculos que no se crean", "", "| Predecesora | Sucesora | Motivo |", "|---|---|---|"]
        lines += [f"| {plan.ref(s.pred)} | {plan.ref(s.succ)} | {s.reason.capitalize()} |" for s in plan.skipped]
        lines.append("")
    if plan.span_changes:
        lines += ["## Resúmenes que cambian de fechas", "",
                  "Sus fechas las fija una subtarea que no se muestra: en el archivo del cliente, el resumen va de la "
                  "primera a la última de sus subtareas que se muestran.", "",
                  "| Resumen | Comienzo | Fin | Comienzo para el cliente | Fin para el cliente |",
                  "|---|---|---|---|---|"]
        lines += [f"| {plan.ref(c.task)} | {_d(c.task.start)} | {_d(c.task.finish)} | {_d(c.start)} | {_d(c.finish)} |"
                  for c in plan.span_changes]
        lines.append("")
    if outcome is not None:
        lines += _outcome_lines(plan, outcome, per_day)
    return "\n".join(lines)


def _dt_min(value: datetime | None) -> str:
    return f"{value:%d/%m/%Y %H:%M}" if value else "—"


def _outcome_lines(plan: ClientPlan, out: WriteResult, per_day: float) -> list[str]:
    """Secciones de la escritura con MS Project."""
    source_open = "Sí: se usó la versión guardada" if out.source_open else "No"
    field_name = "Quitado" if out.field_name_removed else "No se pudo quitar"
    lines = ["## Escritura con MS Project", "",
             "| Dato | Resultado |", "|---|---|",
             f"| Archivo del cliente | {_cell(Path(out.path).name)} (de solo lectura) |",
             f"| Origen abierto en MS Project | {source_open} |",
             f"| Cálculos de MS Project | {out.rounds} |",
             f"| Tareas con el comienzo fijado (No comenzar antes del) | {len(out.fixed)} |",
             f"| Tareas con fechas distintas de la planificación interna | {len(out.differences)} |",
             f"| Vínculos nuevos que MS Project no aceptó | {len(out.rejected)} |",
             f"| Vínculos que MS Project creó al borrar tareas (se quitaron) | {len(out.removed_links)} |",
             f"| Nombre del campo del cliente | {field_name} (valores vaciados) |", ""]
    if out.fixed:
        lines += ["### Fechas fijadas", "",
                  "Sin las tareas ocultas, MS Project las adelantaba: se fijó su comienzo con la delimitación «No "
                  "comenzar antes del» en la fecha de la planificación interna.", "",
                  "| Tarea | Comienzo | Fin |", "|---|---|---|"]
        lines += [f"| {plan.ref(t)} | {_dt_min(t.start)} | {_dt_min(t.finish)} |" for t in out.fixed]
        lines.append("")
    if out.differences:
        lines += ["### Fechas distintas de la planificación interna", "",
                  "| Tarea | Comienzo | Fin | Comienzo para el cliente | Fin para el cliente | Motivo |",
                  "|---|---|---|---|---|---|"]
        lines += [f"| {plan.ref(c.task)} | {_dt_min(c.task.start)} | {_dt_min(c.task.finish)} | {_dt_min(c.start)} | "
                  f"{_dt_min(c.finish)} | {_cell(c.note)} |" for c in out.differences]
        lines.append("")
    if out.rejected:
        lines += ["### Vínculos que MS Project no aceptó", "",
                  "| Predecesora | Sucesora | Tipo | Posición | Error de MS Project |", "|---|---|---|---:|---|"]
        lines += [f"| {plan.ref(x.pred)} | {plan.ref(x.succ)} | {LINK_NAMES[x.kind]} | {_days(x.lag, per_day)} | "
                  f"{_cell(error)} |" for x, error in out.rejected]
        lines.append("")
    if out.removed_links:
        by_uid = {t.uid: t for t in plan.project.tasks}

        def ref(uid: int) -> str:
            return plan.ref(by_uid[uid]) if uid in by_uid else f"UID {uid}"

        lines += ["### Vínculos que MS Project creó al borrar tareas", "",
                  "Se quitaron: no son vínculos de la planificación interna.", "",
                  "| Predecesora | Sucesora |", "|---|---|"]
        lines += [f"| {ref(a)} | {ref(b)} |" for a, b in out.removed_links]
        lines.append("")
    if out.restore_problems:
        lines += ["### Resúmenes convertidos con datos sin restaurar", "", "| Tarea | Detalle |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {_cell(detail)} |" for t, detail in out.restore_problems]
        lines.append("")
    return lines


def write_control(plan: ClientPlan, source: str, path: str | Path, result: str, issues: list[Issue] | None = None,
                  today: date | None = None, outcome: WriteResult | None = None) -> None:
    Path(path).write_text(render_control(plan, source, today or date.today(), result, issues, outcome),
                          encoding="utf-8")


def require_field(cfg: Config) -> None:
    """generar-cliente necesita el campo que indica qué tareas se muestran ([cliente] campo)."""
    if not cfg.client_field:
        raise ConfigError("CFG-054", "Falta el campo que indica qué tareas se muestran al cliente: [cliente] campo "
                                     "(por ejemplo campo = \"Etiqueta Cliente\").")


def require_client_field(project: ProjectData, cfg: Config) -> None:
    """El campo de [cliente] campo debe existir en el archivo (lo indica el lector en fields_found)."""
    if "cliente" not in project.fields_found:
        raise ConfigError("LEC-026", f"No se encontró el campo '{cfg.client_field}' ([cliente] campo) en el archivo de "
                                     "MS Project: no se puede saber qué tareas se muestran al cliente.")
