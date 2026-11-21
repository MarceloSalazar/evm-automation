"""% planificado y % ganado de cada fila del libro (y de cada empresa), sin importes.

MS Project no tiene importes: cada fila del libro recibe un % planificado por mes y un % ganado a la fecha de
corte, y el importe es el BAC que se ingresa en el libro multiplicado por ese %.

Peso de cada tarea de detalle de la fila (fracción del BAC; los pesos de la fila suman 1):
- Tarea con duración: su duración laborable de la línea base, repartida entre los meses que abarcan sus fechas.
- Hito con "Peso hito": ese % del BAC de la empresa en la fila, en el mes de su fecha. El resto
  (100 % - pesos de los hitos) se reparte entre las tareas según su duración. Un hito sin peso no tiene valor.
- Pesos de hitos que suman más de 100: se escalan a 100 % y las tareas no reciben peso.
- Pesos de hitos que suman menos de 100 sin tareas con duración: se escalan a 100 %.
- Sin tareas con duración ni hitos con peso: todas las tareas pesan igual.

Multi-empresa: una tarea pesa para cada empresa su duración multiplicada por la parte de esa empresa (columna
Empresa y % Reparto). Un hito vale su peso para cada empresa que participa en él: el % Reparto reparte la
duración de las tareas, no el peso de los hitos. Si una empresa no participa en ninguna tarea de la fila, se usan
los pesos de la fila completa (por si se le asigna BAC).
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime, time

from evm.fields import fmt_number
from evm.messages import Issue
from evm.methods import (FIXED, LOE, MILESTONES, PCT, PHYSICAL, UNDEFINED, UNITS, WEIGHTED, Method, is_complete,
                         is_started, parse_method)
from evm.model import MonthKey, Task
from evm.timeline import add_into
from evm.workcalendar import WorkCalendar, elapsed_by_month

ROLE_ALL = ""           # libro de una sola empresa
ROLE_A, ROLE_B = "A", "B"

TOL = 1e-9

Share = Callable[[Task], float]


@dataclass
class TaskTime:
    """Ubicación en el tiempo de una tarea de detalle."""

    task: Task
    minutes: float                      # duración laborable (0 en un hito o sin fechas)
    dist: dict[MonthKey, float]         # reparto de la tarea entre meses (suma 1); {} sin fechas
    from_baseline: bool                 # las fechas son las de la línea base


@dataclass
class RowWeights:
    """Pesos de las tareas de una fila para una empresa (fracción del BAC, suman 1)."""

    weights: dict[int, float]                       # UniqueID -> peso
    members: list[TaskTime]                         # tareas en las que participa la empresa
    notes: list[Issue] = field(default_factory=list)    # observaciones para la hoja Validación (sin la fila)
    fallback: bool = False                          # la empresa no participa: pesos de la fila completa


def calendar_for(project, cfg) -> WorkCalendar:
    """Calendario para repartir la duración: el del proyecto, leído del XML o de MS Project (.mpp), con los feriados
    de la configuración; si no se pudo leer, lunes a viernes con los feriados de la configuración."""
    if getattr(project, "calendar", None) is not None:
        return project.calendar
    cal = WorkCalendar()
    cal.add_holidays(cfg.holidays)
    return cal


def _as_datetime(d: datetime | date, end_of_day: bool = False) -> datetime:
    if isinstance(d, datetime):
        return d.replace(tzinfo=None)
    return datetime.combine(d, time(23, 59) if end_of_day else time(0, 0))


def task_dates(t: Task) -> tuple[datetime | None, datetime | None, bool]:
    """Fechas de la línea base de la tarea o, si no las tiene, sus fechas actuales."""
    if t.baseline_start or t.baseline_finish:
        return t.baseline_start, t.baseline_finish, True
    return t.start, t.finish, False


def task_time(t: Task, calendar: WorkCalendar, mode: str = "habiles") -> TaskTime:
    """Duración laborable de la tarea y su reparto entre meses.

    mode = "habiles": según el tiempo laborable del calendario; "calendario": según los días corridos.
    Un hito (o una tarea sin duración) queda entero en el mes de su fecha.
    """
    start, finish, from_bl = task_dates(t)
    if start is None and finish is None:
        return TaskTime(t, 0.0, {}, from_bl)
    if t.milestone:
        when = finish or start
        return TaskTime(t, 0.0, {(when.year, when.month): 1.0}, from_bl)
    s = _as_datetime(start if start is not None else finish)
    f = _as_datetime(finish if finish is not None else start, end_of_day=True)
    if f < s:
        s, f = f, s
    if f == s:
        return TaskTime(t, 0.0, {(s.year, s.month): 1.0}, from_bl)
    weights = calendar.minutes_by_month(s, f) if mode == "habiles" else {}
    if not sum(weights.values()):
        weights = elapsed_by_month(s, f)
    total = sum(weights.values())
    if total <= 0:
        return TaskTime(t, 0.0, {(s.year, s.month): 1.0}, from_bl)
    return TaskTime(t, total, {k: v / total for k, v in weights.items() if v}, from_bl)


def leaf_times(leaves: list[Task], calendar: WorkCalendar, mode: str = "habiles") -> list[TaskTime]:
    return [task_time(t, calendar, mode) for t in leaves]


def row_weights(times: list[TaskTime], share: Share | None = None, company: str = "") -> RowWeights | None:
    """Pesos de las tareas de una fila para una empresa.

    share: parte de la empresa en cada tarea (0 a 1); None = la fila completa (todas las tareas al 100 %).
    company: nombre de la empresa para las observaciones ("" en un libro de una sola empresa).
    Devuelve None si la empresa no participa en ninguna tarea de la fila.
    """
    part = {tt.task.uid: (1.0 if share is None else max(0.0, min(1.0, share(tt.task)))) for tt in times}
    members = [tt for tt in times if part[tt.task.uid] > TOL]
    if not members:
        return None
    who = f" de {company}" if company else ""
    notes: list[Issue] = []
    milestones = [tt for tt in members if tt.task.milestone]
    tasks = [tt for tt in members if not tt.task.milestone]
    negative = [tt for tt in milestones if (tt.task.milestone_weight or 0.0) < 0]
    if negative:
        notes.append(Issue("HIT-006", message=f"{len(negative)} {'hito' if len(negative) == 1 else 'hitos'}{who} "
                                              "con 'Peso hito' negativo: se toma como 0."))
    m_weight = {tt.task.uid: max(0.0, float(tt.task.milestone_weight or 0.0)) for tt in milestones if tt.dist}
    m_total = sum(m_weight.values())
    d_weight = {tt.task.uid: tt.minutes * part[tt.task.uid] for tt in tasks if tt.minutes > 0 and tt.dist}
    d_total = sum(d_weight.values())
    weights: dict[int, float] = {}
    if m_total > TOL and (m_total >= 100 - TOL or d_total <= TOL):
        if m_total > 100 + TOL:
            notes.append(Issue("HIT-007", message=f"Los pesos de los hitos{who} suman {fmt_number(m_total, trim=True)} "
                                                  "%: se escalaron a 100 % y las tareas sin hito no reciben % "
                                                  "planificado."))
        elif m_total < 100 - TOL:
            notes.append(Issue("HIT-008", message=f"Los pesos de los hitos{who} suman {fmt_number(m_total, trim=True)} "
                                                  "% y la fila no tiene tareas con duración: se escalaron a 100 %."))
        weights = {uid: p / m_total for uid, p in m_weight.items() if p > 0}
    elif m_total > TOL or d_total > TOL:
        weights = {uid: p / 100.0 for uid, p in m_weight.items() if p > 0}
        rest = 1.0 - m_total / 100.0
        for uid, d in d_weight.items():
            if d > 0:
                weights[uid] = rest * d / d_total
    else:
        dated = [tt for tt in members if tt.dist]
        if dated:
            notes.append(Issue("HIT-009", message=f"Las tareas{who} no tienen duración ni hitos con peso: todas "
                                                  "pesan igual."))
            weights = {tt.task.uid: 1.0 / len(dated) for tt in dated}
    return RowWeights(weights, members, notes)


def full_weights(times: list[TaskTime]) -> RowWeights:
    """Pesos de la fila completa (todas las tareas al 100 %)."""
    rw = row_weights(times, None)
    return rw if rw is not None else RowWeights({}, [], [])


def company_weights(times: list[TaskTime], share: Share | None, company: str = "",
                    full: RowWeights | None = None) -> RowWeights:
    """Pesos de una empresa; si no participa en la fila, los de la fila completa (fallback=True)."""
    if share is None:
        return full if full is not None else full_weights(times)
    rw = row_weights(times, share, company)
    if rw is not None:
        return rw
    base = full if full is not None else full_weights(times)
    return RowWeights(base.weights, base.members, [], fallback=True)


def curve(times: list[TaskTime], rw: RowWeights) -> dict[MonthKey, float]:
    """% planificado de cada mes (fracciones que suman 1 si la fila tiene pesos)."""
    out: dict[MonthKey, float] = {}
    for tt in times:
        w = rw.weights.get(tt.task.uid)
        if w:
            add_into(out, tt.dist, w)
    return out


def _progress(t: Task, physical: bool) -> float:
    """Avance de una tarea entre 0 y 1. Un hito cuenta entero cuando está cumplido."""
    if t.milestone:
        return 1.0 if is_complete(t) else 0.0
    v = t.physical_percent_complete if physical else t.percent_complete
    return max(0.0, min(1.0, v / 100.0))


def _weighted_progress(rw: RowWeights, physical: bool) -> float:
    if rw.weights:
        by_uid = {tt.task.uid: tt.task for tt in rw.members}
        value = sum(w * _progress(by_uid[uid], physical) for uid, w in rw.weights.items() if uid in by_uid)
        return max(0.0, min(1.0, value))
    if rw.members:                      # sin pesos (tareas sin fechas): promedio simple
        return sum(_progress(tt.task, physical) for tt in rw.members) / len(rw.members)
    return 0.0


def earned_fraction(method: Method, rw: RowWeights, pv_fraction: float) -> tuple[float, Issue | None]:
    """% ganado (0 a 1) de una empresa en una fila a la fecha de corte y una nota explicativa (sin la fila).

    rw: pesos de la empresa (o de la fila completa); pv_fraction: % planificado acumulado de la empresa a la
    fecha de corte (para el avance lineal).
    """
    members = [tt.task for tt in rw.members]
    kind = method.kind
    if kind == FIXED:
        if members and all(is_complete(t) for t in members):
            return 1.0, None
        if any(is_started(t) for t in members):
            return method.start_fraction, None
        return 0.0, None
    if kind == PCT:
        return _weighted_progress(rw, False), None
    if kind in (PHYSICAL, UNITS):
        return _weighted_progress(rw, True), None
    if kind == LOE:
        return max(0.0, min(1.0, pv_fraction)), None
    if kind == MILESTONES:
        milestones = [t for t in members if t.milestone]
        if milestones:
            negative = any((t.milestone_weight or 0.0) < 0 for t in milestones)
            weights = [(t, max(0.0, float(t.milestone_weight or 0.0))) for t in milestones]
            note = None
            if not any(w > 0 for _, w in weights):
                weights = [(t, 1.0) for t in milestones]
                note = (Issue("ACT-029", message="hitos sin peso válido (pesos negativos): se ponderaron por igual")
                        if negative else Issue("ACT-030", message="hitos sin peso definido: se ponderaron por igual"))
            elif negative:
                note = Issue("ACT-031", message="hitos con peso negativo: se tomaron con peso 0")
            total = sum(w for _, w in weights)
            return sum(w for t, w in weights if is_complete(t)) / total, note
        if any(t.physical_percent_complete > 0 for t in members):
            return _weighted_progress(rw, True), Issue("ACT-032", message="sin hitos: se usó % físico completado")
        return _weighted_progress(rw, False), Issue("ACT-033", message="sin hitos ni % físico completado: se usó % "
                                                                       "completado")
    if kind == WEIGHTED:            # «Ponderado» sin pesos (HIT-012): % completado
        return _weighted_progress(rw, False), None
    if kind == UNDEFINED:
        return _weighted_progress(rw, False), Issue("ACT-034", message="método no definido: se usó % completado")
    return _weighted_progress(rw, False), Issue("ACT-035", message=f"método no reconocido '{method.text}': se usó % "
                                                                   "completado")


# ----------------------------------------------------------------------------
# Método «Ponderado» (desde la 2.7.0)
# ----------------------------------------------------------------------------
def task_weight(t: Task) -> float:
    """Peso de una tarea en una fila «Ponderado» (campo de [campos] peso_hito; vacío y negativo = 0)."""
    return max(0.0, float(t.milestone_weight or 0.0))


def is_zero_duration(t: Task) -> bool:
    """Hito: marcado como hito o de duración 0 (sin subtareas)."""
    return not t.children and (t.milestone or (t.duration is not None and t.duration <= 1e-9))


def _has_baseline(t: Task) -> bool:
    return t.baseline_start is not None or t.baseline_finish is not None


@dataclass
class WeightGroup:
    """Tarea resumen de una fila «Ponderado» cuyas subtareas directas tienen pesos que no suman 100.

    Más de 100: error (se escalan a 100 y las subtareas sin peso no cuentan). Menos de 100 y sin subtareas sin peso
    que reciban el resto: se escalan a 100 (aviso)."""

    task: Task
    total: float


@dataclass
class WeightedRow:
    """Reparto de una fila «Ponderado» entre sus tareas de detalle.

    pv: tareas con línea base y su fracción del % planificado (suman 1); ev: todas las tareas y su fracción del %
    ganado (suman 1). Las tareas sin línea base cuentan en el % ganado y no en el planificado. unweighted: tareas (o
    tareas resumen, con sus subtareas) que quedan sin aporte y la causa: NO_SHARE_FULL (sin peso y sus tareas hermanas
    suman 100 o más) o NO_SHARE_DURATION (el resto se reparte por duración y la suya es 0, como un hito). groups:
    tareas resumen con pesos que no suman 100 (WeightGroup).
    """

    pv: list[tuple[TaskTime, float]]
    ev: list[tuple[TaskTime, float]]
    unweighted: list[tuple[Task, str]]
    groups: list[WeightGroup] = field(default_factory=list)


WEIGHT_TOTAL = 100.0
NO_SHARE_FULL = "completo"          # las tareas hermanas con peso suman 100 o más
NO_SHARE_DURATION = "duracion"      # el resto se reparte por duración y la de la tarea es 0
_EPS = 1e-6


def weighted_row(top: Task, time_of: Callable[[Task], TaskTime],
                 allowed: set[int] | None = None) -> WeightedRow | None:
    """Pesos de una fila «Ponderado» (None si ninguna tarea tiene peso).

    En cada tarea resumen (la de la fila y las de nivel inferior que reciben una parte), los pesos de sus subtareas
    directas son % de la parte de la tarea resumen:
    - suman 100: cada subtarea con peso recibe su %; las subtareas sin peso, nada;
    - suman menos de 100: el resto se reparte entre las subtareas sin peso según su duración (un hito sin peso, de
      duración 0, no recibe nada); si no hay subtareas sin peso, los pesos se escalan a 100;
    - suman más de 100: se escalan a 100 (error de los datos: WeightGroup) y las subtareas sin peso no reciben nada;
    - ninguna tiene peso: la parte se reparte según su duración (en partes iguales si todas duran 0).
    Una tarea resumen que recibe una parte la reparte igual entre sus subtareas. allowed: UniqueID de las tareas de
    la fila.
    """
    def kids(t: Task) -> list[Task]:
        return [c for c in t.children if allowed is None or c.uid in allowed]

    def leaves_of(t: Task) -> list[Task]:
        ks = kids(t)
        return [t] if not ks else [x for c in ks for x in leaves_of(c)]

    unweighted: list[tuple[Task, str]] = []
    groups: list[WeightGroup] = []

    def shares(t: Task, ks: list[Task], include, note: bool) -> tuple[dict[int, float], str]:
        """Fracción de la parte de t que recibe cada subtarea (antes de quitar las que no tienen tareas incluidas) y la
        causa por la que las demás no reciben nada."""
        weight = {c.uid: task_weight(c) for c in ks}
        total = sum(weight.values())
        rest = [c for c in ks if weight[c.uid] <= 0]
        over = total > WEIGHT_TOTAL + _EPS
        scaled = over or (total > 0 and not rest)
        if note and scaled and abs(total - WEIGHT_TOTAL) > _EPS:
            groups.append(WeightGroup(t, total))
        base = total if scaled else WEIGHT_TOTAL
        out = {c.uid: weight[c.uid] / base for c in ks if weight[c.uid] > 0}
        remainder = 0.0 if scaled else (WEIGHT_TOTAL - total) / WEIGHT_TOTAL
        if remainder > _EPS:
            minutes = {c.uid: sum(time_of(x).minutes for x in leaves_of(c) if include(x)) for c in rest}
            span = sum(minutes.values())
            if span > 0:
                out.update((c.uid, remainder * minutes[c.uid] / span) for c in rest)
            elif total <= 0:                                    # sin pesos y todas de duración 0: partes iguales
                found = [c for c in rest if any(include(x) for x in leaves_of(c))]
                out.update((c.uid, remainder / len(found)) for c in found)
        return out, (NO_SHARE_DURATION if remainder > _EPS else NO_SHARE_FULL)

    def alloc(t: Task, amount: float, include, note: bool) -> list[tuple[Task, float]]:
        ks = kids(t)
        if not ks:
            return [(t, amount)] if include(t) else []
        frac, why = shares(t, ks, include, note)
        if note:
            unweighted.extend((c, why) for c in ks if frac.get(c.uid, 0.0) <= 0)
        parts = [(c, frac[c.uid]) for c in ks
                 if frac.get(c.uid, 0.0) > 0 and any(include(x) for x in leaves_of(c))]
        total = sum(f for _c, f in parts)
        return [x for c, f in parts for x in alloc(c, amount * f / total, include, note)] if total > 0 else []

    if not kids(top) or not any(task_weight(t) > 0 for c in kids(top) for t in _subtree(c, kids)):
        return None

    def split(include, note: bool) -> list[tuple[TaskTime, float]]:
        pairs = alloc(top, 1.0, include, note)
        s = sum(f for _t, f in pairs)
        return [(time_of(t), f / s) for t, f in pairs if f > 0] if s > 0 else []

    ev = split(lambda t: True, True)
    pv = split(_has_baseline, False)
    return WeightedRow(pv, ev, unweighted, groups)


def _subtree(t: Task, kids: Callable[[Task], list[Task]]) -> list[Task]:
    return [t] + [x for c in kids(t) for x in _subtree(c, kids)]


def company_pairs(pairs: list[tuple[TaskTime, float]], share: Share | None) -> list[tuple[TaskTime, float]]:
    """Fracciones de una empresa (su parte de cada tarea, normalizada); si no participa, las de la fila completa."""
    if share is None:
        return pairs
    own = [(tt, f * max(0.0, min(1.0, share(tt.task)))) for tt, f in pairs]
    total = sum(f for _tt, f in own)
    return [(tt, f / total) for tt, f in own if f > 0] if total > TOL else pairs


def weighted_curve(pairs: list[tuple[TaskTime, float]]) -> dict[MonthKey, float]:
    out: dict[MonthKey, float] = {}
    for tt, f in pairs:
        add_into(out, tt.dist, f)
    return out


def subtask_progress(t: Task, elapsed: float | None = None) -> float:
    """Avance (0 a 1) de una tarea de una fila «Ponderado» según su propio método: un hito, 0/100; una tarea, el de su
    campo de método (fórmula fija, % físico, unidades o avance lineal) o, si no tiene, % completado."""
    if is_zero_duration(t):
        return 1.0 if is_complete(t) else 0.0
    m = parse_method(t.method)
    if m.kind == FIXED:
        return 1.0 if is_complete(t) else (m.start_fraction if is_started(t) else 0.0)
    if m.kind in (PHYSICAL, UNITS):
        return max(0.0, min(1.0, t.physical_percent_complete / 100.0))
    if m.kind == LOE and elapsed is not None:
        return max(0.0, min(1.0, elapsed))
    return max(0.0, min(1.0, t.percent_complete / 100.0))


def weighted_earned(pairs: list[tuple[TaskTime, float]], cut: MonthKey | None,
                    state: Callable[[Task], Task] | None = None) -> float:
    """% ganado (0 a 1) de una fila «Ponderado»: suma de la fracción de cada tarea por su avance. cut: mes de corte
    (para el avance lineal de una tarea: lo planificado hasta ese mes); state: estado de la tarea a esa fecha."""
    total = 0.0
    for tt, f in pairs:
        elapsed = sum(v for k, v in tt.dist.items() if cut is None or k <= cut) if tt.dist else None
        total += f * subtask_progress(state(tt.task) if state else tt.task, elapsed)
    return max(0.0, min(1.0, total))
