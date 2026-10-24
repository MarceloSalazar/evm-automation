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

from evm.methods import FIXED, LOE, MILESTONES, PCT, PHYSICAL, UNDEFINED, UNITS, Method, is_complete, is_started
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
    notes: list[str] = field(default_factory=list)  # observaciones para la hoja Validación
    fallback: bool = False                          # la empresa no participa: pesos de la fila completa


def calendar_for(project, cfg) -> WorkCalendar:
    """Calendario para repartir la duración: el del proyecto (XML, con los feriados de la configuración) o
    lunes a viernes con los feriados de la configuración (.mpp)."""
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


def _fmt(v: float) -> str:
    return f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")


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
    notes: list[str] = []
    milestones = [tt for tt in members if tt.task.milestone]
    tasks = [tt for tt in members if not tt.task.milestone]
    negative = [tt for tt in milestones if (tt.task.milestone_weight or 0.0) < 0]
    if negative:
        notes.append(f"{len(negative)} {'hito' if len(negative) == 1 else 'hitos'}{who} con 'Peso hito' negativo: "
                     "se toma como 0.")
    m_weight = {tt.task.uid: max(0.0, float(tt.task.milestone_weight or 0.0)) for tt in milestones if tt.dist}
    m_total = sum(m_weight.values())
    d_weight = {tt.task.uid: tt.minutes * part[tt.task.uid] for tt in tasks if tt.minutes > 0 and tt.dist}
    d_total = sum(d_weight.values())
    weights: dict[int, float] = {}
    if m_total > TOL and (m_total >= 100 - TOL or d_total <= TOL):
        if m_total > 100 + TOL:
            notes.append(f"Los pesos de los hitos{who} suman {_fmt(m_total)} %: se escalaron a 100 % y las tareas "
                         "sin hito no reciben % planificado.")
        elif m_total < 100 - TOL:
            notes.append(f"Los pesos de los hitos{who} suman {_fmt(m_total)} % y la fila no tiene tareas con "
                         "duración: se escalaron a 100 %.")
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
            notes.append(f"Las tareas{who} no tienen duración ni hitos con peso: todas pesan igual.")
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


def earned_fraction(method: Method, rw: RowWeights, pv_fraction: float) -> tuple[float, str]:
    """% ganado (0 a 1) de una empresa en una fila a la fecha de corte y una nota explicativa.

    rw: pesos de la empresa (o de la fila completa); pv_fraction: % planificado acumulado de la empresa a la
    fecha de corte (para el avance lineal).
    """
    members = [tt.task for tt in rw.members]
    kind = method.kind
    if kind == FIXED:
        if members and all(is_complete(t) for t in members):
            return 1.0, ""
        if any(is_started(t) for t in members):
            return method.start_fraction, ""
        return 0.0, ""
    if kind == PCT:
        return _weighted_progress(rw, False), ""
    if kind in (PHYSICAL, UNITS):
        return _weighted_progress(rw, True), ""
    if kind == LOE:
        return max(0.0, min(1.0, pv_fraction)), ""
    if kind == MILESTONES:
        milestones = [t for t in members if t.milestone]
        if milestones:
            negative = any((t.milestone_weight or 0.0) < 0 for t in milestones)
            weights = [(t, max(0.0, float(t.milestone_weight or 0.0))) for t in milestones]
            if not any(w > 0 for _, w in weights):
                weights = [(t, 1.0) for t in milestones]
                note = ("hitos sin peso válido (pesos negativos): se ponderaron por igual" if negative
                        else "hitos sin peso definido: se ponderaron por igual")
            else:
                note = "hitos con peso negativo: se tomaron con peso 0" if negative else ""
            total = sum(w for _, w in weights)
            return sum(w for t, w in weights if is_complete(t)) / total, note
        if any(t.physical_percent_complete > 0 for t in members):
            return _weighted_progress(rw, True), "sin hitos: se usó % físico completado"
        return _weighted_progress(rw, False), "sin hitos ni % físico completado: se usó % completado"
    note = "método no definido" if kind == UNDEFINED else f"método no reconocido '{method.text}'"
    return _weighted_progress(rw, False), note + ": se usó % completado"
