"""% planificado y % ganado de cada fila del libro (y de cada empresa), sin importes.

MS Project no tiene importes: cada fila del libro recibe un % planificado por mes y un % ganado a la fecha de
corte, y el importe es el BAC que se ingresa en el libro multiplicado por ese %.

Reparto de la fila entre sus tareas (RowTree, desde la 2.7.0 el mismo en todos los métodos): en cada tarea
resumen, el peso de cada subtarea directa es un % de la parte de la tarea resumen y no se propaga más arriba; las
subtareas sin peso se reparten el resto según su duración (un hito sin peso no recibe nada). Una fila sin pesos se
reparte según la duración de sus tareas. El método de la fila decide cómo se mide el avance de cada tarea en el %
ganado («Ponderado»: el método propio de cada tarea).

Multi-empresa: el peso de una tarea vale para cada empresa que participa en ella (es un % de la parte de la
empresa) y la duración de cada tarea se multiplica por la parte de la empresa (columna Empresa y % Reparto). Si una
empresa no participa en ninguna tarea de la fila, se usa el reparto de la fila completa (por si se le asigna BAC).
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import date, datetime, time

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
    equal_milestones: bool = False                  # «Hitos ponderados» sin pesos: los hitos pesan igual


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


# ----------------------------------------------------------------------------
# Pesos
# ----------------------------------------------------------------------------
WEIGHT_TOTAL = 100.0
NO_SHARE_FULL = "completo"          # las tareas hermanas con peso suman 100 o más
NO_SHARE_DURATION = "duracion"      # el resto se reparte por duración y la de la tarea es 0
_EPS = 1e-6


def task_weight(t: Task) -> float:
    """Peso de una tarea (campo de [campos] peso_hito; vacío y negativo = sin peso)."""
    return max(0.0, float(t.milestone_weight or 0.0))


def zero_weight(t: Task) -> bool:
    """Peso = 0 escrito en el campo (desde la 2.7.0): la tarea se hace, pero no cuenta en el libro EVM. Solo en un campo
    de texto se distingue del vacío (sin definir); en un campo numérico, el lector lee el 0 como vacío."""
    return t.milestone_weight is not None and t.milestone_weight == 0


def no_count(t: Task) -> bool:
    """La tarea no cuenta en el libro: tiene Peso = 0, o todas sus subtareas no cuentan."""
    return zero_weight(t) or (bool(t.children) and all(no_count(c) for c in t.children))


def is_zero_duration(t: Task) -> bool:
    """Hito: marcado como hito o de duración 0 (sin subtareas)."""
    return not t.children and (t.milestone or (t.duration is not None and t.duration <= 1e-9))


@dataclass
class WeightGroup:
    """Tarea resumen cuyas subtareas directas tienen pesos que no suman 100.

    Más de 100: error (se escalan a 100 y las subtareas sin peso no cuentan). Menos de 100 y sin subtareas sin peso
    que reciban el resto: se escalan a 100 (aviso)."""

    task: Task | None
    total: float


Pairs = list[tuple[Task, float]]


class RowTree:
    """Reparto de una fila del libro entre sus tareas (desde la 2.7.0, el mismo en todos los métodos).

    En cada tarea resumen (la de la fila y las de nivel inferior que reciben una parte), los pesos de sus subtareas
    directas son % de la parte de la tarea resumen; un peso no se propaga más arriba de su tarea resumen:
    - suman 100: cada subtarea con peso recibe su %; las subtareas sin peso, nada;
    - suman menos de 100: el resto se reparte entre las subtareas sin peso según su duración (un hito sin peso, de
      duración 0, no recibe nada); si no hay subtareas sin peso, los pesos se escalan a 100;
    - suman más de 100: se escalan a 100 (error de los datos) y las subtareas sin peso no reciben nada;
    - ninguna tiene peso: la parte se reparte según su duración (en partes iguales si todas duran 0).
    Una fila sin pesos se reparte, así, según la duración de sus tareas. Las tareas con Peso = 0 no cuentan.

    % ganado: todas las tareas de detalle de la fila. % planificado: las tareas del PV (units, de wbs.pv_selection:
    las de detalle con línea base y las tareas desglosadas después, como un bloque); el reparto por duración usa su
    duración planificada, y la parte de las tareas que no forman el PV (sin línea base) la reciben sus tareas hermanas
    de la misma tarea resumen (reparto local). units = None: el PV lo forman las tareas de detalle
    (--pv-fechas-actuales); [] = la fila no tiene % planificado.

    Multi-empresa: el peso de una tarea vale para cada empresa que participa en ella (es un % de la parte de la empresa
    en la tarea resumen) y la duración se multiplica por la parte de la empresa en cada tarea.

    «Hitos ponderados» sin ningún peso: el % planificado y el % ganado se reparten por igual entre los hitos (las
    tareas no cuentan).
    """

    def __init__(self, top: Task | None, leaves: list[Task], time_of: Callable[[Task], TaskTime],
                 units: list[Task] | None = None, method: Method | None = None):
        self.top = top
        self.time_of = time_of
        self.method = method or Method(PCT)
        self.units = units
        self._unit_ids = None if units is None else {u.uid for u in units}
        self.leaves: list[Task] = []
        inside: dict[int, Task] = {}
        orphans: list[Task] = []
        for t in leaves:
            if zero_weight(t):
                continue
            chain: list[Task] = []
            p = t.parent if t is not top else None
            while p is not None and p is not top and not zero_weight(p):
                chain.append(p)
                p = p.parent
            if p is not None and p is not top:          # dentro de una tarea con Peso = 0
                continue
            self.leaves.append(t)
            if p is None:                               # fuera de la rama de la fila, o la fila es una tarea de detalle
                orphans.append(t)
                continue
            inside[t.uid] = t
            inside.update((x.uid, x) for x in chain)
        self._inside = inside
        self._root = ([c for c in top.children if c.uid in inside] if top is not None else []) + orphans
        self.nodes: list[Task] = [x for x in list(inside.values()) + orphans if x is not top]
        self._leaf_ids = {t.uid for t in self.leaves}
        self.has_weights = any(task_weight(x) > 0 for x in self.nodes)
        self.negative = [x for x in self.nodes if (x.milestone_weight or 0.0) < 0]
        self.equal_milestones = (self.method.kind == MILESTONES and not self.has_weights
                                 and any(t.milestone for t in self.leaves))
        # Observaciones del reparto de la fila completa
        self.unweighted: list[tuple[Task, str]] = []    # tareas sin parte y la causa (NO_SHARE_*)
        self.groups: list[WeightGroup] = []             # tareas resumen con pesos que no suman 100
        self.equal: list[Task | None] = []              # tareas resumen repartidas en partes iguales (sin duración)
        self.pv_lost: list[Task] = []                   # tareas sin parte del PV por duración planificada 0
        self._memo: dict[tuple[int, bool], list[Task]] = {}
        self._full = self._split(lambda t: 1.0, True)

    # --- Árbol de la fila -------------------------------------------------------------------------------------
    def kids(self, t: Task | None) -> list[Task]:
        if t is None:
            return self._root
        return [c for c in t.children if c.uid in self._inside]

    def _inner(self, t: Task, pv: bool) -> list[Task]:
        """Tareas de detalle (o del PV, con pv) de la rama de t."""
        key = (t.uid, pv)
        if key not in self._memo:
            if pv and self._unit_ids is not None and t.uid in self._unit_ids:
                out = [t]
            else:
                ks = self.kids(t)
                if ks:
                    out = [x for c in ks for x in self._inner(c, pv)]
                elif pv and self._unit_ids is not None:
                    out = []
                else:
                    out = [t] if t.uid in self._leaf_ids else []
            self._memo[key] = out
        return self._memo[key]

    def branch_leaves(self, t: Task) -> list[Task]:
        """Tareas de detalle de la fila en la rama de t."""
        return self._inner(t, False)

    # --- Reparto ----------------------------------------------------------------------------------------------
    def _shares(self, t: Task | None, ks: list[Task], part: Callable[[Task], float], pv: bool,
                note: bool) -> tuple[dict[int, float], str, bool]:
        """Fracción de la parte de t que recibe cada subtarea (antes de quitar las que no tienen tareas incluidas), la
        causa por la que las demás no reciben nada y si se repartió en partes iguales."""
        weight = {c.uid: task_weight(c) for c in ks}
        total = sum(weight.values())
        rest = [c for c in ks if weight[c.uid] <= 0]
        scaled = total > WEIGHT_TOTAL + _EPS or (total > 0 and not rest)
        if note and scaled and abs(total - WEIGHT_TOTAL) > _EPS:
            self.groups.append(WeightGroup(t if t is not None else self.top, total))
        base = total if scaled else WEIGHT_TOTAL
        out = {c.uid: weight[c.uid] / base for c in ks if weight[c.uid] > 0}
        remainder = 0.0 if scaled else (WEIGHT_TOTAL - total) / WEIGHT_TOTAL
        equal = False
        if remainder > _EPS:
            minutes = {c.uid: sum(self.time_of(x).minutes * part(x) for x in self._inner(c, pv)) for c in rest}
            span = sum(minutes.values())
            if span > 0:
                out.update((c.uid, remainder * minutes[c.uid] / span) for c in rest)
            elif total <= 0:                                    # sin pesos y todas de duración 0: partes iguales
                found = [c for c in rest if any(part(x) > TOL for x in self._inner(c, pv))]
                out.update((c.uid, remainder / len(found)) for c in found)
                equal = len(found) > 1
        return out, (NO_SHARE_DURATION if remainder > _EPS else NO_SHARE_FULL), equal

    def _alloc(self, t: Task | None, amount: float, part: Callable[[Task], float], pv: bool, note: bool) -> Pairs:
        if t is not None:
            if pv and self._unit_ids is not None and t.uid in self._unit_ids:
                return [(t, amount)] if part(t) > TOL else []
            if not self.kids(t):
                if pv and self._unit_ids is not None:
                    return []
                return [(t, amount)] if part(t) > TOL else []
        # Multi-empresa: solo las subtareas en las que participa la empresa (los pesos son % de su parte)
        ks = [c for c in self.kids(t) if any(part(x) > TOL for x in self._inner(c, False))]
        if not ks:
            return []
        frac, why, equal = self._shares(t, ks, part, pv, note and not pv)
        if note and not pv:
            if equal:
                self.equal.append(t)
            self.unweighted.extend((c, why) for c in ks if frac.get(c.uid, 0.0) <= 0)
        elif note and pv:
            ev = self._shares(t, ks, part, False, False)[0]
            self.pv_lost.extend(c for c in ks if frac.get(c.uid, 0.0) <= 0 < ev.get(c.uid, 0.0)
                                and any(part(x) > TOL for x in self._inner(c, True)))
        parts = [(c, frac[c.uid]) for c in ks
                 if frac.get(c.uid, 0.0) > 0 and any(part(x) > TOL for x in self._inner(c, pv))]
        total = sum(f for _c, f in parts)
        return [x for c, f in parts for x in self._alloc(c, amount * f / total, part, pv, note)] if total > 0 else []

    def _split(self, part: Callable[[Task], float], note: bool) -> tuple[Pairs, Pairs]:
        """Tareas y fracción del % ganado y del % planificado (cada una suma 1, o vacía)."""
        if self.equal_milestones:
            ev = [t for t in self.leaves if t.milestone and part(t) > TOL]
            pool = self.leaves if self.units is None else self.units
            pv = [t for t in pool if t.milestone and part(t) > TOL]
            return [(t, 1.0 / len(ev)) for t in ev], [(t, 1.0 / len(pv)) for t in pv]
        ev = self._alloc(None, 1.0, part, False, note)
        if self._unit_ids is None:
            pv = ev
        elif self.top is not None and self.top.uid in self._unit_ids:
            pv = [(self.top, 1.0)] if part(self.top) > TOL else []
        else:
            pv = self._alloc(None, 1.0, part, True, note)
        return _normalized(ev), _normalized(pv)

    def full(self) -> tuple["RowWeights", list[tuple[TaskTime, float]]]:
        """% ganado y % planificado de la fila completa (todas las tareas al 100 %)."""
        ev, pv = self._full
        return self._weights(ev, [self.time_of(t) for t in self.leaves]), [(self.time_of(t), f) for t, f in pv]

    def company(self, share: Share | None) -> tuple["RowWeights", list[tuple[TaskTime, float]]]:
        """% ganado y % planificado de una empresa (share: su parte en cada tarea; None = la fila completa). Si la
        empresa no participa en ninguna tarea de la fila (o del PV), se usa el reparto de la fila completa (fallback
        en el % ganado)."""
        if share is None:
            return self.full()

        def part(t: Task) -> float:
            return max(0.0, min(1.0, share(t)))

        members = [self.time_of(t) for t in self.leaves if part(t) > TOL]
        ev, pv = self._split(part, False) if members else ([], [])
        full_ev, full_pv = self.full()
        rw = self._weights(ev, members) if ev else replace(full_ev, fallback=bool(self.leaves))
        return rw, ([(self.time_of(t), f) for t, f in pv] if pv else full_pv)

    def _weights(self, pairs: Pairs, members: list[TaskTime]) -> "RowWeights":
        return RowWeights({t.uid: f for t, f in pairs}, members, equal_milestones=self.equal_milestones)


def _normalized(pairs: Pairs) -> Pairs:
    s = sum(f for _t, f in pairs)
    return [(t, f / s) for t, f in pairs if f > 0] if s > TOL else []


def weighted_curve(pairs: list[tuple[TaskTime, float]]) -> dict[MonthKey, float]:
    """% planificado de cada mes (fracciones que suman 1 si la fila tiene % planificado)."""
    out: dict[MonthKey, float] = {}
    for tt, f in pairs:
        add_into(out, tt.dist, f)
    return out


# ----------------------------------------------------------------------------
# % ganado
# ----------------------------------------------------------------------------
def _progress(t: Task, physical: bool) -> float:
    """Avance de una tarea entre 0 y 1. Un hito cuenta entero cuando está cumplido."""
    if t.milestone:
        return 1.0 if is_complete(t) else 0.0
    v = t.physical_percent_complete if physical else t.percent_complete
    return max(0.0, min(1.0, v / 100.0))


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


_PCT, _PHYSICAL, _OWN = "pct", "fisico", "propio"


def _task_progress(tt: TaskTime, mode: str, cut: MonthKey | None) -> float:
    if mode == _OWN:
        elapsed = sum(v for k, v in tt.dist.items() if cut is None or k <= cut) if tt.dist else None
        return subtask_progress(tt.task, elapsed)
    return _progress(tt.task, mode == _PHYSICAL)


def _weighted_progress(rw: RowWeights, mode: str = _PCT, cut: MonthKey | None = None) -> float:
    if rw.weights:
        by_uid = {tt.task.uid: tt for tt in rw.members}
        value = sum(w * _task_progress(by_uid[uid], mode, cut) for uid, w in rw.weights.items() if uid in by_uid)
        return max(0.0, min(1.0, value))
    if rw.members:                      # sin pesos (tareas sin fechas): promedio simple
        return sum(_task_progress(tt, mode, cut) for tt in rw.members) / len(rw.members)
    return 0.0


def earned_fraction(method: Method, rw: RowWeights, pv_fraction: float,
                    cut: MonthKey | None = None) -> tuple[float, Issue | None]:
    """% ganado (0 a 1) de una empresa en una fila a la fecha de corte y una nota explicativa (sin la fila).

    rw: pesos de la empresa (o de la fila completa); pv_fraction: % planificado acumulado de la empresa a la
    fecha de corte (para el avance lineal); cut: mes de corte (para el avance lineal de una tarea de una fila
    «Ponderado»; None = toda su duración).
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
        return _weighted_progress(rw), None
    if kind in (PHYSICAL, UNITS):
        return _weighted_progress(rw, _PHYSICAL), None
    if kind == LOE:
        return max(0.0, min(1.0, pv_fraction)), None
    if kind == MILESTONES:
        milestones = [t for t in members if t.milestone]
        if milestones:
            # Cada hito cumplido suma su parte y las tareas con parte, su % completado (el mismo reparto que el %
            # planificado); sin pesos, los hitos pesan igual y las tareas no cuentan
            negative = any((t.milestone_weight or 0.0) < 0 for t in milestones)
            note = None
            if rw.equal_milestones:
                note = (Issue("ACT-029", message="hitos sin peso válido (pesos negativos): se ponderaron por igual")
                        if negative else Issue("ACT-030", message="hitos sin peso definido: se ponderaron por igual"))
            elif negative:
                note = Issue("ACT-031", message="hitos con peso negativo: se tomaron sin peso")
            return _weighted_progress(rw), note
        if any(t.physical_percent_complete > 0 for t in members):
            return _weighted_progress(rw, _PHYSICAL), Issue("ACT-032", message="sin hitos: se usó % físico completado")
        return _weighted_progress(rw), Issue("ACT-033", message="sin hitos ni % físico completado: se usó % completado")
    if kind == WEIGHTED:            # cada tarea con su propio método
        return _weighted_progress(rw, _OWN, cut), None
    if kind == UNDEFINED:
        return _weighted_progress(rw), Issue("ACT-034", message="método no definido: se usó % completado")
    return _weighted_progress(rw), Issue("ACT-035", message=f"método no reconocido '{method.text}': se usó % "
                                                            "completado")
