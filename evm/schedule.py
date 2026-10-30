"""Plazo según el cronograma de MS Project: fin previsto, desvío de fin, holgura y camino crítico de cada fila.

El SPI compara valores, no fechas: no indica si un atraso afecta la fecha de fin. Cada lectura de MS Project
('generar' y 'actualizar') guarda en el libro, para cada fila de la EDT, con sus tareas de detalle (las mismas
del % ganado):

- fin de línea base y fin previsto (o real): el último de sus tareas;
- desvío de fin: días laborables entre los dos (> 0 = atraso), con el calendario que reparte la duración;
- holgura total: la menor de sus tareas sin terminar, en días laborables (la calcula MS Project);
- crítica: alguna de sus tareas sin terminar es crítica según MS Project.

Las filas resumen y el proyecto combinan los valores de sus filas.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

from evm.methods import is_complete
from evm.model import Task
from evm.workcalendar import WorkCalendar

CRITICAL_YES, CRITICAL_NO, FINISHED = "Sí", "No", "Terminado"


@dataclass
class RowSchedule:
    """Plazo de una fila (o del proyecto)."""

    baseline_finish: datetime | None = None
    finish: datetime | None = None
    variance: int | None = None         # días laborables entre el fin de línea base y el fin previsto (> 0 = atraso)
    slack: float | None = None          # holgura total mínima de las tareas sin terminar (días laborables)
    critical: bool = False              # alguna tarea sin terminar es crítica
    open_tasks: bool = False            # tiene tareas sin terminar

    @property
    def status(self) -> str:
        """Texto de la columna Crítica: 'Sí', 'No', 'Terminado' ('' sin tareas con fechas)."""
        if self.finish is None:
            return ""
        if not self.open_tasks:
            return FINISHED
        return CRITICAL_YES if self.critical else CRITICAL_NO

    @property
    def late_critical(self) -> bool:
        """Fila crítica con el fin atrasado respecto de la línea base."""
        return self.open_tasks and self.critical and (self.variance or 0) > 0


def finished(t: Task) -> bool:
    """Tarea terminada: el mismo criterio que el % ganado (methods.is_complete)."""
    return is_complete(t)


def _day(d: datetime | date) -> date:
    return d.date() if isinstance(d, datetime) else d


def working_days(calendar: WorkCalendar, a: datetime | date, b: datetime | date) -> int:
    """Días laborables de a a b: positivo si b es posterior, negativo si es anterior. Cuenta los días laborables
    posteriores a la fecha inicial hasta la final inclusive (de lunes a miércoles: 2)."""
    x, y = _day(a), _day(b)
    if x == y:
        return 0
    sign = 1 if y > x else -1
    lo, hi = (x, y) if sign > 0 else (y, x)
    days = 0
    d = lo + timedelta(days=1)
    while d <= hi:
        if calendar.is_working_day(d):
            days += 1
        d += timedelta(days=1)
    return sign * days


def _variance(calendar: WorkCalendar, baseline: datetime | None, finish: datetime | None) -> int | None:
    if baseline is None or finish is None:
        return None
    return working_days(calendar, baseline, finish)


def row_schedule(leaves: list[Task], calendar: WorkCalendar, minutes_per_day: float) -> RowSchedule:
    """Plazo de una fila con sus tareas de detalle."""
    baselines = [t.baseline_finish for t in leaves if t.has_baseline and t.baseline_finish]
    finishes = [t.finish for t in leaves if t.finish]
    pending = [t for t in leaves if not finished(t)]
    slacks = [t.total_slack / minutes_per_day for t in pending if t.total_slack is not None]
    out = RowSchedule(max(baselines) if baselines else None, max(finishes) if finishes else None,
                      slack=round(min(slacks), 2) if slacks else None,
                      critical=any(t.critical for t in pending), open_tasks=bool(pending))
    out.variance = _variance(calendar, out.baseline_finish, out.finish)
    return out


def combine(parts: list[RowSchedule], calendar: WorkCalendar) -> RowSchedule:
    """Plazo de una fila resumen (o del proyecto) con el de sus filas."""
    baselines = [p.baseline_finish for p in parts if p.baseline_finish]
    finishes = [p.finish for p in parts if p.finish]
    slacks = [p.slack for p in parts if p.slack is not None]
    out = RowSchedule(max(baselines) if baselines else None, max(finishes) if finishes else None,
                      slack=min(slacks) if slacks else None, critical=any(p.critical for p in parts),
                      open_tasks=any(p.open_tasks for p in parts))
    out.variance = _variance(calendar, out.baseline_finish, out.finish)
    return out


@dataclass
class ScheduleSummary:
    """Plazo del proyecto y de cada fila, para escribirlo en el libro."""

    status_date: date | None            # fecha de estado de la lectura (None = sin fecha de estado)
    rows: dict[int, RowSchedule]        # fila de las hojas mensuales -> plazo (sin las filas sin tareas)
    total: RowSchedule
    critical: int = 0                   # filas de ingreso críticas sin terminar
    late_critical: int = 0              # de ellas, con el fin atrasado

    @classmethod
    def build(cls, status_date: date | None, rows: dict[int, RowSchedule], inputs: list[int],
              children: dict[int, list[int]], calendars: dict[int, WorkCalendar],
              default: WorkCalendar) -> "ScheduleSummary":
        """Completa las filas resumen (de abajo hacia arriba) y el total.

        rows: plazo de las filas de ingreso; children: fila resumen -> sus filas hijas (en orden, las resumen
        después de sus hijas no importa); calendars: calendario de cada fila (varios archivos).
        """
        out = dict(rows)

        def visit(r: int) -> RowSchedule | None:
            if r in out:
                return out[r]
            kids = [x for x in (visit(c) for c in children.get(r, [])) if x is not None]
            if not kids:
                return None
            out[r] = combine(kids, calendars.get(r, default))
            return out[r]

        tops = [r for r in children if not any(r in v for v in children.values())]
        for r in list(children):
            visit(r)
        level1 = [out[r] for r in tops if r in out] + [out[r] for r in inputs if r in out
                                                      and not any(r in v for v in children.values())]
        latest = max((s for s in level1 if s.finish), key=lambda s: s.finish, default=None)
        cal = default
        if latest is not None:
            cal = next((calendars.get(r, default) for r, s in out.items() if s is latest), default)
        total = combine(level1, cal)
        crit = [out[r] for r in inputs if r in out and out[r].open_tasks and out[r].critical]
        return cls(status_date, out, total, len(crit), sum(1 for s in crit if (s.variance or 0) > 0))


def from_report(report, cfg, row_of) -> ScheduleSummary:
    """Plazo de las filas del informe de 'generar' (row_of: ReportRow -> fila de las hojas mensuales)."""
    from evm.percent import calendar_for
    sources = {s.name: s.project for s in report.sources}
    cals = {name: calendar_for(project, cfg) for name, project in sources.items()}
    default = next(iter(cals.values()), WorkCalendar())
    rows: dict[int, RowSchedule] = {}
    inputs: list[int] = []
    children: dict[int, list[int]] = {}
    calendars: dict[int, WorkCalendar] = {}
    for r in report.rows:
        n = row_of(r)
        calendars[n] = cals.get(r.source, default)
        if r.is_input:
            inputs.append(n)
            if r.leaves:
                rows[n] = row_schedule(r.leaves, calendars[n], sources[r.source].minutes_per_day
                                       if r.source in sources else 480.0)
        else:
            children[n] = [row_of(c) for c in r.children]
    status = next((p.status_date for p in sources.values() if p.status_date), None)
    return ScheduleSummary.build(_day(status) if status else None, rows, inputs, children, calendars, default)
