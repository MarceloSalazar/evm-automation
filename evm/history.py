"""% ganado de los meses hasta la fecha de estado al generar el libro (desde la 2.7.0).

MS Project guarda el avance acumulado a la fecha de estado y las fechas reales, no el avance de cada mes. Para cada
mes del horizonte hasta el de la fecha de estado se reconstruye el estado de cada tarea de detalle al cierre del mes
y se calcula el % ganado de cada fila (y de cada empresa) con el mismo cálculo que 'actualizar'
(percent.earned_fraction):

- una tarea sin avance a la fecha de estado no tiene avance en ningún mes;
- antes de su comienzo real no tiene avance; desde su fin real está terminada;
- entre los dos (o, si no terminó, entre el comienzo real y la fecha de estado) su avance a la fecha de estado se
  reparte en línea recta según el tiempo laborable (o los días corridos, con distribucion = "calendario").

Así las fórmulas fijas (0/100, 50/50...) y los hitos ponderados ganan su valor en el mes de su comienzo o de su fin
real (exacto), el avance lineal gana el PV de cada mes (exacto) y el % completado y el % físico completado se
reparten en línea recta (estimación). En el mes de la fecha de estado el % ganado acumulado es el mismo que
calcularía 'actualizar' con ese archivo.
"""
from __future__ import annotations

import calendar as _calendar
from dataclasses import dataclass, field, replace
from datetime import date, datetime

from evm import percent as pct
from evm.messages import Issue
from evm.methods import is_complete, is_started, parse_method
from evm.model import MonthKey, Task
from evm.workcalendar import WorkCalendar, elapsed_by_month


@dataclass
class EarnedHistory:
    """Resultado de la carga: la fecha de estado y su mes (None = no se cargó) y las observaciones."""

    status: date | None = None
    index: int | None = None
    issues: list[Issue] = field(default_factory=list)

    @property
    def loaded(self) -> bool:
        return self.index is not None


def _month_end(d: date) -> datetime:
    return datetime(d.year, d.month, _calendar.monthrange(d.year, d.month)[1], 23, 59, 59)


def _without_progress(t: Task) -> Task:
    return replace(t, percent_complete=0.0, physical_percent_complete=0.0, actual_start=None, actual_finish=None)


class _TaskProgress:
    """Avance de una tarea de detalle al cierre de un mes anterior al de la fecha de estado."""

    def __init__(self, t: Task, status: datetime, calendar: WorkCalendar, mode: str):
        self.task = t
        self.start: datetime | None = None
        self.finish: datetime | None = None
        self.cumulative: dict[MonthKey, float] = {}     # fracción del avance acumulada al cierre de cada mes
        if not (is_started(t) or is_complete(t)):
            return
        self.start = t.actual_start or t.start
        if self.start is None:              # comenzada sin fecha de comienzo: todo su avance, en el mes de corte
            return
        done = is_complete(t)
        self.finish = (t.actual_finish or t.finish) if done else None
        end = self.finish if self.finish is not None else status
        if end <= self.start:
            return
        weights = calendar.minutes_by_month(self.start, end) if mode == "habiles" else {}
        if not sum(weights.values()):
            weights = elapsed_by_month(self.start, end)
        total = sum(weights.values())
        acc = 0.0
        for key in sorted(weights):
            acc += weights[key]
            self.cumulative[key] = acc / total if total else 1.0

    def at(self, when: datetime) -> Task:
        t = self.task
        if self.start is None:
            return t if not (is_started(t) or is_complete(t)) else _without_progress(t)
        if when < self.start:
            return _without_progress(t)
        if self.finish is not None and when >= self.finish:
            return t
        key = (when.year, when.month)
        frac = max((v for k, v in self.cumulative.items() if k <= key), default=0.0)
        return replace(t, percent_complete=t.percent_complete * frac,
                       physical_percent_complete=t.physical_percent_complete * frac,
                       actual_start=self.start, actual_finish=None)


def earned_history(report, cfg) -> EarnedHistory:
    """Calcula el % ganado de cada mes hasta la fecha de estado (ReportRow.ev_hist) si el archivo tiene avance."""
    out = EarnedHistory()
    inputs = [r for r in report.rows if r.is_input and r.weights]
    if not any(is_started(t) or is_complete(t) for r in inputs for t in r.leaves):
        return out
    statuses = {s.name: s.project.status_date for s in report.sources}
    later = "se cargará con 'actualizar'"
    missing = [n for n, d in statuses.items() if d is None]
    if missing:
        which = ("El archivo" if missing == [""] else
                 ("El archivo " if len(missing) == 1 else "Los archivos ") + ", ".join(missing))
        out.issues.append(Issue("LIB-017", message=f"{which} tiene avance pero no fecha de estado: no se cargó el % "
                                                   f"ganado, que {later} (se debe definir Proyecto > Fecha de estado "
                                                   "en MS Project)."))
        return out
    days = {n: d.date() for n, d in statuses.items()}
    if len(set(days.values())) > 1:
        shown = ", ".join(f"{n}: {d:%d/%m/%Y}" for n, d in days.items())
        out.issues.append(Issue("LIB-017", message=f"Los archivos tienen fechas de estado distintas ({shown}): no se "
                                                   f"cargó el % ganado, que {later}."))
        return out
    status = next(iter(days.values()))
    idx = report.horizon.index(status)
    if not 0 <= idx < report.horizon.months:
        out.issues.append(Issue("LIB-017", message=f"La fecha de estado {status:%d/%m/%Y} está fuera del horizonte "
                                                   f"del libro ({report.horizon.start:%m/%Y} a "
                                                   f"{report.horizon.end:%m/%Y}): no se cargó el % ganado."))
        return out

    status_end = datetime(status.year, status.month, status.day, 23, 59, 59)
    projects = {s.name: s.project for s in report.sources}
    calendars = {n: pct.calendar_for(p, cfg.project_config(n)) for n, p in projects.items()}
    default = next(iter(calendars.values()))
    month_ends = [_month_end(report.horizon.month_start(i)) for i in range(idx)]
    states: dict[tuple[str, int], _TaskProgress] = {}
    for row in inputs:
        cal = calendars.get(row.source, default)
        method = parse_method(row.method)

        def state(t: Task) -> _TaskProgress:
            key = (row.source, t.uid)
            if key not in states:
                states[key] = _TaskProgress(t, status_end, cal, cfg.distribution)
            return states[key]

        for role in row.weights:
            rw = row.weights[role]
            pv = row.pct_vec.get(role, [])
            values: list[float] = []
            prev = 0.0
            for i in range(idx + 1):
                if i < idx:
                    members = [replace(tt, task=state(tt.task).at(month_ends[i])) for tt in rw.members]
                    rw_i = replace(rw, members=members)
                else:
                    rw_i = rw                       # mes de la fecha de estado: el avance de MS Project
                frac, _note = pct.earned_fraction(method, rw_i, sum(pv[: i + 1]), report.horizon.key(i))
                values.append(round(frac - prev, 12))
                prev = frac
            row.ev_hist[role] = values

    first = report.horizon.month_start(0)
    months = f"{idx + 1} {'mes' if idx == 0 else 'meses'}"
    out.status, out.index = status, idx
    out.issues.append(Issue("LIB-016", message=(
        f"Se cargó el % ganado de {months} ({first:%m/%Y} a {status:%m/%Y}), hasta la fecha de estado "
        f"{status:%d/%m/%Y}, que es la fecha de corte del Resumen. El acumulado a la fecha de estado es el de MS "
        "Project; el de cada mes anterior se reconstruyó con las fechas reales: exacto en las fórmulas fijas, los "
        "hitos ponderados y el avance lineal, y en línea recta desde el comienzo real en el % completado y el % físico "
        "completado. El AC de esos meses se debe ingresar en la hoja AC.")))
    return out
