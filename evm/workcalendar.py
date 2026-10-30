"""Calendario laboral para repartir montos en el tiempo.

Con XML se usa el calendario del proyecto. Con COM, MS Project calcula el PV
con sus propios calendarios; aquí solo se usa el calendario estándar para las
tareas sin distribución mensual.
"""
from __future__ import annotations

import calendar as _cal
from datetime import date, datetime, time, timedelta

from evm.messages import Issue
from evm.model import MonthKey

Window = tuple[int, int]          # minutos desde las 00:00 (fin 1440 = 24:00)

STANDARD_DAY: list[Window] = [(8 * 60, 12 * 60), (13 * 60, 17 * 60)]


class WorkCalendar:
    """Jornada semanal + semanas laborales + excepciones por fecha."""

    def __init__(self, weekly: dict[int, list[Window]] | None = None, name: str = "Estándar"):
        # weekday(): 0 = lunes ... 6 = domingo
        if weekly is None:
            weekly = {d: list(STANDARD_DAY) for d in range(5)}
            weekly.update({5: [], 6: []})
        self.weekly = weekly
        self.name = name
        self.work_weeks: list[tuple[date, date, dict[int, list[Window]]]] = []
        self.exceptions: dict[date, list[Window]] = {}

    def add_holidays(self, days) -> None:
        for d in days:
            self.exceptions[d] = []

    def windows(self, d: date) -> list[Window]:
        if d in self.exceptions:
            return self.exceptions[d]
        for a, b, week in reversed(self.work_weeks):     # el calendario derivado tiene prioridad
            if a <= d <= b and d.weekday() in week:
                return week[d.weekday()]
        return self.weekly.get(d.weekday(), [])

    def is_working_day(self, d: date) -> bool:
        return bool(self.windows(d))

    def minutes_by_month(self, start: datetime, finish: datetime) -> dict[MonthKey, float]:
        """Minutos laborables del intervalo [start, finish) agrupados por mes."""
        out: dict[MonthKey, float] = {}
        d = start.date()
        last = finish.date()
        while d <= last:
            day0 = datetime.combine(d, time(0, 0))
            for a, b in self.windows(d):
                ws = day0 + timedelta(minutes=a)
                we = day0 + timedelta(minutes=b)
                lo, hi = max(ws, start), min(we, finish)
                if hi > lo:
                    k = (d.year, d.month)
                    out[k] = out.get(k, 0.0) + (hi - lo).total_seconds() / 60.0
            d += timedelta(days=1)
        return out


def elapsed_by_month(start: datetime, finish: datetime) -> dict[MonthKey, float]:
    """Minutos transcurridos (días corridos) del intervalo [start, finish) por mes."""
    out: dict[MonthKey, float] = {}
    cur = start
    while cur < finish:
        y, m = cur.year, cur.month
        nxt = datetime(y + (m == 12), 1 if m == 12 else m + 1, 1)
        b = min(nxt, finish)
        out[(y, m)] = out.get((y, m), 0.0) + (b - cur).total_seconds() / 60.0
        cur = b
    return out


# ----------------------------------------------------------------------------
# Lectura desde el XML de MS Project (<Calendars>)
# ----------------------------------------------------------------------------
def _minutes(text: str) -> int | None:
    try:
        h, m, *_ = (int(x) for x in text.split(":"))
    except (ValueError, AttributeError):
        return None
    return h * 60 + m


def _windows(r, parent) -> list[Window]:
    out: list[Window] = []
    wt = parent.find(r.q("WorkingTimes"))
    if wt is None:
        return out
    for w in wt.findall(r.q("WorkingTime")):
        a, b = _minutes(r.text(w, "FromTime")), _minutes(r.text(w, "ToTime"))
        if a is None or b is None:
            continue
        if b <= a:
            b = 24 * 60     # 'ToTime' 00:00 = fin del día
        out.append((a, b))
    return out


def _weekdays(r, parent) -> tuple[dict[int, list[Window]], list[tuple[date, date, list[Window]]]]:
    """Lee <WeekDays>: jornada por día de la semana y excepciones del formato antiguo (DayType 0)."""
    week: dict[int, list[Window]] = {}
    legacy: list[tuple[date, date, list[Window]]] = []
    wds = parent.find(r.q("WeekDays"))
    if wds is None:
        return week, legacy
    for wd in wds.findall(r.q("WeekDay")):
        try:
            day_type = int(r.text(wd, "DayType"))
        except ValueError:
            continue
        working = r.text(wd, "DayWorking") == "1"
        wins = _windows(r, wd) if working else []
        if working and not wins:
            wins = list(STANDARD_DAY)
        if day_type == 0:
            tp = wd.find(r.q("TimePeriod"))
            a, b = r.date(tp, "FromDate"), r.date(tp, "ToDate")
            if a and b:
                legacy.append((a.date(), b.date(), wins))
        elif 1 <= day_type <= 7:
            week[(day_type - 2) % 7] = wins     # 1 = domingo ... 7 = sábado
    return week, legacy


def _exception_days(r, ex) -> list[date] | None:
    """Días de una excepción del calendario; None si no se puede interpretar."""
    tp = ex.find(r.q("TimePeriod"))
    a, b = r.date(tp, "FromDate"), r.date(tp, "ToDate")
    if a is None or b is None:
        return None
    a, b = a.date(), b.date()
    try:
        kind = int(r.text(ex, "Type") or 1)
    except ValueError:
        kind = 1
    try:
        period = max(1, int(r.text(ex, "Period") or 1))
    except ValueError:
        period = 1
    days: list[date] = []
    if kind in (1, 7):                                   # diaria
        d = a
        while d <= b:
            days.append(d)
            d += timedelta(days=period)
    elif kind == 2:                                      # anual por fecha
        try:
            month = int(r.text(ex, "Month")) + 1
            mday = int(r.text(ex, "MonthDay"))
        except ValueError:
            return None
        for y in range(a.year, b.year + 1):
            if 1 <= month <= 12:
                dd = date(y, month, min(mday, _cal.monthrange(y, month)[1]))
                if a <= dd <= b:
                    days.append(dd)
    elif kind == 4:                                      # mensual por fecha (p. ej. día 15)
        try:
            mday = int(r.text(ex, "MonthDay"))
        except ValueError:
            return None
        y, m = a.year, a.month
        while date(y, m, 1) <= b:
            dd = date(y, m, min(mday, _cal.monthrange(y, m)[1]))
            if a <= dd <= b:
                days.append(dd)
            m += period
            while m > 12:
                y, m = y + 1, m - 12
    elif kind == 6:                                      # semanal (máscara de días)
        try:
            mask = int(r.text(ex, "DaysOfWeek") or 0)
        except ValueError:
            mask = 0
        week_start = a - timedelta(days=(a.weekday() + 1) % 7)   # domingo
        w = 0
        while week_start + timedelta(weeks=w) <= b:
            if w % period == 0:
                for i in range(7):                       # bit 0 = domingo
                    dd = week_start + timedelta(weeks=w, days=i)
                    if mask & (1 << i) and a <= dd <= b:
                        days.append(dd)
            w += 1
    else:               # tipos 3 y 5 (anual y mensual relativas, p. ej. 'primer lunes') y otros
        return None
    return days


def calendar_from_xml(r, root, uid: int | None, warnings: list[Issue] | None = None) -> WorkCalendar | None:
    """Calendario del proyecto (con herencia del calendario base) o None si no existe.

    Las excepciones que no se pueden interpretar se omiten y se informan en `warnings`.
    """
    cals = root.find(r.q("Calendars"))
    if cals is None or uid is None:
        return None
    by_uid = {}
    for c in cals.findall(r.q("Calendar")):
        try:
            by_uid[int(r.text(c, "UID"))] = c
        except ValueError:
            continue
    if uid not in by_uid:
        return None

    chain = []
    seen = set()
    cur = uid
    while cur in by_uid and cur not in seen:
        seen.add(cur)
        chain.append(by_uid[cur])
        try:
            cur = int(r.text(by_uid[cur], "BaseCalendarUID") or -1)
        except ValueError:
            break

    cal = WorkCalendar(weekly={})
    cal.name = r.text(chain[0], "Name")
    # Del calendario base al derivado: el derivado reemplaza lo que define
    for c in reversed(chain):
        week, legacy = _weekdays(r, c)
        cal.weekly.update(week)
        for a, b, wins in legacy:
            d = a
            while d <= b:
                cal.exceptions[d] = wins
                d += timedelta(days=1)
        wws = c.find(r.q("WorkWeeks"))
        if wws is not None:
            for ww in wws.findall(r.q("WorkWeek")):
                tp = ww.find(r.q("TimePeriod"))
                a, b = r.date(tp, "FromDate"), r.date(tp, "ToDate")
                wk, _ = _weekdays(r, ww)
                if a and b and wk:
                    cal.work_weeks.append((a.date(), b.date(), wk))
        exs = c.find(r.q("Exceptions"))
        if exs is not None:
            for ex in exs.findall(r.q("Exception")):
                working = r.text(ex, "DayWorking") == "1"
                wins = _windows(r, ex) if working else []
                if working and not wins:
                    wins = list(STANDARD_DAY)
                days = _exception_days(r, ex)
                if days is None:
                    if warnings is not None:
                        name = r.text(ex, "Name")
                        label = f"'{name}'" if name else "sin nombre"
                        warnings.append(Issue("LEC-018", message=(
                            f"Calendario '{r.text(c, 'Name')}': se omite la excepción {label} "
                            f"(tipo {r.text(ex, 'Type') or '?'}) porque su periodicidad no se interpreta. Solo "
                            "afecta el reparto de la duración de las tareas entre los meses (% planificado).")))
                    continue
                for d in days:
                    cal.exceptions[d] = wins
    for d in range(7):
        cal.weekly.setdefault(d, list(STANDARD_DAY) if d < 5 else [])
    if not any(cal.weekly.values()) and not cal.work_weeks:
        return None
    return cal
