"""Constructor mínimo de XML de MS Project (formato MSPDI) para las pruebas.

Genera solo los elementos que leen evm.reader_xml y evm.workcalendar. Los
montos se indican en unidades de moneda y se escriben multiplicados por 100,
como lo hace MS Project. Las fechas pueden ser date, datetime o texto ISO.

Ejemplo:

    xml = project_xml([
        task_xml(0, 0, "Proyecto"),
        task_xml(1, 1, "A", wbs="1", start=s, finish=f, cost=100,
                 baselines=[(0, s, f, 100)], timephased=[(10, s, f, 100)],
                 ext={TEXT10: "50/50"}),
    ], definitions=[(TEXT10, "Text10", "Método EVM")])
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from xml.sax.saxutils import escape

from evm.fields import TASK_NUMBER_IDS, TASK_TEXT_IDS

NS = "http://schemas.microsoft.com/project"

# FieldID de campos personalizados de tarea
TEXT2 = TASK_TEXT_IDS[2]
TEXT3 = TASK_TEXT_IDS[3]
TEXT10 = TASK_TEXT_IDS[10]
NUMBER1 = TASK_NUMBER_IDS[1]
NUMBER2 = TASK_NUMBER_IDS[2]

# FieldID del rango de campos de RECURSO (0x0C400000 = 205520896 en adelante). En el XML
# las definiciones de campos de tarea, recurso y asignación comparten <ExtendedAttributes>.
RESOURCE_FIELD_IDS = (205520904, 205520905)

# Jornada estándar: lunes a viernes 08:00-12:00 y 13:00-17:00
STANDARD_HOURS = (("08:00:00", "12:00:00"), ("13:00:00", "17:00:00"))


def iso(value: date | datetime | str | None, end_of_day: bool = False) -> str:
    """Fecha en formato MSPDI. Un date se completa con 00:00 (o 23:59 si end_of_day)."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    return f"{value.isoformat()}T{'23:59:00' if end_of_day else '00:00:00'}"


def money(amount: float) -> str:
    """Monto en MSPDI: multiplicado por 100."""
    return f"{amount * 100:.0f}" if float(amount * 100).is_integer() else f"{amount * 100:.2f}"


def _el(name: str, value) -> str:
    return "" if value is None else f"<{name}>{escape(str(value))}</{name}>"


def _working_times(hours) -> str:
    if not hours:
        return ""
    return "<WorkingTimes>" + "".join(
        f"<WorkingTime><FromTime>{a}</FromTime><ToTime>{b}</ToTime></WorkingTime>" for a, b in hours) + "</WorkingTimes>"


# ----------------------------------------------------------------------------
# Calendarios
# ----------------------------------------------------------------------------
def weekday_xml(day_type: int, hours=STANDARD_HOURS, working: bool = True,
                period: tuple | None = None) -> str:
    """<WeekDay>: day_type 1 = domingo ... 7 = sábado; 0 = excepción del formato antiguo
    (en ese caso `period` = (desde, hasta))."""
    tp = ""
    if period is not None:
        tp = (f"<TimePeriod><FromDate>{iso(period[0])}</FromDate>"
              f"<ToDate>{iso(period[1], end_of_day=True)}</ToDate></TimePeriod>")
    return (f"<WeekDay><DayType>{day_type}</DayType><DayWorking>{1 if working else 0}</DayWorking>{tp}"
            f"{_working_times(hours) if working else ''}</WeekDay>")


def standard_week() -> list[str]:
    """Semana estándar: lunes a viernes laborables, sábado y domingo no laborables."""
    return [weekday_xml(d, working=2 <= d <= 6) for d in range(1, 8)]


def exception_xml(start, finish, kind: int | None = 1, working: bool = False, hours=(), name: str = "Excepción",
                  **fields) -> str:
    """<Exception> del calendario. kind = Type (1 diaria, 2 anual, 3 anual relativa, 4 mensual,
    5 mensual relativa, 6 semanal). `fields` agrega elementos como Period=2, DaysOfWeek=10,
    Month=1 (base 0), MonthDay=29, MonthItem=4, MonthPosition=4."""
    extra = "".join(_el(k, v) for k, v in fields.items())
    return (f"<Exception><TimePeriod><FromDate>{iso(start)}</FromDate><ToDate>{iso(finish, end_of_day=True)}"
            f"</ToDate></TimePeriod>{_el('Name', name)}{_el('Type', kind)}{extra}"
            f"<DayWorking>{1 if working else 0}</DayWorking>{_working_times(hours) if working else ''}</Exception>")


def work_week_xml(start, finish, weekdays: list[str], name: str = "Semana laboral") -> str:
    """<WorkWeek>: jornada distinta de la predeterminada entre dos fechas."""
    return (f"<WorkWeek><TimePeriod><FromDate>{iso(start)}</FromDate><ToDate>{iso(finish, end_of_day=True)}"
            f"</ToDate></TimePeriod>{_el('Name', name)}<WeekDays>{''.join(weekdays)}</WeekDays></WorkWeek>")


def calendar_xml(uid: int = 1, name: str = "Estándar", base: int = -1, weekdays: list[str] | None = None,
                 exceptions=(), work_weeks=()) -> str:
    """<Calendar> con su semana (por defecto, la estándar), semanas laborales y excepciones."""
    weekdays = standard_week() if weekdays is None else weekdays
    ww = f"<WorkWeeks>{''.join(work_weeks)}</WorkWeeks>" if work_weeks else ""
    ex = f"<Exceptions>{''.join(exceptions)}</Exceptions>" if exceptions else ""
    return (f"<Calendar><UID>{uid}</UID>{_el('Name', name)}<IsBaseCalendar>1</IsBaseCalendar>"
            f"<BaseCalendarUID>{base}</BaseCalendarUID><WeekDays>{''.join(weekdays)}</WeekDays>{ww}{ex}</Calendar>")


# ----------------------------------------------------------------------------
# Tareas y proyecto
# ----------------------------------------------------------------------------
def task_xml(uid: int, level: int | None, name: str = "Tarea", *, tid: int | None = None, wbs: str | None = None,
             start=None, finish=None, cost: float | None = None, baselines=(), timephased=(),
             ext: dict[int, str] | None = None, extra: str = "") -> str:
    """<Task>.

    level None omite <OutlineLevel>. baselines: tuplas (número, inicio, fin, costo) con
    inicio/fin None si no existen. timephased: tuplas (tipo, inicio, fin, monto) de
    <TimephasedData> (tipo 10 = costo de línea base; 19 + 6 * (N - 1) = línea base N).
    ext: {FieldID: valor} de campos personalizados. extra: XML agregado tal cual.
    """
    parts = [f"<Task><UID>{uid}</UID><ID>{uid if tid is None else tid}</ID>{_el('Name', name)}",
             _el("OutlineLevel", level), _el("WBS", wbs), _el("Start", iso(start) or None),
             _el("Finish", iso(finish) or None), _el("Cost", None if cost is None else money(cost)), extra]
    for fid, value in (ext or {}).items():
        parts.append(f"<ExtendedAttribute><FieldID>{fid}</FieldID>{_el('Value', value)}</ExtendedAttribute>")
    for number, bl_start, bl_finish, bl_cost in baselines:
        parts.append(f"<Baseline><Number>{number}</Number>{_el('Start', iso(bl_start) or None)}"
                     f"{_el('Finish', iso(bl_finish) or None)}<Cost>{money(bl_cost)}</Cost></Baseline>")
    for kind, seg_start, seg_finish, amount in timephased:
        parts.append(f"<TimephasedData><Type>{kind}</Type><UID>{uid}</UID><Start>{iso(seg_start)}</Start>"
                     f"<Finish>{iso(seg_finish)}</Finish><Unit>2</Unit><Value>{money(amount)}</Value>"
                     "</TimephasedData>")
    parts.append("</Task>")
    return "".join(parts)


def project_xml(tasks: list[str] | None, *, title: str | None = "Prueba", header: str = "", definitions=(),
                calendars=(), calendar_uid: int | None = None) -> str:
    """<Project> completo.

    tasks None omite <Tasks>. header: elementos de proyecto agregados tal cual (p. ej.
    <BaselineForEarnedValue>). definitions: tuplas (FieldID, FieldName, alias) de
    <ExtendedAttributes>; alias "" o None si el campo no tiene alias.
    """
    defs = ""
    if definitions:
        defs = "<ExtendedAttributes>" + "".join(
            f"<ExtendedAttribute><FieldID>{fid}</FieldID>{_el('FieldName', fname)}{_el('Alias', alias or None)}"
            "</ExtendedAttribute>" for fid, fname, alias in definitions) + "</ExtendedAttributes>"
    cals = f"<Calendars>{''.join(calendars)}</Calendars>" if calendars else ""
    body = "" if tasks is None else f"<Tasks>{''.join(tasks)}</Tasks>"
    return (f'<?xml version="1.0" encoding="UTF-8"?>\n<Project xmlns="{NS}">{_el("Title", title)}'
            f"{_el('CalendarUID', calendar_uid)}{header}{defs}{cals}{body}</Project>")


def write_xml(path: Path, text: str) -> Path:
    path.write_text(text, encoding="utf-8")
    return path
