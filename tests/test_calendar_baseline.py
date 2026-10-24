from datetime import date, datetime

from evm.config import Config
from evm.reader_xml import read_xml

XML = """<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
<Title>Calendario</Title><CalendarUID>7</CalendarUID>{bfe}
<Calendars>
 <Calendar><UID>1</UID><Name>Standard</Name><IsBaseCalendar>1</IsBaseCalendar><BaseCalendarUID>-1</BaseCalendarUID>
  <WeekDays>
   <WeekDay><DayType>1</DayType><DayWorking>0</DayWorking></WeekDay>
   <WeekDay><DayType>2</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
   <WeekDay><DayType>3</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
   <WeekDay><DayType>4</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
   <WeekDay><DayType>5</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
   <WeekDay><DayType>6</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>12:00:00</ToTime></WorkingTime><WorkingTime><FromTime>13:00:00</FromTime><ToTime>17:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay>
   <WeekDay><DayType>7</DayType><DayWorking>0</DayWorking></WeekDay>
  </WeekDays>
  <Exceptions>
   <Exception><TimePeriod><FromDate>2026-12-25T00:00:00</FromDate><ToDate>2028-12-25T23:59:00</ToDate></TimePeriod>
    <Occurrences>3</Occurrences><Name>Navidad</Name><Type>2</Type><Month>11</Month><MonthDay>25</MonthDay><DayWorking>0</DayWorking></Exception>
   <Exception><TimePeriod><FromDate>2027-02-15T00:00:00</FromDate><ToDate>2027-02-16T23:59:00</ToDate></TimePeriod>
    <Occurrences>2</Occurrences><Name>Carnaval</Name><Type>1</Type><DayWorking>0</DayWorking></Exception>
  </Exceptions>
 </Calendar>
 <Calendar><UID>7</UID><Name>Turno sábado</Name><IsBaseCalendar>1</IsBaseCalendar><BaseCalendarUID>1</BaseCalendarUID>
  <WeekDays><WeekDay><DayType>7</DayType><DayWorking>1</DayWorking><WorkingTimes><WorkingTime><FromTime>08:00:00</FromTime><ToTime>13:00:00</ToTime></WorkingTime></WorkingTimes></WeekDay></WeekDays>
 </Calendar>
</Calendars>
<Tasks>
<Task><UID>0</UID><ID>0</ID><Name>Calendario</Name><OutlineLevel>0</OutlineLevel></Task>
<Task><UID>1</UID><ID>1</ID><Name>A</Name><WBS>1</WBS><OutlineLevel>1</OutlineLevel>
<Start>2026-01-05T08:00:00</Start><Finish>2026-01-09T17:00:00</Finish><Cost>10000</Cost>
<Baseline><Number>0</Number><Start>2026-01-05T08:00:00</Start><Finish>2026-01-09T17:00:00</Finish><Cost>10000</Cost></Baseline>
<Baseline><Number>3</Number><Start>2026-02-02T08:00:00</Start><Finish>2026-02-06T17:00:00</Finish><Cost>30000</Cost></Baseline>
</Task>
</Tasks></Project>"""


def test_calendar_inheritance_and_exceptions(tmp_path):
    f = tmp_path / "cal.xml"
    f.write_text(XML.format(bfe=""), encoding="utf-8")
    p = read_xml(f, Config())
    cal = p.calendar
    assert cal.name == "Turno sábado"
    assert cal.windows(date(2026, 1, 10)) == [(480, 780)]          # sábado (derivado)
    assert cal.windows(date(2026, 1, 9)) == [(480, 720), (780, 1020)]   # viernes (base)
    assert not cal.is_working_day(date(2026, 1, 11))                # domingo
    assert not cal.is_working_day(date(2026, 12, 25))               # anual por fecha
    assert not cal.is_working_day(date(2028, 12, 25))
    assert not cal.is_working_day(date(2027, 2, 16))                # diaria
    # Tarea que cruza de mes: sábado 31/01 (5 h) + lunes 02/02 (8 h)
    assert cal.minutes_by_month(datetime(2026, 1, 31, 8), datetime(2026, 2, 2, 17)) == {(2026, 1): 300,
                                                                                         (2026, 2): 480}


def test_baseline_default_from_project_setting(tmp_path):
    f = tmp_path / "bl.xml"
    f.write_text(XML.format(bfe="<BaselineForEarnedValue>3</BaselineForEarnedValue>"), encoding="utf-8")
    p = read_xml(f, Config())
    assert p.baseline_label == "Línea base 3"
    assert p.tasks[0].baseline_start == datetime(2026, 2, 2, 8)
    p0 = read_xml(f, Config(baseline=0))
    assert p0.baseline_label == "Línea base" and p0.tasks[0].baseline_start == datetime(2026, 1, 5, 8)
