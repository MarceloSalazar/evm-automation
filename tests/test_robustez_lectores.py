"""Robustez de los lectores: XML de MS Project, calendario laboral, campos y COM (objetos simulados).

Cubren casos límite de archivos reales (campos de recurso con el mismo alias,
tareas sin nivel de esquema, excepciones de calendario no interpretadas,
fondos de recursos abiertos con el archivo, etc.) sin requerir MS Project.
"""
import sys
import types
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from pathlib import Path

import pytest

from evm.config import Config
from evm.fields import TASK_NUMBER_IDS, TASK_TEXT_IDS, builtin_kind, fmt_number, parse_custom_name, parse_number
import evm.reader_com as reader_com
from evm.reader_com import PJ_DO_NOT_SAVE, ComSession, _com_date, _dt, _f
from evm.reader_xml import XmlFormatError, read_xml
from evm.wbs import build_report
from mspdi import (NUMBER1, RESOURCE_FIELD_IDS, TEXT3, TEXT10, calendar_xml, exception_xml, project_xml,
                   standard_week, task_xml, weekday_xml, work_week_xml)
from test_reader_com_fake import FakeApp, FakeProject, FakeTask, make_session, program

UTC = timezone.utc
S, F = datetime(2026, 1, 5, 8), datetime(2026, 1, 9, 17)
STANDARD = [(480, 720), (780, 1020)]


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
@pytest.fixture()
def read(tmp_path):
    """Escribe el XML en un archivo temporal y lo lee con read_xml."""
    def run(xml: str, cfg: Config | None = None):
        path = tmp_path / "proyecto.xml"
        path.write_text(xml, encoding="utf-8")
        return read_xml(path, cfg or Config())
    return run


def simple_tasks(**kw) -> list[str]:
    """Tarea resumen del proyecto (UID 0) y una tarea con línea base 0."""
    return [task_xml(0, 0, "Proyecto"),
            task_xml(1, 1, "A", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)], **kw)]


def calendar_of(read, *calendars, uid=1, cfg=None):
    return read(project_xml(simple_tasks(), calendars=calendars, calendar_uid=uid), cfg).calendar


def weekdays_between(a: date, b: date):
    d = a
    while d <= b:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


# ----------------------------------------------------------------------------
# Lector XML: campos personalizados
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("definitions, cfg, attr, expected", [
    pytest.param([(TEXT3, "Text3", "Empresa"), (RESOURCE_FIELD_IDS[0], "Text1", "Empresa")],
                 Config(multi_enabled=True), "company", "EmpresaA", id="alias de tarea antes que el de recurso"),
    pytest.param([(RESOURCE_FIELD_IDS[0], "Text1", "Empresa"), (TEXT3, "Text3", "Empresa")],
                 Config(multi_enabled=True), "company", "EmpresaA", id="alias de recurso antes que el de tarea"),
    pytest.param([(RESOURCE_FIELD_IDS[1], "Text10", "Departamento"), (TEXT10, "Text10", "")],
                 Config(method_field="Text10"), "method", "50/50", id="nombre Text10 de recurso antes que el de tarea"),
])
def test_custom_field_resolution_uses_task_fields(read, definitions, cfg, attr, expected):
    """Las definiciones de campos de recurso no ocultan los campos de tarea."""
    tasks = simple_tasks(ext={TEXT3: "EmpresaA", TEXT10: "50/50"})
    p = read(project_xml(tasks, definitions=definitions), cfg)
    assert getattr(p.tasks[0], attr) == expected


def test_alias_only_in_resource_field_is_reported_as_missing(read):
    tasks = simple_tasks(ext={TEXT3: "EmpresaA"})
    p = read(project_xml(tasks, definitions=[(RESOURCE_FIELD_IDS[0], "Text1", "Empresa")]), Config(multi_enabled=True))
    assert any("No se encontró el campo de empresa" in w.message for w in p.warnings)


def test_alias_matching_ignores_case_accents_and_spaces(read):
    tasks = simple_tasks(ext={TEXT10: "50/50", NUMBER1: "30.5"})
    defs = [(TEXT10, "Text10", "Método EVM"), (NUMBER1, "Number1", "Peso hito")]
    p = read(project_xml(tasks, definitions=defs), Config(method_field="  metodo evm ", milestone_weight_field="PESO HITO"))
    assert (p.tasks[0].method, p.tasks[0].milestone_weight, p.warnings) == ("50/50", 30.5, [])
    # Sin <ExtendedAttributes>: el nombre del campo (en inglés o en español) basta
    p2 = read(project_xml(tasks), Config(method_field="Texto 10", milestone_weight_field="Número1"))
    assert (p2.tasks[0].method, p2.tasks[0].milestone_weight, p2.warnings) == ("50/50", 30.5, [])


def test_missing_fields_produce_warnings(read):
    tasks = [task_xml(0, 0, "Proyecto"), task_xml(1, 1, "A", wbs="1.1", start=S, finish=F, cost=100,
                                                  baselines=[(0, S, F, 100)])]
    cfg = Config(wbs_field="Código cliente", method_field="Método EVM", milestone_weight_field="Peso",
                 multi_enabled=True)
    p = read(project_xml(tasks), cfg)
    assert p.tasks[0].wbs == "1.1"                          # se usa el campo EDT (WBS) de MS Project
    for text in ("campo EDT 'Código cliente'", "campo de método 'Método EVM'", "peso de hitos 'Peso'",
                 "campo de empresa 'Empresa'", "campo de reparto '% Reparto'"):
        assert any(text in w.message for w in p.warnings), text


# ----------------------------------------------------------------------------
# Lector XML: línea base y estructura
# ----------------------------------------------------------------------------
def test_task_duration_is_split_with_project_calendar(read):
    s, f = datetime(2026, 1, 26, 8), datetime(2026, 2, 6, 17)
    tasks = [task_xml(0, 0, "Proyecto"), task_xml(1, 1, "A", wbs="1", start=s, finish=f, baselines=[(0, s, f, 0)])]
    feriado = exception_xml(date(2026, 2, 2), date(2026, 2, 2))
    xml = project_xml(tasks, calendars=[calendar_xml(1, exceptions=[feriado])], calendar_uid=1)
    # Enero: 26 al 30 (5 días); febrero: 3 al 6 (4 días, el 2 es feriado del calendario)
    assert build_report(read(xml), Config()).rows[0].pct[""] == pytest.approx({(2026, 1): 5 / 9, (2026, 2): 4 / 9})
    # Los feriados de la configuración se suman a los del calendario del proyecto
    cfg = Config(holidays={date(2026, 1, 30)})
    assert build_report(read(xml, cfg), cfg).rows[0].pct[""] == pytest.approx({(2026, 1): 0.5, (2026, 2): 0.5})


def test_baseline_n_reads_its_dates(read):
    s, f = datetime(2026, 1, 26, 8), datetime(2026, 2, 6, 17)
    tasks = [task_xml(0, 0, "Proyecto"),
             task_xml(1, 1, "A", wbs="1", start=s, finish=f,
                      baselines=[(0, s, f, 0), (2, s, datetime(2026, 4, 10, 17), 0)])]
    p = read(project_xml(tasks, header="<BaselineForEarnedValue>0</BaselineForEarnedValue>"), Config(baseline=2))
    assert p.baseline_label == "Línea base 2" and p.tasks[0].baseline_finish == datetime(2026, 4, 10, 17)
    # Sin indicar la línea base: la que el proyecto usa para el valor acumulado
    p = read(project_xml(tasks, header="<BaselineForEarnedValue>2</BaselineForEarnedValue>"))
    assert p.baseline_label == "Línea base 2"
    only_bl1 = [task_xml(0, 0, "Proyecto"), task_xml(1, 1, "A", wbs="1", start=S, finish=F, baselines=[(1, S, F, 0)])]
    p1 = read(project_xml(only_bl1, header="<BaselineForEarnedValue>0</BaselineForEarnedValue>"), Config(baseline=1))
    assert p1.baseline_found and p1.tasks[0].baseline_start == S


@pytest.mark.parametrize("value", ["11", "-1", "abc", ""])
def test_invalid_ev_baseline_setting_defaults_to_baseline_0(read, value):
    p = read(project_xml(simple_tasks(), header=f"<BaselineForEarnedValue>{value}</BaselineForEarnedValue>"))
    assert (p.baseline_label, p.tasks[0].baseline_start) == ("Línea base", S)


def test_requested_baseline_missing_uses_current_schedule(read):
    s, f = datetime(2026, 2, 2, 8), datetime(2026, 2, 6, 17)
    tasks = [task_xml(0, 0, "Proyecto", baselines=[(0, S, F, 0)]),
             task_xml(1, 1, "A", wbs="1", start=s, finish=f, baselines=[(0, S, F, 0)])]
    p = read(project_xml(tasks), Config(baseline=3))
    t = p.tasks[0]
    assert not p.baseline_found and p.baseline_label == "Programa actual (sin línea base 3)"
    assert (t.baseline_start, t.baseline_finish) == (s, f)


@pytest.mark.parametrize("text, message", [
    ("<root/>", "falta el elemento <Project>"),
    ("<Project><Tasks>", "no es válido"),
    (project_xml(None), "no contiene tareas"),
], ids=["sin Project", "XML truncado", "sin Tasks"])
def test_invalid_xml_raises_xml_format_error(tmp_path, text, message):
    path = tmp_path / "malo.xml"
    path.write_text(text, encoding="utf-8")
    with pytest.raises(XmlFormatError, match=message):
        read_xml(path, Config())


def test_external_tasks_title_fallback_order_and_flags(read):
    tasks = [
        task_xml(3, 1, "C", wbs="3", start=S, finish=F, baselines=[(0, S, F, 0)]),
        task_xml(1, 1, "A", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)],
                 extra="<Milestone>true</Milestone><Active>0</Active>"),
        task_xml(2, 1, "Externa", wbs="2", start=S, finish=F, baselines=[(0, S, F, 0)],
                 extra="<ExternalTask>1</ExternalTask>"),
        task_xml(4, 1, "D", wbs="4", start=S, finish=F, baselines=[(0, S, F, 0)]),
        task_xml(5, 1, "", extra="<IsNull>1</IsNull>"),
    ]
    p = read(project_xml(tasks, title=None, header="<Name>archivo.xml</Name>"))
    assert [t.name for t in p.tasks] == ["A", "C", "D"]          # orden por ID; sin externas ni filas en blanco
    assert p.name == "archivo"
    assert any("extern" in w.message for w in p.warnings)
    assert [(t.milestone, t.active) for t in p.tasks] == [(True, False), (False, True), (False, True)]


def test_dates_before_1984_are_empty(read):
    tasks = [task_xml(0, 0, "Proyecto"),
             task_xml(1, 1, "A", wbs="1", start="1983-12-31T08:00:00", finish="2026-01-09T17:00:00Z",
                      baselines=[(0, date(2026, 1, 5), F, 0)])]
    t = read(project_xml(tasks)).tasks[0]
    assert t.start is None and t.finish == F                    # sin zona horaria
    assert t.baseline_start == datetime(2026, 1, 5)             # fecha sin hora


def test_task_without_outline_level_is_not_project_summary(read):
    """Solo la tarea de UID 0 es la tarea resumen del proyecto; sin nivel de esquema (o con 0) se usa el 1."""
    tasks = [task_xml(0, 0, "Proyecto", baselines=[(0, S, F, 0)]),
             task_xml(1, None, "Sin nivel", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)]),
             task_xml(2, 0, "Nivel cero", wbs="2", start=S, finish=F, baselines=[(0, S, F, 0)]),
             task_xml(3, 2, "B", wbs="2.1", start=S, finish=F, baselines=[(0, S, F, 0)])]
    p = read(project_xml(tasks))
    assert p.name == "Proyecto"
    assert [(t.name, t.outline_level) for t in p.tasks] == [("Sin nivel", 1), ("Nivel cero", 1), ("B", 2)]


# ----------------------------------------------------------------------------
# Calendario laboral (<Calendars> del XML)
# ----------------------------------------------------------------------------
def test_weekly_and_monthly_recurring_exceptions(read):
    weekly = exception_xml(date(2026, 3, 2), date(2026, 3, 31), kind=6, Period=2, DaysOfWeek=10)  # lunes y miércoles
    monthly = exception_xml(date(2026, 1, 1), date(2026, 6, 30), kind=4, MonthDay=31)            # día 31 o el último
    cal = calendar_of(read, calendar_xml(1, exceptions=[weekly, monthly]))
    off = [d for d in weekdays_between(date(2026, 3, 1), date(2026, 4, 30)) if not cal.is_working_day(d)]
    assert off == [date(2026, 3, 2), date(2026, 3, 4), date(2026, 3, 16), date(2026, 3, 18), date(2026, 3, 30),
                   date(2026, 3, 31), date(2026, 4, 30)]


def test_yearly_exception_on_february_29(read):
    yearly = exception_xml(date(2027, 1, 1), date(2029, 12, 31), kind=2, Month=1, MonthDay=29)   # Month en base 0
    cal = calendar_of(read, calendar_xml(1, exceptions=[yearly]))
    assert not cal.is_working_day(date(2028, 2, 29))
    assert not cal.is_working_day(date(2029, 2, 28))            # año no bisiesto: último día de febrero
    assert cal.is_working_day(date(2028, 2, 28))


def test_legacy_exceptions_work_weeks_and_working_exception(read):
    legacy = weekday_xml(0, working=False, period=(date(2026, 7, 6), date(2026, 7, 8)))
    four_by_ten = work_week_xml(date(2026, 2, 1), date(2026, 2, 28),
                                [weekday_xml(d, hours=(("07:00:00", "17:00:00"),)) for d in (2, 3, 4, 5)]
                                + [weekday_xml(6, working=False)])
    saturday = exception_xml(date(2026, 1, 31), date(2026, 1, 31), working=True, hours=(("09:00:00", "11:30:00"),))
    cal = calendar_of(read, calendar_xml(1, weekdays=standard_week() + [legacy], work_weeks=[four_by_ten],
                                         exceptions=[saturday]))
    assert [cal.is_working_day(date(2026, 7, d)) for d in (6, 7, 8, 9)] == [False, False, False, True]
    assert cal.windows(date(2026, 2, 2)) == [(420, 1020)] and cal.windows(date(2026, 2, 6)) == []
    assert cal.windows(date(2026, 3, 6)) == STANDARD
    assert cal.windows(date(2026, 1, 31)) == [(540, 690)]
    # Sábado 31/01 de 09:00 a 11:30 (150 min) y lunes 02/02 de 07:00 a 17:00 (600 min)
    assert cal.minutes_by_month(datetime(2026, 1, 31, 8), datetime(2026, 2, 2, 17)) == {(2026, 1): 150,
                                                                                         (2026, 2): 600}


@pytest.mark.parametrize("calendars, uid", [
    pytest.param([calendar_xml(1)], None, id="sin CalendarUID"),
    pytest.param([calendar_xml(1)], 99, id="calendario inexistente"),
    pytest.param([calendar_xml(1, base=2), calendar_xml(2, name="Otro", base=1)], 1, id="herencia circular"),
    pytest.param([calendar_xml(1, weekdays=[weekday_xml(d, working=False) for d in range(1, 8)])], 1,
                 id="sin días laborables"),
])
def test_calendar_fallbacks_use_standard_week(read, calendars, uid):
    cal = calendar_of(read, *calendars, uid=uid)
    assert cal.windows(date(2026, 1, 5)) == STANDARD and not cal.is_working_day(date(2026, 1, 4))


@pytest.mark.parametrize("kind, fields", [
    (3, {"Month": 4, "MonthItem": 4, "MonthPosition": 4}),      # anual: último lunes de mayo
    (5, {"Period": 1, "MonthItem": 4, "MonthPosition": 4}),     # mensual: último lunes de cada mes
    (8, {}),
], ids=["anual relativa", "mensual relativa", "tipo 8"])
def test_relative_exceptions_are_applied_or_reported(read, kind, fields):
    exc = exception_xml(date(2026, 5, 25), date(2026, 5, 25), kind=kind, Occurrences=1, **fields)
    p = read(project_xml(simple_tasks(), calendars=[calendar_xml(1, exceptions=[exc])], calendar_uid=1))
    assert not p.calendar.is_working_day(date(2026, 5, 25)) or any("excepci" in w.message.lower() for w in p.warnings)


def test_unreadable_exception_warning_names_calendar_and_exception(read):
    exceptions = [exception_xml(date(2026, 5, 25), date(2026, 5, 25), kind=99, name="Feriado móvil"),
                  exception_xml(date(2026, 1, 1), date(2027, 12, 31), kind=2, name="Sin mes", MonthDay=1),
                  exception_xml(date(2026, 2, 2), date(2026, 2, 2), name="Feriado fijo")]
    p = read(project_xml(simple_tasks(), calendars=[calendar_xml(1, "Turno", exceptions=exceptions)], calendar_uid=1))
    omitted = [w.message for w in p.warnings if "se omite la excepción" in w.message]
    assert len(omitted) == 2
    assert omitted[0].startswith("Calendario 'Turno': se omite la excepción 'Feriado móvil' (tipo 99)")
    assert "'Sin mes' (tipo 2)" in omitted[1] and "reparto de la duración de las tareas" in omitted[1]
    assert not p.calendar.is_working_day(date(2026, 2, 2))      # las demás excepciones se aplican


# ----------------------------------------------------------------------------
# Campos y números
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("text, expected", [
    ("1.234,56", 1234.56), ("1,234.56", 1234.56), ("12,5", 12.5), ("12.5", 12.5), ("60 %", 60.0),
    ("\u00a060\u00a0", 60.0), ("1\u202f234,5", 1234.5), ("1,234,567", 1234567.0), ("-3,5", -3.5),
    ("", None), ("-", None), ("abc", None), (None, None), (30, 30.0), (0.0, 0.0),
])
def test_parse_number_locale_formats(text, expected):
    assert parse_number(text) == expected


@pytest.mark.parametrize("value, decimals, trim, expected", [
    (44210000, 0, False, "44.210.000"), (-1901492.4, 0, False, "-1.901.492"), (0.84, 2, False, "0,84"),
    (1234.5, 2, False, "1.234,50"), (27.44, 1, False, "27,4"), (-0.004, 2, False, "0,00"),
    (2.5, 2, True, "2,5"), (100.0, 2, True, "100"), (1000.0, 1, True, "1.000"), (1000.0, 0, True, "1.000"),
])
def test_fmt_number_spanish_separators(value, decimals, trim, expected):
    # Punto de miles y coma decimal, como Excel en español; sin '-0,00'
    assert fmt_number(value, decimals, trim) == expected


def test_field_id_tables_and_custom_names():
    # Constantes PjField de MS Project (pjTaskTextN / pjTaskNumberN)
    assert [TASK_TEXT_IDS[i] for i in (1, 10, 11, 30)] == [188743731, 188743750, 188743997, 188744016]
    assert [TASK_NUMBER_IDS[i] for i in (1, 5, 6, 20)] == [188743767, 188743771, 188743982, 188743996]
    ids = list(TASK_TEXT_IDS.values()) + list(TASK_NUMBER_IDS.values())
    assert len(ids) == len(set(ids)) == 50
    assert parse_custom_name("Texto 10") == ("Text", 10) and parse_custom_name("Número20") == ("Number", 20)
    assert parse_custom_name("TEXTO30") == ("Text", 30)
    assert [parse_custom_name(n) for n in ("Text0", "Text31", "Number21", "Costo1")] == [None] * 4
    assert builtin_kind("Número de esquema") == builtin_kind("Outline Number") == "OutlineNumber"
    assert builtin_kind(" wbs ") == "WBS" and builtin_kind("EDT") is None


# ----------------------------------------------------------------------------
# Lector COM (objetos simulados: no requiere MS Project ni pywin32)
# ----------------------------------------------------------------------------
ORIGINAL_STATUS = datetime(2026, 5, 31, 17, 0, tzinfo=UTC)
FAKE_PYTHONCOM = types.SimpleNamespace(Empty=object(), CoUninitialize=lambda: None)


@pytest.fixture()
def fake_pywintypes(monkeypatch):
    """ComSession._open importa pywintypes: se reemplaza por un módulo simulado."""
    module = types.ModuleType("pywintypes")
    module.com_error = type("com_error", (Exception,), {})
    monkeypatch.setitem(sys.modules, "pywintypes", module)


class OpenProject(FakeProject):
    """Proyecto abierto en MS Project (con ruta, fecha de estado y activación)."""

    def __init__(self, full_name: str, tasks=None):
        super().__init__(program() if tasks is None else tasks)
        self.FullName = full_name
        self.StatusDate = ORIGINAL_STATUS
        self.activations = 0

    def Activate(self):
        self.activations += 1


class Projects:
    def __init__(self, items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]


class SessionApp(FakeApp):
    """MS Project simulado: registra las llamadas que abren, cierran o recalculan."""

    def __init__(self, open_projects=(), to_open=None, pool=None, open_ok=True):
        self.Projects = Projects(open_projects)
        self.ActiveProject = self.Projects.items[0] if self.Projects.items else None
        self.DisplayAlerts = True
        self.AutomationSecurity = 1
        self.calls = []
        self.to_open, self.pool, self.open_ok = to_open, pool, open_ok

    def FileOpenEx(self, *args):
        self.calls.append(("FileOpenEx", len(args)))
        if not self.open_ok:
            return False
        self.Projects.items.append(self.to_open)
        if self.pool is not None:                 # fondo de recursos abierto junto con el archivo
            self.Projects.items.append(self.pool)
        self.ActiveProject = self.to_open
        return True

    def FileCloseEx(self, save):
        self.calls.append(("FileCloseEx", save))
        self.Projects.items.remove(self.ActiveProject)
        self.ActiveProject = self.Projects.items[0] if self.Projects.items else None

    def Quit(self, save):
        self.calls.append(("Quit", save))

    def CalculateProject(self):
        self.calls.append(("CalculateProject",))


def com_session(app: SessionApp, origen: str, started_by_us: bool = False) -> ComSession:
    s = ComSession(origen)
    s.app, s.started_by_us, s._pythoncom = app, started_by_us, FAKE_PYTHONCOM
    return s


def test_com_cleanup_restores_status_date_and_keeps_user_project_open(tmp_path, fake_pywintypes):
    origen = str(tmp_path / "programa.mpp")
    other = OpenProject(r"C:\otro.mpp", tasks=[])
    user_p = OpenProject(str(Path(origen).resolve()))
    app = SessionApp([other, user_p])
    s = com_session(app, origen)
    s._open()
    assert s.project is user_p and not s.opened_by_us
    s.read_project(Config(), status_date=date(2026, 9, 30))
    assert user_p.StatusDate == datetime(2026, 9, 30, 23, 59, tzinfo=UTC) and ("CalculateProject",) in app.calls
    s._cleanup()
    assert user_p.StatusDate == ORIGINAL_STATUS                 # se restauró la fecha de estado del usuario
    assert not [c for c in app.calls if c[0] in ("FileCloseEx", "Quit")]
    assert other.activations == 1 and app.DisplayAlerts is True


@pytest.mark.parametrize("started_by_us, with_pool, expect_quit", [
    pytest.param(False, False, False, id="instancia del usuario"),
    pytest.param(False, True, False, id="instancia del usuario con fondo de recursos"),
    pytest.param(True, False, True, id="instancia propia"),
    pytest.param(True, True, True, id="instancia propia con fondo de recursos"),
])
def test_com_session_closes_file_and_quits_only_its_own_instance(tmp_path, fake_pywintypes, started_by_us,
                                                                 with_pool, expect_quit):
    origen = str(tmp_path / "programa.mpp")
    other = None if started_by_us else OpenProject(r"C:\otro.mpp", tasks=[])
    app = SessionApp([] if other is None else [other], to_open=OpenProject(str(Path(origen).resolve())),
                     pool=OpenProject(r"C:\fondo.mpp", tasks=[]) if with_pool else None)
    s = com_session(app, origen, started_by_us)
    s._open()
    assert s.opened_by_us and app.DisplayAlerts is False and app.AutomationSecurity == 3
    s._cleanup()
    assert ("FileCloseEx", PJ_DO_NOT_SAVE) in app.calls      # se cierra sin guardar
    if other is not None:                                     # instancia del usuario: se restaura su estado
        assert app.DisplayAlerts is True and app.AutomationSecurity == 1 and other.activations == 1
    assert (("Quit", PJ_DO_NOT_SAVE) in app.calls) is expect_quit


def test_com_started_instance_is_kept_if_user_opened_other_projects(tmp_path, fake_pywintypes):
    origen = str(tmp_path / "programa.mpp")
    app = SessionApp([], to_open=OpenProject(str(Path(origen).resolve())))
    s = com_session(app, origen, started_by_us=True)
    s._open()
    app.Projects.items.append(OpenProject(r"C:\del_usuario.mpp", tasks=[]))   # abierto por el usuario mientras tanto
    s._cleanup()
    assert ("FileCloseEx", PJ_DO_NOT_SAVE) in app.calls
    assert not [c for c in app.calls if c[0] == "Quit"]      # no se descarta el trabajo del usuario


def test_com_open_failure_restores_app_properties(tmp_path, fake_pywintypes):
    origen = str(tmp_path / "programa.mpp")
    other = OpenProject(r"C:\otro.mpp", tasks=[])
    app = SessionApp([other], to_open=OpenProject(origen), open_ok=False)
    s = com_session(app, origen)
    with pytest.raises(RuntimeError, match="no pudo abrir"):
        s._open()
    s._cleanup()                                              # __enter__ lo llama ante cualquier error
    assert (app.DisplayAlerts, app.AutomationSecurity) == (True, 1)
    assert not [c for c in app.calls if c[0] in ("FileCloseEx", "Quit")] and other.activations == 1


RPC_SERVER_UNAVAILABLE = -2147023174      # 0x800706BA con signo, como lo entrega pywin32


class ComError(Exception):
    """pywintypes.com_error simulado (con su HRESULT)."""

    def __init__(self, hresult: int):
        super().__init__(hresult, "error COM", None, None)
        self.hresult = hresult


class ConnectApp:
    """Instancia de MS Project para _connect: puede rechazar las primeras llamadas (cerrándose u ocupada)."""

    def __init__(self, count: int = 0, user_control: bool = False, failures: int = 0):
        self._oleobj_ = self
        self.UserControl = user_control
        self.Visible = None
        self.count, self.failures = count, failures

    @property
    def Projects(self):
        if self.failures:
            self.failures -= 1
            raise ComError(RPC_SERVER_UNAVAILABLE)
        return Projects([object()] * self.count)


@pytest.fixture()
def fake_win32com(monkeypatch):
    """win32com y pywintypes simulados: GetActiveObject y DispatchEx se definen en cada prueba."""
    pywintypes = types.ModuleType("pywintypes")
    pywintypes.com_error = ComError
    client = types.ModuleType("win32com.client")
    client.gencache = types.SimpleNamespace(EnsureDispatch=lambda obj: obj)
    client.dynamic = types.SimpleNamespace(Dispatch=lambda obj: obj)
    package = types.ModuleType("win32com")
    package.client = client
    for name, module in (("pywintypes", pywintypes), ("win32com", package), ("win32com.client", client)):
        monkeypatch.setitem(sys.modules, name, module)
    monkeypatch.setattr(reader_com.time, "sleep", lambda seconds: None)
    client.calls = []

    def sequence(name, results):
        def call(progid):
            client.calls.append(name)
            item = results.pop(0)
            if isinstance(item, Exception):
                raise item
            return item
        setattr(client, name, call)

    client.sequence = sequence
    return client


def test_com_connect_waits_for_a_closing_instance(fake_win32com):
    """Una instancia que se está cerrando (p. ej. la de una ejecución anterior) no se usa: se espera a que
    termine de cerrarse y se inicia otra."""
    closing, fresh = ConnectApp(failures=99), ConnectApp(count=0)
    fake_win32com.sequence("GetActiveObject", [closing, closing, ComError(-2147221021)])
    fake_win32com.sequence("DispatchEx", [fresh])
    s = ComSession("programa.mpp")
    s._connect()
    assert s.app is fresh and s.started_by_us
    assert fake_win32com.calls == ["GetActiveObject", "GetActiveObject", "GetActiveObject", "DispatchEx"]


def test_com_connect_uses_the_user_instance(fake_win32com):
    user = ConnectApp(count=1, user_control=True)
    fake_win32com.sequence("GetActiveObject", [user])
    s = ComSession("programa.mpp")
    s._connect()
    assert s.app is user and not s.started_by_us and fake_win32com.calls == ["GetActiveObject"]


def test_com_connect_waits_for_its_own_busy_instance_without_starting_more(fake_win32com):
    busy = ConnectApp(count=0, failures=2)
    fake_win32com.sequence("GetActiveObject", [ComError(-2147221021)])
    fake_win32com.sequence("DispatchEx", [busy])
    s = ComSession("programa.mpp")
    s._connect()
    assert s.app is busy and s.started_by_us
    assert fake_win32com.calls == ["GetActiveObject", "DispatchEx"]      # no se inician otras instancias


def test_com_connect_gives_up_with_a_clear_message(fake_win32com, monkeypatch):
    monkeypatch.setattr(reader_com, "CONNECT_WAIT", 0.0)
    fake_win32com.sequence("GetActiveObject", [ConnectApp(failures=99)])
    s = ComSession("programa.mpp")
    with pytest.raises(RuntimeError, match="no responde.*0x800706ba"):
        s._connect()


def test_com_cleanup_releases_objects_before_quitting(tmp_path, fake_pywintypes):
    """Al cerrar MS Project ya no quedan referencias al proyecto (liberarlas después produce errores de RPC
    0x800706BA)."""
    origen = str(tmp_path / "programa.mpp")
    app = SessionApp([], to_open=OpenProject(str(Path(origen).resolve())))
    s = com_session(app, origen, started_by_us=True)
    s._open()
    seen = {}
    app.Quit = lambda save: seen.update(project=s.project)
    s._cleanup()
    assert seen == {"project": None}
    assert s.app is None


def test_com_field_resolution_and_external_tasks():
    company_fid = 188744000                                   # constante devuelta por FieldNameToFieldConstant

    class App(FakeApp):
        def FieldNameToFieldConstant(self, name, kind):
            if name == "Empresa":
                return company_fid
            raise RuntimeError("no existe")

    class CompanyTask(FakeTask):
        def GetField(self, fid):
            return "EmpresaB" if fid == company_fid else ""

    tasks = program()
    tasks[3] = CompanyTask(3, 3, "Paquete 1", "1.1.1", {(2026, 1): 100.0, (2026, 2): 300.0}, method="50/50")
    external = FakeTask(9, 1, "Externa", "9", {(2026, 1): 10.0})
    external.ExternalTask = True
    tasks.append(external)
    s = make_session(tasks)
    s.app = App()
    cfg = Config(wbs_field="EDT", multi_enabled=True)
    p = s.read_project(cfg)
    t3 = next(t for t in p.tasks if t.uid == 3)
    assert (t3.wbs, t3.company) == ("1.1.1", "EmpresaB")      # 'EDT' sin alias: campo EDT (WBS)
    assert 9 not in [t.uid for t in p.tasks] and any("extern" in w.message for w in p.warnings)
    assert [r.code for r in build_report(p, cfg).rows if r.is_input] == ["1.1.1", "1.1.2"]


def test_com_empty_number_field():
    s = make_session(program())
    p = s.read_project(Config(milestone_weight_field="Peso hito"))
    t3 = next(t for t in p.tasks if t.uid == 3)
    # En MS Project un campo numérico vacío vale 0: se lee 0 (los métodos lo tratan como vacío)
    assert t3.milestone_weight == 0.0


def test_com_value_conversions():
    assert (_f("NA"), _f(Decimal("1234.56")), _f("1.234,5"), _f(None, None), _f("", 3)) == (0.0, 1234.56, 1234.5,
                                                                                           None, 3)
    assert _dt("NA") is None and _dt(datetime(1983, 12, 31)) is None and _dt(datetime(2150, 1, 1)) is None
    # Fecha con zona (pywin32): se conserva la hora del reloj, sin convertir
    assert _dt(datetime(2026, 1, 31, 23, 59, 59, 965000, tzinfo=UTC)) == datetime(2026, 1, 31, 23, 59, 59)
    assert _com_date(datetime(2026, 1, 5, 8, 0, 30)) == datetime(2026, 1, 5, 8, 0, tzinfo=UTC)
    assert _com_date(date(2026, 1, 5)) == datetime(2026, 1, 5, tzinfo=UTC)
