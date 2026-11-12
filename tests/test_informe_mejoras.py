"""validar como revisión de mejoras: informe de mejoras (--informe, sin nombres), evaluación DCMA, varias
planificaciones con su comparación y controles: fin previsto fuera del horizonte (CFG-052), líneas base incompletas
(FEC-015), tareas con calendario propio (FEC-017) y vínculos con otras planificaciones (LOG-018)."""
from datetime import date, datetime

import pytest

from evm.cli import main
from evm.config import Config
from evm.messages import SEV_ERROR, SEV_INFO, SEV_WARN, Issue
from evm.reader_xml import read_xml
from evm.review import PRIORITY_HIGH, PRIORITY_LOW, PRIORITY_MEDIUM, priority, sanitizer
from evm.validation import validate_part
from evm.version import version_label
from evm.wbs import build_report

from builders import P, T, by_name, tree
from mspdi import calendar_xml, project_xml, task_xml, write_xml

S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)

# Nombres inventados que el informe no debe mostrar
SECRET = ("Estación Secreta Norte", "Montaje Confidencial", "Programa Reservado", "Calendario Sede Privada")


def ids(issues, code):
    return [i for i in issues if i.id == code]


# ----------------------------------------------------------------------------
# Controles nuevos
# ----------------------------------------------------------------------------
def test_forecast_finish_after_the_horizon():
    """CFG-052: el fin previsto (cronograma actual) pasa del último mes del horizonte."""
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "A", finish=datetime(2026, 6, 30, 17))]
    [issue] = ids(build_report(P(tasks), Config(months=3, start_month=date(2026, 1, 1))).issues, "CFG-052")
    assert issue.severity == SEV_WARN
    assert issue.message.startswith("El fin previsto del cronograma (06/2026) es posterior al último mes del horizonte "
                                    "(03/2026)")
    assert "Se necesitan al menos 6 meses desde 01/2026 (hoy son 3)" in issue.message
    assert not ids(build_report(P(tasks), Config(months=6, start_month=date(2026, 1, 1))).issues, "CFG-052")


def test_incomplete_baseline_of_a_detail_task():
    """FEC-015 (generar y validar): la tarea de detalle no pesa en el % planificado de su fila."""
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "A"), T(3, 2, "1.2", "Solo comienzo", bf=None)]
    rep = build_report(P(tasks), Config())
    [issue] = ids(rep.issues, "FEC-015")
    assert (issue.code, issue.severity) == ("1.2.0", SEV_WARN)
    assert issue.message == ("La línea base solo tiene el comienzo (05/01/2026): se toma como si durara 0 y no pesa en "
                             "el % planificado de su fila.")


def test_incomplete_baselines_summary_in_validar(tmp_path):
    """FEC-015 (resumen de validar): tareas resumen y de detalle con solo una fecha en la línea base usada y, con XML,
    en las demás líneas base."""
    path = write_xml(tmp_path / "p.xml", project_xml([
        task_xml(1, 1, "Fase", wbs="1", start=S, finish=F, baselines=[(0, S, None, 0)],
                 extra="<Summary>1</Summary>"),
        task_xml(2, 2, "A", wbs="1.1", start=S, finish=F, baselines=[(0, S, F, 0), (1, None, F, 0)]),
        task_xml(3, 2, "B", wbs="1.2", start=S, finish=F, baselines=[(0, S, F, 0), (2, S, F, 0)]),
    ]))
    p = read_xml(path, Config())
    assert [t.incomplete_baselines for t in p.tasks] == [(0,), (1,), ()]
    [issue] = [i for i in ids(validate_part(p, Config()).checks, "FEC-015") if not i.code]
    assert issue.message.startswith("2 tareas tienen una línea base incompleta, con solo el comienzo o solo el fin "
                                    "(Línea base: 1 tarea y Línea base 1: 1 tarea): 1.0.0, 1.1.0.")


def test_tasks_with_their_own_calendar(tmp_path):
    """FEC-017 (validar): con XML, el calendario de la tarea distinto del del proyecto."""
    path = write_xml(tmp_path / "p.xml", project_xml([
        task_xml(1, 1, "Fase", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)]),
        task_xml(2, 2, "A", wbs="1.1", start=S, finish=F, baselines=[(0, S, F, 0)], extra="<CalendarUID>2</CalendarUID>"),
        task_xml(3, 2, "B", wbs="1.2", start=S, finish=F, baselines=[(0, S, F, 0)], extra="<CalendarUID>1</CalendarUID>"),
    ], calendars=[calendar_xml(1, "Estándar"), calendar_xml(2, SECRET[3])], calendar_uid=1))
    p = read_xml(path, Config())
    assert [t.calendar for t in p.tasks] == ["", SECRET[3], ""]
    [issue] = ids(validate_part(p, Config()).checks, "FEC-017")
    assert issue.message == ("1 tarea tiene un calendario de tarea distinto del calendario del proyecto (1 calendario): "
                             "MS Project calcula sus fechas con él, pero el programa reparte su duración entre los meses "
                             "con el calendario del proyecto (1.1.0).")


def test_com_reads_task_calendars_and_incomplete_baseline_for_validar():
    from test_reader_com_fake import FakeTask, make_session

    class Calendar:
        def __init__(self, name):
            self.Name = name

    a, b, c = FakeTask(1, 1, "A", "1"), FakeTask(2, 1, "B", "2"), FakeTask(3, 1, "C", "3")
    a.CalendarObject, b.CalendarObject = Calendar("Sede"), Calendar("Estándar")       # b: el del proyecto
    c.BaselineFinish = "NA"
    session = make_session([a, b, c])
    session.project.Calendar = Calendar("Estándar")         # solo el nombre: el reparto usa el calendario fijo
    p = session.read_project(Config(), review=True)
    assert [t.calendar for t in p.tasks] == ["Sede", "", ""]
    assert [t.incomplete_baselines for t in p.tasks] == [(), (), (0,)]
    assert [t.calendar for t in make_session([a]).read_project(Config()).tasks] == [""]   # solo con review


# ----------------------------------------------------------------------------
# Evaluación DCMA
# ----------------------------------------------------------------------------
def dcma_tasks():
    from evm.model import LINK_FS, LINK_SS, Link
    tasks = tree('''
        1 Fase
          1.1 A
          1.2 B
          1.3 C
    ''')
    t = by_name(tasks)
    t["B"].predecessors = [Link(t["A"].uid, LINK_FS)]
    t["C"].predecessors = [Link(t["B"].uid, LINK_SS, 480)]
    t["A"].start, t["A"].finish = S, datetime(2026, 1, 30, 17)
    t["C"].start, t["C"].finish = datetime(2026, 2, 2, 8), F
    return tasks


def test_dcma_metrics_of_validar():
    part = validate_part(P(dcma_tasks(), status_date=datetime(2026, 1, 31, 17)), Config())
    assert sorted(part.dcma) == list(range(1, 15))
    assert part.dcma[1] == ("0 % (0 sin predecesoras, 0 sin sucesoras)", "≤ 5 %", True)
    assert part.dcma[3] == ("1 (50 %)", "≤ 5 %", False)
    assert part.dcma[4] == ("50 %", "≥ 90 %", False)
    assert part.dcma[9][2] is False                     # A debía terminar el 30/01 y no tiene avance
    assert part.dcma[10] == ("no aplica (BAC en el libro)", "—", None)
    assert part.dcma[12][0] == "prueba manual"
    part = validate_part(P(dcma_tasks()), Config())     # sin fecha de estado
    assert part.dcma[9] == ("sin fecha de estado", "0", None)


# ----------------------------------------------------------------------------
# Prioridad y nombres
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("issue, expected", [
    (Issue("EDT-005", "1.2.0", "X", "m", SEV_ERROR), PRIORITY_HIGH),
    (Issue("CFG-052", message="m"), PRIORITY_HIGH),            # advertencia que también da generar
    (Issue("FEC-013", message="m"), PRIORITY_MEDIUM),          # advertencia que solo da validar
    (Issue("LOG-018", message="m"), PRIORITY_MEDIUM),
    (Issue("LEC-024", message="m"), PRIORITY_LOW),
])
def test_priority_of_each_observation(issue, expected):
    assert priority(issue) == expected


def test_sanitizer_removes_task_project_calendar_and_company_names():
    tasks = [T(1, 1, "1", SECRET[0], calendar=SECRET[3]), T(2, 2, "1.1", SECRET[1]), T(3, 1, "", "INTERNO")]
    project = P(tasks)
    project.name = SECRET[2]
    clean = sanitizer([project], Config(multi_enabled=True, company_a="Alfa SA"), keep={"INTERNO"})
    text = (f"{SECRET[0]} y «{SECRET[1]}» en {SECRET[2]} con {SECRET[3]}: Alfa SA participa; INTERNO; "
            f"{SECRET[1]}s no es una palabra completa.")
    assert clean(text) == ("… y «…» en … con …: empresa A participa; INTERNO; Montaje Confidencials no es una "
                           "palabra completa.")


# ----------------------------------------------------------------------------
# Comando validar
# ----------------------------------------------------------------------------
def two_plans_xml(tmp_path):
    """Dos planificaciones con los mismos códigos, una vinculada a la otra, y nombres que el informe no muestra."""
    late = datetime(2026, 4, 30, 17)

    def task(uid, level, name, wbs=None, **kw):
        return task_xml(uid, level, name, wbs=wbs, start=S, finish=kw.pop("finish", F),
                        baselines=[(0, S, kw.pop("bf", F), 0)], **kw)

    return write_xml(tmp_path / "plan.xml", project_xml([
        task_xml(0, 0, SECRET[2]),
        task(1, 1, "CLIENTE"),
        task(2, 2, SECRET[0], "1"),
        task(3, 3, SECRET[1], "1.1", extra="<PredecessorLink><PredecessorUID>12</PredecessorUID><Type>3</Type>"
                                           "</PredecessorLink>"),
        task(4, 3, "Solo en cliente", "1.2"),
        task(10, 1, "INTERNO"),
        task(11, 2, SECRET[0], "1"),
        task(12, 3, SECRET[1], "1.1", finish=late, bf=late),
    ], title=SECRET[2], header="<StatusDate>2026-03-31T17:00:00</StatusDate>"))


def test_validar_report_has_no_names(tmp_path, capsys):
    path = two_plans_xml(tmp_path)
    assert main(["validar", "--origen", str(path), "--planificacion", "INTERNO", "--informe"]) in (0, 1)
    out = capsys.readouterr().out
    report = tmp_path / "plan_mejoras.md"
    assert f"Informe de mejoras: {report.resolve()}" in out and "Evaluación DCMA de 14 puntos:" in out
    text = report.read_text(encoding="utf-8")
    for heading in ("# Informe de mejoras del cronograma", "## 1. Resumen", "## 2. Mejoras por prioridad",
                    "### Prioridad alta", "### Prioridad media", "## 3. Evaluación DCMA de 14 puntos"):
        assert heading + "\n" in text
    assert "#### LOG-018 · Vínculos con otras planificaciones" in text and "**Qué hacer:**" in text
    assert "| 1. Lógica (sin predecesoras o sucesoras) |" in text
    for name in SECRET:
        assert name not in text, name
    assert f"| **Versión** | Exportador EVM {version_label()} · `validar --informe` |" in text
    assert f"| **Fecha** | {date.today():%d/%m/%Y} " in text
    # Las informativas de la estructura no son mejoras: sus datos están en el resumen
    assert "| Filas «Otras tareas» (-D) |" in text and "| Tareas sumadas en la fila de su elemento superior |" in text
    for code in ("EDT-016", "EDT-017", "EDT-018", "EDT-022", "VIS-007", "LEC-024"):
        assert f"#### {code} ·" not in text, code
    assert main(["validar", "--origen", str(path), "--informe"]) == 2          # ya existe: CFG-042
    assert "CFG-042" in capsys.readouterr().err
    assert main(["validar", "--origen", str(path), "--planificacion", "INTERNO", "--informe", "--sobrescribir"]) \
        in (0, 1)


def test_validar_reviews_several_plans_and_compares_them(tmp_path, capsys):
    path = two_plans_xml(tmp_path)
    report = tmp_path / "ambas.md"
    main(["validar", "--origen", str(path), "--planificacion", "INTERNO,CLIENTE", "--informe", str(report)])
    out = capsys.readouterr().out
    assert out.count("Validación: ") == 2
    assert "Planificación: INTERNO" in out and "Planificación: CLIENTE" in out
    assert "Comparación de planificaciones:" in out
    assert "Filas solo en CLIENTE: 1 (1.2.0)" in out
    assert "Filas comunes a INTERNO y CLIENTE: 2; con otras fechas de plan (línea base): 2 (1.0.0, 1.1.0)" in out
    text = report.read_text(encoding="utf-8")
    assert "## 4. Comparación de planificaciones" in text and "### INTERNO" in text and "### CLIENTE" in text
    assert "#### Prioridad alta" in text
    assert "| Dato | INTERNO | CLIENTE |" in text
    for name in SECRET:
        assert name not in text, name


@pytest.mark.parametrize("command", [["generar", "--salida", "x.xlsx"], ["generar-fichas"]])
def test_several_plans_only_in_validar(tmp_path, capsys, command):
    path = two_plans_xml(tmp_path)
    cmd = [command[0], "--origen", str(path), "--planificacion", "INTERNO,CLIENTE"] + \
        [str(tmp_path / c) if c.endswith(".xlsx") else c for c in command[1:]]
    assert main(cmd) == 2
    assert "CFG-053" in capsys.readouterr().err


def test_report_groups_the_tasks_of_a_code_by_case_and_leaves_structure_codes_out():
    """Un código con varios casos (LOG-003) lista sus tareas por caso; las informativas de la estructura no son
    mejoras (sus datos están en el resumen)."""
    from evm.review import _improvements
    issues = [Issue("LOG-003", "1.1", "A", "Sin sucesoras.", variant="sin sucesoras"),
              Issue("LOG-003", "1.2", "B", "Sin predecesoras.", variant="sin predecesoras"),
              Issue("LOG-003", "1.3", "C", "Sin sucesoras.", variant="sin sucesoras"),
              Issue("FEC-006", "", "", "2 tareas tienen duración estimada."),
              Issue("FEC-006", "2.1", "D", "Duración estimada («?»)."),
              Issue("EDT-022", message="3 tareas se agregaron en su elemento EDT superior."),
              Issue("LEC-024", message="Líneas base guardadas.")]
    text = "\n".join(_improvements(issues, lambda x: x, "###"))
    assert "- **Dónde (sin sucesoras, 2 tareas):** `1.1`, `1.3`" in text
    assert "- **Dónde (sin predecesoras, 1 tarea):** `1.2`" in text
    assert "#### FEC-006 · Duración estimada" in text and "- **Dónde:** `2.1`" in text
    assert "- **Observación:** 2 tareas tienen duración estimada." in text
    assert "EDT-022" not in text and "LEC-024" not in text
