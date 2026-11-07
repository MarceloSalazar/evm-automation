"""Planificación para el cliente (generar-cliente): tareas que se muestran, resúmenes convertidos, vínculos nuevos e
informe de control."""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime

import pytest

from evm.cli import main
from evm.client_plan import (ORIGIN_BRIDGE, ORIGIN_SUMMARY, SKIP_CYCLE, SKIP_LINKED, SKIP_SUMMARY, build_client_plan,
                             render_control, rollup_percent, working_minutes)
from evm.config import Config, load_config
from evm.fields import TASK_TEXT_IDS
from evm.messages import EvmError
from evm.model import LINK_FF, LINK_FS, LINK_SF, LINK_SS, Link, Task
from evm.reader_xml import read_xml
from evm.workcalendar import WorkCalendar

from builders import P
from mspdi import project_xml, task_xml, write_xml

TEXT15 = TASK_TEXT_IDS[15]
DAY = 480.0                                    # minutos de un día laborable (calendario estándar)
CFG = Config(plan="INTERNO", client_field="Etiqueta Cliente")


def d(day: int, hour: int = 8) -> datetime:
    """Fecha de enero de 2026 (el 5 es lunes)."""
    return datetime(2026, 1, day, hour)


def task(uid: int, level: int, name: str, mark: str = "Sí", start: datetime | None = None,
         finish: datetime | None = None, duration: float | None = None, wbs: str = "", **kw) -> Task:
    start = start or d(5)
    finish = finish or d(9, 17)
    return Task(uid=uid, id=uid, name=name, outline_level=level, outline_number="", wbs=wbs or str(uid),
                client_mark=mark, start=start, finish=finish, duration=duration, **kw)


def link(pred: Task, succ: Task, kind: int = LINK_FS, lag: float = 0.0, fmt: int = 7) -> None:
    succ.predecessors.append(Link(pred.uid, kind, lag, fmt))


def plan_of(tasks: list[Task], cfg: Config = CFG):
    return build_client_plan(P(tasks), cfg)


def root(uid: int = 1, name: str = "INTERNO") -> Task:
    return task(uid, 1, name, mark="", wbs="")


# ----------------------------------------------------------------------------
# Tareas que se muestran
# ----------------------------------------------------------------------------
def test_only_tasks_with_yes_in_the_whole_branch_are_shown():
    tasks = [root(),
             task(2, 2, "Fase A"), task(3, 3, "A1"), task(4, 3, "A2", "No"), task(5, 3, "A3", ""),
             task(6, 2, "Fase B", ""), task(7, 3, "B1"), task(8, 4, "B11"),
             task(9, 2, "Fase C", "quizá"),
             task(10, 1, "CLIENTE", wbs=""), task(11, 2, "Copia")]
    plan = plan_of(tasks)
    assert plan.shown == {2, 3}
    assert plan.marks == Counter({"Sí": 4, "No": 1, "vacío": 2, "otro valor": 1})
    assert [t.uid for t in plan.orphans] == [7] and plan.orphan_count == 2
    assert [t.uid for t in plan.unknown] == [9]
    assert [t.uid for t in plan.outside] == [10, 11]
    assert plan.root.uid == 1 and plan.new_name == "CLIENTE"
    assert [t.uid for t in plan.delete_roots] == [10, 4, 5, 6, 9]
    assert {i.id for i in plan.issues()} == {"CLI-003", "CLI-004"}


@pytest.mark.parametrize("value", ["Sí", "Si", "si", "SI", "SÍ", "S", "s", " Sí "])
def test_yes_values(value):
    assert plan_of([root(), task(2, 2, "A", value)]).shown == {2}


@pytest.mark.parametrize("value, category", [("No", "No"), ("n", "No"), ("x", "otro valor"), ("1", "otro valor"),
                                             ("Yes", "otro valor"), ("Sí, revisar", "otro valor")])
def test_only_si_shows_a_task(value, category):
    plan = plan_of([root(), task(2, 2, "A"), task(3, 2, "B", value)])
    assert plan.shown == {2} and plan.marks[category] == 1
    assert ("CLI-003" in {i.id for i in plan.issues()}) == (category == "otro valor")


def test_without_plan_the_whole_project_is_used():
    plan = plan_of([task(1, 1, "A"), task(2, 1, "B", "No")], Config(client_field="Etiqueta Cliente"))
    assert plan.shown == {1} and plan.root is None and plan.new_name == "" and plan.outside == []


def test_plan_must_be_a_top_level_task():
    tasks = [task(1, 1, "Programa", ""), task(2, 2, "INTERNO", "", wbs=""), task(3, 3, "A")]
    with pytest.raises(EvmError) as exc:
        plan_of(tasks)
    assert exc.value.id == "CLI-001"


def test_nothing_to_show():
    with pytest.raises(EvmError) as exc:
        plan_of([root(), task(2, 2, "A", "No"), task(3, 2, "B", "")])
    assert exc.value.id == "CLI-002"


def test_new_plan_name_from_config():
    plan = plan_of([root(), task(2, 2, "A")], Config(plan="INTERNO", client_field="X", client_plan_name="OBRA"))
    assert plan.new_name == "OBRA"


# ----------------------------------------------------------------------------
# Resúmenes convertidos y vínculos de sus subtareas
# ----------------------------------------------------------------------------
def collapsed_case(succ_start: datetime):
    tasks = [root(),
             task(2, 2, "Resumen", start=d(5), finish=d(9, 17)),
             task(3, 3, "Oculta 1", "No", start=d(5), finish=d(7, 17), duration=3 * DAY),
             task(4, 3, "Oculta 2", "No", start=d(8), finish=d(9, 17), duration=2 * DAY),
             task(5, 2, "Siguiente", start=succ_start, finish=d(16, 17), duration=5 * DAY)]
    link(tasks[2], tasks[3])
    link(tasks[2], tasks[4])
    return tasks, plan_of(tasks)


def test_summary_without_shown_subtasks_receives_their_links():
    tasks, plan = collapsed_case(d(12))
    assert [t.uid for t in plan.collapsed] == [2]
    assert plan.kept_links == 0 and plan.removed_links == 2
    [new] = plan.new_links
    assert (new.pred.uid, new.succ.uid, new.kind, new.lag, new.origin) == (2, 5, LINK_FS, 0.0, ORIGIN_SUMMARY)
    assert not new.adjusted
    assert {i.id for i in plan.issues()} == {"CLI-005", "CLI-006"}


def test_link_moved_to_a_summary_does_not_delay_its_successor():
    # La sucesora depende solo de la primera subtarea: desde el fin del resumen, la retrasaría. En vez de un adelanto
    # de 2 días desde el fin, el vínculo parte del comienzo del resumen, con 3 días de posición: la misma fecha
    tasks, plan = collapsed_case(d(8))
    [new] = plan.new_links
    assert new.adjusted and (new.kind, new.lag, new.slack) == (LINK_SS, 3 * DAY, 0)


# ----------------------------------------------------------------------------
# Puentes
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("kind_in, kind_out, a, h, b, kind, lag", [
    # A fin a comienzo de la oculta (3 días), y de esta a B: B comienza 3 días después del fin de A
    (LINK_FS, LINK_FS, (d(5), d(6, 17)), (d(7), d(9, 17)), d(12), LINK_FS, 3 * DAY),
    # Comienzo a comienzo y después fin a comienzo: B comienza 3 días después del comienzo de A
    (LINK_SS, LINK_FS, (d(5), d(6, 17)), (d(5), d(7, 17)), d(8), LINK_SS, 3 * DAY),
    # Fin a comienzo y comienzo a comienzo: la duración de la oculta no cuenta
    (LINK_FS, LINK_SS, (d(5), d(6, 17)), (d(7), d(9, 17)), d(7), LINK_FS, 0.0),
    # Fin a fin y comienzo a comienzo: B comienza 3 días antes del fin de A
    (LINK_FF, LINK_SS, (d(5), d(9, 17)), (d(7), d(9, 17)), d(7), LINK_FS, -3 * DAY),
])
def test_bridge_through_a_hidden_task(kind_in, kind_out, a, h, b, kind, lag):
    tasks = [root(), task(2, 2, "A", start=a[0], finish=a[1]),
             task(3, 2, "Oculta", "No", start=h[0], finish=h[1], duration=3 * DAY),
             task(4, 2, "B", start=b, finish=d(23, 17))]
    link(tasks[1], tasks[2], kind_in)
    link(tasks[2], tasks[3], kind_out)
    plan = plan_of(tasks)
    [new] = plan.new_links
    assert (new.pred.uid, new.succ.uid, new.kind, new.lag, new.origin, new.hidden) == (2, 4, kind, lag,
                                                                                      ORIGIN_BRIDGE, 1)
    assert not new.adjusted and new.slack == 0


def test_bridge_through_several_hidden_tasks_with_lags():
    tasks = [root(), task(2, 2, "A", start=d(5), finish=d(5, 17)),
             task(3, 2, "H1", "No", start=d(7), finish=d(7, 17), duration=DAY),
             task(4, 2, "H2", "No", start=d(8), finish=d(9, 17), duration=2 * DAY),
             task(5, 2, "B", start=d(13), finish=d(13, 17))]
    link(tasks[1], tasks[2], lag=DAY)            # 1 día de posición
    link(tasks[2], tasks[3])
    link(tasks[3], tasks[4], lag=DAY)
    [new] = plan_of(tasks).new_links
    assert (new.kind, new.lag, new.hidden, new.adjusted) == (LINK_FS, 5 * DAY, 2, False)


def test_one_link_per_pair_the_most_restrictive():
    tasks = [root(), task(2, 2, "A", start=d(5), finish=d(5, 17)),
             task(3, 2, "Corta", "No", start=d(6), finish=d(6, 17), duration=DAY),
             task(4, 2, "Larga", "No", start=d(6), finish=d(8, 17), duration=3 * DAY),
             task(5, 2, "B", start=d(9), finish=d(9, 17))]
    for h in tasks[2:4]:
        link(tasks[1], h)
        link(h, tasks[4])
    [new] = plan_of(tasks).new_links
    assert new.lag == 3 * DAY and new.slack == 0


def test_bridge_that_would_delay_is_adjusted():
    # La oculta terminó antes de lo que dice su duración (por ejemplo, con avance real): B empezó antes
    tasks = [root(), task(2, 2, "A", start=d(5), finish=d(5, 17)),
             task(3, 2, "Oculta", "No", start=d(6), finish=d(6, 17), duration=5 * DAY),
             task(4, 2, "B", start=d(7), finish=d(7, 17))]
    link(tasks[1], tasks[2])
    link(tasks[2], tasks[3])
    [new] = plan_of(tasks).new_links
    assert new.adjusted and new.lag == DAY


def test_elapsed_lag_takes_the_lag_from_the_dates():
    tasks = [root(), task(2, 2, "A", start=d(5), finish=d(5, 17)),
             task(3, 2, "Oculta", "No", start=d(8), finish=d(8, 17), duration=DAY),
             task(4, 2, "B", start=d(9), finish=d(9, 17))]
    link(tasks[1], tasks[2], lag=2 * 1440, fmt=8)          # 2 días transcurridos
    link(tasks[2], tasks[3])
    [new] = plan_of(tasks).new_links
    assert new.from_dates and new.lag == 3 * DAY and not new.adjusted


def test_tasks_already_linked_keep_their_link():
    tasks = [root(), task(2, 2, "A", start=d(5), finish=d(5, 17)),
             task(3, 2, "Oculta", "No", start=d(6), finish=d(6, 17), duration=DAY),
             task(4, 2, "B", start=d(7), finish=d(7, 17))]
    link(tasks[1], tasks[2])
    link(tasks[2], tasks[3])
    link(tasks[1], tasks[3])
    plan = plan_of(tasks)
    assert plan.new_links == [] and plan.kept_links == 1
    assert [(s.pred.uid, s.succ.uid, s.reason) for s in plan.skipped] == [(2, 4, SKIP_LINKED)]
    assert "CLI-007" not in {i.id for i in plan.issues()}


def test_bridge_that_would_close_a_cycle_is_not_created():
    tasks = [root(), task(2, 2, "A"), task(3, 2, "Oculta", "No", duration=DAY), task(4, 2, "B"), task(5, 2, "C")]
    a, h, b, c = tasks[1:]
    link(a, h)
    link(h, b)
    link(b, c)
    link(c, a)
    plan = plan_of(tasks)
    assert plan.new_links == []
    assert [(s.pred.uid, s.succ.uid, s.reason) for s in plan.skipped] == [(2, 4, SKIP_CYCLE)]
    assert "CLI-007" in {i.id for i in plan.issues()}


def test_no_link_between_a_summary_and_its_subtasks():
    tasks = [root(), task(2, 2, "Resumen"), task(3, 3, "X"), task(4, 3, "Y"), task(5, 2, "Oculta", "No", duration=DAY)]
    link(tasks[3], tasks[4])
    link(tasks[4], tasks[1])
    plan = plan_of(tasks)
    assert plan.new_links == []
    assert [(s.pred.uid, s.succ.uid, s.reason) for s in plan.skipped] == [(4, 2, SKIP_SUMMARY)]


def test_inactive_hidden_tasks_do_not_bridge():
    tasks = [root(), task(2, 2, "A"), task(3, 2, "Inactiva", "No", duration=DAY, active=False), task(4, 2, "B")]
    link(tasks[1], tasks[2])
    link(tasks[2], tasks[3])
    plan = plan_of(tasks)
    assert plan.new_links == [] and plan.removed_links == 2


def test_links_with_other_plans_disappear():
    tasks = [root(), task(2, 2, "A"), task(3, 1, "CLIENTE", wbs=""), task(4, 2, "Copia de A")]
    link(tasks[1], tasks[3], LINK_SS)
    link(tasks[3], tasks[1])
    plan = plan_of(tasks)
    assert plan.outside_links == 2 and plan.new_links == [] and plan.kept_links == 0


# ----------------------------------------------------------------------------
# Resúmenes que cambian de fechas
# ----------------------------------------------------------------------------
def test_summary_dates_set_by_a_hidden_subtask():
    tasks = [root(), task(2, 2, "Resumen", start=d(5), finish=d(9, 17)),
             task(3, 3, "Oculta", "No", start=d(5), finish=d(6, 17)),
             task(4, 3, "X", start=d(7), finish=d(9, 17))]
    plan = plan_of(tasks)
    [change] = plan.span_changes
    assert (change.task.uid, change.start, change.finish) == (2, d(7), d(9, 17))
    assert "CLI-008" in {i.id for i in plan.issues()}


# ----------------------------------------------------------------------------
# Avance de los resúmenes
# ----------------------------------------------------------------------------
def test_rollup_percent_like_ms_project():
    a = task(1, 1, "A", duration=2 * DAY, percent_complete=100)
    b = task(2, 1, "B", duration=6 * DAY, percent_complete=50)
    assert rollup_percent([a, b]) == 62.5                     # (2 × 100 + 6 × 50) / 8
    h1 = task(3, 1, "H1", duration=0, percent_complete=100)
    h2 = task(4, 1, "H2", duration=0, percent_complete=0)
    assert rollup_percent([h1, h2]) == 0 and rollup_percent([h1]) == 100   # solo hitos: 100 si están todos
    assert rollup_percent([]) is None


def test_summary_progress_with_hidden_and_collapsed_subtasks():
    tasks = [root(),
             task(2, 2, "Fase", percent_complete=50),
             task(3, 3, "Visible", duration=2 * DAY, percent_complete=0),
             task(4, 3, "Oculta", "No", duration=2 * DAY, percent_complete=100),
             task(5, 2, "Otra fase", percent_complete=50),
             task(6, 3, "Resumen", percent_complete=50, duration=4 * DAY),
             task(7, 4, "Oculta 1", "No", duration=2 * DAY, percent_complete=100),
             task(8, 4, "Oculta 2", "No", duration=2 * DAY, percent_complete=0)]
    plan = plan_of(tasks)
    # Fase pierde el avance de la oculta: 0 %. Otra fase conserva el 50 % (el resumen convertido lo lleva)
    assert [(c.task.uid, c.task.percent_complete, c.client, c.hidden) for c in plan.progress_changes] == [(2, 50, 0, 1)]
    assert "CLI-020" in {i.id for i in plan.issues()}
    text = render_control(plan, "INTERNO.mpp", date(2026, 11, 7), "Simulación")
    assert "## Avance de los resúmenes" in text and "| 2 (UID 2) | 50 % | 0 % | 1 tarea oculta |" in text


def test_converted_summary_weighs_by_its_own_duration():
    """Dos subtareas ocultas en paralelo (2 días cada una, una al 100 %): el resumen convertido conserva su 50 %, pero
    en el resumen superior pesa 2 días (su duración), no 4."""
    tasks = [root(),
             task(2, 2, "Fase", percent_complete=67),
             task(3, 3, "Visible", duration=2 * DAY, percent_complete=100),
             task(4, 3, "Resumen", duration=2 * DAY, percent_complete=50),
             task(5, 4, "Oculta 1", "No", duration=2 * DAY, percent_complete=100),
             task(6, 4, "Oculta 2", "No", duration=2 * DAY, percent_complete=0)]
    # Interno: (2 × 100 + 2 × 100 + 2 × 0) / 6 = 66,7 %; cliente: (2 × 100 + 2 × 50) / 4 = 75 %
    [change] = plan_of(tasks).progress_changes
    assert (change.task.uid, change.client, change.hidden, change.converted) == (2, 75.0, 0, 1)
    assert "1 resumen convertido" in change.reason


def test_working_minutes_signed():
    cal = WorkCalendar()
    assert working_minutes(cal, d(9, 17), d(12)) == 0
    assert working_minutes(cal, d(5), d(7)) == 2 * DAY
    assert working_minutes(cal, d(7), d(5)) == -2 * DAY


# ----------------------------------------------------------------------------
# Informe de control
# ----------------------------------------------------------------------------
def test_control_report_has_no_task_names():
    tasks, plan = collapsed_case(d(8))
    tasks.append(task(6, 2, "", "quizá"))
    for t in tasks[1:]:                                 # la raíz se busca por su nombre (INTERNO)
        t.name = f"Confidencial{t.uid}"
    plan = plan_of(tasks)
    text = render_control(plan, "INTERNO.mpp", date(2026, 11, 7), "Simulación: no se generó el archivo del cliente.")
    assert "Confidencial" not in text
    assert "## Resúmenes que pasan a ser tareas" in text and "## Vínculos nuevos" in text
    assert "| 2 (UID 2) | 5 (UID 5) | CC | 3 | Resumen | ajustada |" in text
    assert "CLI-003" in text and "«quizá»" in text


# ----------------------------------------------------------------------------
# Lectura y configuración
# ----------------------------------------------------------------------------
def client_xml(tmp_path, mark_a="Sí", mark_h="No"):
    pred = ("<PredecessorLink><PredecessorUID>{}</PredecessorUID><Type>1</Type><LinkLag>{}</LinkLag>"
            "<LagFormat>{}</LagFormat></PredecessorLink>")
    tasks = [task_xml(0, 0, "Programa"),
             task_xml(1, 1, "INTERNO", start=d(5), finish=d(9, 17)),
             task_xml(2, 2, "A", wbs="1", start=d(5), finish=d(5, 17), ext={TEXT15: mark_a},
                      extra="<Duration>PT8H0M0S</Duration>"),
             task_xml(3, 2, "Oculta", wbs="2", start=d(7), finish=d(7, 17), ext={TEXT15: mark_h},
                      extra="<Duration>PT8H0M0S</Duration>" + pred.format(2, 4800, 7)),
             task_xml(4, 2, "B", wbs="3", start=d(8), finish=d(8, 17), ext={TEXT15: "Sí"},
                      extra="<Duration>PT8H0M0S</Duration>" + pred.format(3, 0, 7))]
    return write_xml(tmp_path / "interno.xml",
                     project_xml(tasks, definitions=[(TEXT15, "Text15", "Etiqueta Cliente")]))


def test_xml_reads_the_client_field_and_lags_in_minutes(tmp_path):
    path = client_xml(tmp_path)
    project = read_xml(path, CFG, client=True)
    t = {x.name: x for x in project.tasks}
    assert (t["A"].client_mark, t["Oculta"].client_mark) == ("Sí", "No")
    assert t["Oculta"].predecessors == [Link(2, LINK_FS, 480.0, 7)]
    assert project.fields_found["cliente"] == "Etiqueta Cliente"
    assert read_xml(path, CFG).tasks[1].client_mark == ""        # solo generar-cliente lo lee
    [new] = build_client_plan(project, CFG).new_links
    assert (new.pred.name, new.succ.name, new.lag) == ("A", "B", 2 * DAY)


def test_com_reads_the_client_field_and_lag_format():
    from test_reader_com_fake import FakeTask, make_session

    class Dep:
        def __init__(self, a, b, lag, lag_type):
            self.From, self.To, self.Type, self.Lag, self.LagType = a, b, LINK_FS, lag, lag_type

    a, b = FakeTask(1, 1, "A", "1"), FakeTask(2, 1, "B", "2")
    a.Text15, b.Text15 = "Sí", "No"
    dep = Dep(a, b, 2880, 8)                                    # 2 días transcurridos
    a.TaskDependencies = b.TaskDependencies = [dep]
    cfg = Config(client_field="Text15")
    session = make_session([a, b])
    ta, tb = session.read_project(cfg, client=True).tasks
    assert (ta.client_mark, tb.client_mark, tb.predecessors) == ("Sí", "No", [Link(1, LINK_FS, 2880, 8)])
    ta, tb = session.read_project(cfg).tasks                    # los demás comandos no los leen
    assert (ta.client_mark, tb.predecessors) == ("", [Link(1, LINK_FS, 2880, 7)])


def test_config_client_section(tmp_path):
    path = tmp_path / "c.toml"
    path.write_text('[cliente]\ncampo = "Etiqueta Cliente"\nnombre_planificacion = "OBRA"\nfijar_fechas = false\n',
                    encoding="utf-8")
    cfg = load_config(path)
    assert (cfg.client_field, cfg.client_plan_name, cfg.client_fix_dates) == ("Etiqueta Cliente", "OBRA", False)
    assert (Config().client_field, Config().client_plan_name, Config().client_fix_dates) == ("", "CLIENTE", True)


# ----------------------------------------------------------------------------
# Comando
# ----------------------------------------------------------------------------
def write_config(tmp_path, field="Etiqueta Cliente"):
    path = tmp_path / "c.toml"
    path.write_text(f'[exportacion]\nplanificacion = "INTERNO"\n\n[cliente]\ncampo = "{field}"\n', encoding="utf-8")
    return str(path)


def test_cli_generar_cliente_simulation(tmp_path, capsys):
    path = client_xml(tmp_path)
    cfg = write_config(tmp_path)
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular"]) == 0
    out = capsys.readouterr().out
    assert "Se muestran: 2" in out and "No se muestran: 1" in out and "1 nuevos" in out
    assert "Simulación: no se generó el archivo del cliente." in out
    report = tmp_path / "interno_cliente_control.md"
    assert report.is_file() and "Oculta" not in report.read_text(encoding="utf-8")
    assert not (tmp_path / "interno_cliente.mpp").exists()
    # El informe existe: hace falta --sobrescribir
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular"]) == 2
    assert "CFG-042" in capsys.readouterr().err
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular", "--sobrescribir",
                 "--destino", str(tmp_path / "cliente.mpp")]) == 0
    assert (tmp_path / "cliente_control.md").is_file()


def test_cli_generar_cliente_errors(tmp_path, capsys):
    path = client_xml(tmp_path)
    assert main(["generar-cliente", "--origen", str(path), "--simular"]) == 2
    assert "CFG-054" in capsys.readouterr().err
    assert main(["generar-cliente", "--origen", str(path), "--config", write_config(tmp_path, "Otro campo"),
                 "--simular"]) == 2
    assert "LEC-026" in capsys.readouterr().err
    cfg = write_config(tmp_path)
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg]) == 2
    assert "CLI-010" in capsys.readouterr().err
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular",
                 "--destino", str(tmp_path / "cliente.xml")]) == 2
    assert "CFG-041" in capsys.readouterr().err
