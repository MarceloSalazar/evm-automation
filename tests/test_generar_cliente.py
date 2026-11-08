"""Planificación para el cliente (generar-cliente): tareas que se muestran, resúmenes convertidos, vínculos nuevos e
informe de control."""
from __future__ import annotations

from collections import Counter
from datetime import date, datetime

import pytest

from evm.cli import main
from evm.client_check import client_model, compare_tasks, predict
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
# Borrado de las tareas que no se muestran (actividad_interna = false) y «Otras tareas» (por defecto)
CFG = Config(plan="INTERNO", client_field="Etiqueta Cliente", client_internal_activity=False)
# El código EDT de las pruebas está en un campo de texto: se puede quitar de las «Otras tareas»
ANON = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="Text5")


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
    assert [t.uid for t in plan.unmarked] == [5]        # sin valor debajo de una que se muestra (no Fase B)
    assert {i.id for i in plan.issues()} == {"CLI-003", "CLI-004", "CLI-031"}


def test_tasks_without_mark_below_a_shown_task():
    """Una tarea nueva sin marcar debajo de una que se muestra: se avisa (CLI-031) con la de más arriba de cada
    rama. No se avisa de las que tienen No, de las inactivas ni de las de primer nivel."""
    tasks = [root(),
             task(2, 2, "Fase A", wbs="5"), task(3, 3, "Nueva", "", wbs="5.1"), task(4, 4, "Nueva 1", "", wbs="5.1.1"),
             task(5, 4, "Nueva 2", "Sí", wbs="5.1.2"), task(6, 3, "Interna", "No", wbs="5.2"),
             task(7, 3, "Cancelada", "", wbs="5.3", active=False), task(8, 3, "Otra nueva", " ", wbs="5.4"),
             task(9, 2, "Sin marcar", "", wbs="6")]
    plan = plan_of(tasks, ANON)
    assert [t.uid for t in plan.unmarked] == [3, 8]
    message = {i.id: i.message for i in plan.issues()}["CLI-031"]
    assert message.startswith("2 tareas no tienen valor en 'Etiqueta Cliente' y su tarea superior se muestra (4 con "
                              "sus subtareas): quedan como «Otras tareas». Si el cliente las debe ver")
    assert "5.1 (UID 3), 5.4 (UID 8)." in message
    report = render_control(plan, "origen.xml", date(2026, 11, 8), "Simulación")
    assert "| Sin valor y con la tarea superior que se muestra (sin sus subtareas) | 2 |" in report
    assert "## Tareas sin valor debajo de una tarea que se muestra" in report
    assert "| 5.1 (UID 3) | 5 (UID 2) | 3 |" in report and "| 5.4 (UID 8) | 5 (UID 2) | 1 |" in report
    # Borrado: no se muestran; una sola tarea
    plan = plan_of(tasks[:3] + tasks[5:7], CFG)
    message = {i.id: i.message for i in plan.issues()}["CLI-031"]
    assert message.startswith("1 tarea no tiene valor en 'Etiqueta Cliente' y su tarea superior se muestra: no se "
                              "muestra. Si el cliente la debe ver")


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


# ----------------------------------------------------------------------------
# actividad_interna: las tareas que no se muestran quedan como «Otras tareas»
# ----------------------------------------------------------------------------
def internal_case():
    """Fase 1 con A (Sí, 2 días al 100 %), Interna (No, 3 días al 100 %) y B (Sí, 2 días al 0 %), encadenadas; una
    tarea inactiva que no se muestra y la planificación CLIENTE."""
    tasks = [task(1, 1, "INTERNO", mark="", wbs="INTERNO"),
             task(2, 2, "Fase", wbs="1", start=d(5), finish=d(13, 17)),
             task(3, 3, "A", wbs="1.1", start=d(5), finish=d(6, 17), duration=2 * DAY, percent_complete=100),
             task(4, 3, "Interna", "No", wbs="1.2", start=d(7), finish=d(9, 17), duration=3 * DAY,
                  percent_complete=100),
             task(5, 3, "B", wbs="1.3", start=d(12), finish=d(13, 17), duration=2 * DAY),
             task(6, 3, "Inactiva", "", wbs="1.4", start=d(12), finish=d(12, 17), duration=DAY, active=False),
             task(7, 1, "CLIENTE", wbs=""), task(8, 2, "Copia")]
    link(tasks[2], tasks[3])
    link(tasks[3], tasks[4])
    link(tasks[5], tasks[4])
    return tasks


def test_internal_activity_keeps_the_hidden_tasks():
    plan = plan_of(internal_case(), ANON)
    assert plan.internal_activity and plan.codes_removed
    assert [t.uid for t in plan.anonymous] == [4] and [t.uid for t in plan.dropped] == [6]
    assert [t.uid for t in plan.delete_roots] == [7, 6]             # otra planificación e inactiva oculta
    assert plan.remaining == {1, 2, 3, 4, 5}
    assert (plan.kept_links, plan.removed_links) == (2, 1) and plan.kept_pairs == {(3, 4), (4, 5)}
    assert plan.new_links == [] and plan.collapsed == [] and plan.progress_changes == []
    issues = {i.id: i.message for i in plan.issues()}
    assert "1 tarea que no se muestra queda como «Otras tareas» (sin nombre ni código EDT)" in issues["CLI-022"]
    assert "1 tarea inactiva se quita" in issues["CLI-022"]


def test_internal_activity_keeps_the_percent_and_the_book():
    """Con «Otras tareas», el % del proyecto y el libro EVM son los de la planificación interna; la fila de la tarea
    interna, sin código, pasa a la fila -D de su resumen (mismas tareas y valores) y se avisa."""
    plan = plan_of(internal_case(), ANON)
    model = client_model(plan)
    t = {x.uid: x for x in model.tasks}
    assert sorted(t) == [1, 2, 3, 4, 5]
    assert (t[4].name, t[4].wbs, t[4].client_mark, t[4].percent_complete) == ("Otras tareas", "", "", 100)
    assert (t[1].name, t[3].name, t[3].wbs) == ("CLIENTE", "A", "1.1")
    comp = predict(plan, ANON)
    assert comp.percent_internal == comp.percent_client == pytest.approx(500 / 7)      # (2 × 100 + 3 × 100) / 7
    book = comp.book
    assert (book.same, book.deviations, book.leaves_internal, book.leaves_client) == (2, 0, 3, 3)
    assert [(a.code, b.code) for a, b in book.moved] == [("1.2.0", "1.0.0-D")] and book.d_rows == ["1.0.0-D"]
    plan.comparison = comp
    issues = {i.id: i.message for i in plan.issues()}
    assert "CLI-023" not in issues and "CLI-024" not in issues
    assert "1 fila -D («Otras tareas de …»): 1.0.0-D (antes 1.2.0)" in issues["CLI-025"]


def test_without_internal_activity_the_deviations_are_reported():
    plan = plan_of(internal_case())
    plan.comparison = comp = predict(plan, CFG)
    assert comp.percent_internal == pytest.approx(500 / 7) and comp.percent_client == pytest.approx(50)
    book = comp.book
    assert [r.code for r in book.only_internal] == ["1.2.0"] and book.where["1.2.0"] == []
    assert (book.same, book.deviations, book.leaves_client) == (2, 1, 2)
    issues = {i.id: i.message for i in plan.issues()}
    assert "sería 50 % en el archivo del cliente y es 71,4 % en la planificación interna" in issues["CLI-023"]
    assert "1 fila del libro EVM del cliente tendría otras tareas o valores" in issues["CLI-024"]
    text = render_control(plan, "INTERNO.mpp", date(2026, 11, 8), "Simulación")
    assert "| 1.2.0 | — (sus tareas: ninguna) | 1 / 0 |" in text


def test_merged_rows_show_where_their_tasks_are():
    """Si no se muestra ningún elemento EDT de un resumen, su fila reúne las tareas de las filas internas."""
    tasks = [task(1, 1, "INTERNO", mark="", wbs="INTERNO"), task(2, 2, "Fase", wbs="1"),
             task(3, 3, "X", "No", wbs="1.1", duration=DAY), task(4, 3, "Y", "No", wbs="1.2", duration=DAY),
             task(5, 2, "Otra", wbs="2", duration=DAY)]
    plan = plan_of(tasks, ANON)
    book = predict(plan, ANON).book
    assert sorted(r.code for r in book.only_internal) == ["1.1.0", "1.2.0"]
    assert [r.code for r in book.only_client] == ["1.0.0"]
    assert book.where == {"1.1.0": ["1.0.0"], "1.2.0": ["1.0.0"], "1.0.0": ["1.1.0", "1.2.0"]}


def test_compare_tasks_reports_each_change():
    plan = plan_of(internal_case(), ANON)
    model = client_model(plan)
    t = {x.uid: x for x in model.tasks}
    t[5].start, t[4].percent_complete = d(14), 40
    model.tasks.remove(t[3])
    diffs, checked = compare_tasks(plan, model)
    assert checked == 5
    found = {x.task.uid: x.fields for x in diffs}
    assert found[3] == [("Tarea", "Sí", "No está")]
    assert found[4] == [("% completado", "100 %", "40 %")]
    assert found[5] == [("Comienzo", "12/01/2026 08:00", "14/01/2026 08:00")]


def test_codes_in_a_builtin_field_cannot_be_removed():
    cfg = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="EDT")
    plan = plan_of(internal_case(), cfg)
    assert not plan.codes_removed and client_model(plan).tasks[3].wbs == "1.2"
    assert "conservan su código EDT" in {i.id: i.message for i in plan.issues()}["CLI-027"]


def test_inactive_hidden_tasks_are_kept_when_the_book_includes_them():
    cfg = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="Text5", include_inactive=True)
    plan = plan_of(internal_case(), cfg)
    assert plan.dropped == [] and [t.uid for t in plan.anonymous] == [4, 6]


def test_control_report_with_internal_activity():
    tasks = internal_case()
    for x in tasks[1:]:
        x.name = f"Confidencial{x.uid}"
    plan = plan_of(tasks, ANON)
    plan.comparison = predict(plan, ANON)
    text = render_control(plan, "INTERNO.mpp", date(2026, 11, 8), "Simulación")
    assert "Confidencial" not in text
    assert "| **Tareas que no se muestran** | Quedan como «Otras tareas»" in text
    assert "| Quedan como «Otras tareas» | 1 (0 resúmenes; sin código EDT) |" in text
    assert "## Comparación con la planificación interna" in text and "| 1.0.0-D | 1.2.0 |" in text
    assert "## Ramas que no se muestran" in text and "| 1.2 (UID 4) | 1 | 0 |" in text
    assert "## Vínculos nuevos" not in text


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
    path.write_text('[cliente]\ncampo = "Etiqueta Cliente"\nnombre_planificacion = "OBRA"\nfijar_fechas = false\n'
                    'actividad_interna = false\nfiltrar_vista = true\n', encoding="utf-8")
    cfg = load_config(path)
    assert (cfg.client_field, cfg.client_plan_name, cfg.client_fix_dates) == ("Etiqueta Cliente", "OBRA", False)
    assert (cfg.client_internal_activity, cfg.client_filter) == (False, True)
    default = Config()
    assert (default.client_field, default.client_plan_name, default.client_fix_dates) == ("", "CLIENTE", True)
    assert (default.client_internal_activity, default.client_filter) == (True, False)


# ----------------------------------------------------------------------------
# Comando
# ----------------------------------------------------------------------------
def write_config(tmp_path, field="Etiqueta Cliente", extra=""):
    path = tmp_path / "c.toml"
    path.write_text(f'[exportacion]\nplanificacion = "INTERNO"\n\n[cliente]\ncampo = "{field}"\n{extra}',
                    encoding="utf-8")
    return str(path)


def test_cli_generar_cliente_simulation(tmp_path, capsys):
    path = client_xml(tmp_path)
    cfg = write_config(tmp_path)
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular"]) == 0
    out = capsys.readouterr().out
    assert "Se muestran: 2" in out and "No se muestran: 1" in out
    assert "Tareas que no se muestran: 1 como «Otras tareas», 0 inactivas se quitan" in out
    assert "% completado del proyecto: 0 % (interno) y 0 % (cliente)" in out
    # El campo EDT es el de MS Project: las «Otras tareas» conservan su código y el libro es el mismo
    assert "Libro EVM: 3 filas iguales, 0 con otro código, 0 con otras tareas o valores" in out
    assert "CLI-027" in out and "Simulación: no se generó el archivo del cliente." in out
    report = tmp_path / "interno_cliente_control.md"
    assert report.is_file() and "Oculta" not in report.read_text(encoding="utf-8")
    assert not (tmp_path / "interno_cliente.mpp").exists()
    # El informe existe: hace falta --sobrescribir
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular"]) == 2
    assert "CFG-042" in capsys.readouterr().err
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular", "--sobrescribir",
                 "--destino", str(tmp_path / "cliente.mpp")]) == 0
    assert (tmp_path / "cliente_control.md").is_file()


def test_cli_generar_cliente_simulation_deleting_hidden_tasks(tmp_path, capsys):
    path = client_xml(tmp_path)
    cfg = write_config(tmp_path, extra="actividad_interna = false\n")
    assert main(["generar-cliente", "--origen", str(path), "--config", cfg, "--simular"]) == 0
    out = capsys.readouterr().out
    assert "Tareas que no se muestran: se borran (actividad_interna = false)" in out and "1 nuevos" in out
    assert "Libro EVM: 2 filas iguales, 0 con otro código, 1 con otras tareas o valores" in out and "CLI-024" in out


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
