"""Peso = 0 escrito en el campo de peso (desde la 2.7.0): la tarea se hace, pero no cuenta en el libro EVM. Peso
vacío: sin definir, depende de sus tareas hermanas. En un campo numérico el 0 no se distingue del vacío."""
from datetime import datetime

import pytest

from evm.config import Config
from evm.messages import SEV_WARN
from evm.percent import RowTree, task_time
from evm.reader_xml import read_xml
from evm.wbs import KIND_RESIDUAL, build_report
from evm.workcalendar import WorkCalendar

from builders import P, T, rows_by_code
from mspdi import NUMBER1, TEXT3, project_xml, task_xml, write_xml

JAN = (datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17))
FEB = (datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17))
MAR = (datetime(2026, 3, 2, 8), datetime(2026, 3, 27, 17))


def cfg(**kw) -> Config:
    c = Config(max_level=2, months=12, milestone_weight_field="Peso", **kw)
    c.validate()
    return c


def task(uid, level, name, period, weight=None, code="", **kw):
    return T(uid, level, code, name, bs=period[0], bf=period[1], start=period[0], finish=period[1],
             milestone_weight=weight, **kw)


def report(tasks, **kw):
    return build_report(P(tasks), cfg(**kw))


def ids(rep, code):
    return [i for i in rep.issues if i.id == code]


def test_zero_does_not_count_and_empty_takes_the_rest_by_duration():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="Ponderado"),
             task(3, 3, "Con peso", JAN, 60), task(4, 3, "Cero", FEB, 0), task(5, 3, "Vacía", MAR)]
    rep = report(tasks)
    pv = rows_by_code(rep)["1.1"].pct_vec[""]
    assert pv[0] == pytest.approx(0.60) and pv[1] == pytest.approx(0.0) and pv[2] == pytest.approx(0.40)
    assert not ids(rep, "HIT-011") and len(ids(rep, "HIT-015")) == 1


def test_the_zero_branch_is_not_used_by_other_methods():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="% Completado"),
             task(3, 3, "A", JAN), task(4, 3, "Fuera", FEB, 0), task(5, 4, "Fuera hija", FEB), task(6, 3, "B", MAR)]
    rep = report(tasks)
    row = rows_by_code(rep)["1.1"]
    assert [t.name for t in row.leaves] == ["A", "B"]
    assert row.pct_vec[""][1] == pytest.approx(0.0)


def test_a_zero_task_does_not_form_a_residual_row():
    def tasks(weight):
        return [T(1, 1, "1", "Fase"), task(2, 2, "Paquete", JAN, code="1.1"), task(3, 2, "Suelta", FEB, weight)]
    assert [r.code for r in report(tasks(None)).rows if r.kind == KIND_RESIDUAL] == ["1.0-D"]
    assert [r.code for r in report(tasks(0)).rows if r.kind == KIND_RESIDUAL] == []


def test_a_row_with_zero_weight_is_kept_with_nothing_planned():
    tasks = [T(1, 1, "1", "Fase"), task(2, 2, "Paquete cero", JAN, 0, code="1.1"), task(3, 3, "Detalle", JAN),
             task(4, 2, "Otro", FEB, code="1.2")]
    rep = report(tasks)
    rows = rows_by_code(rep)
    assert "1.1" in rows and sum(rows["1.1"].pct_vec[""]) == pytest.approx(0.0)
    assert sum(rows["1.2"].pct_vec[""]) == pytest.approx(1.0)
    [issue] = ids(rep, "HIT-016")
    assert issue.code == "1.1" and issue.severity == SEV_WARN


def test_weights_inside_a_zero_task_are_reported():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="Ponderado"),
             task(3, 3, "Bloque", JAN, 0), task(4, 4, "Con peso", JAN, 50), task(5, 3, "Resto", FEB)]
    rep = report(tasks)
    [issue] = ids(rep, "HIT-014")
    assert issue.severity == SEV_WARN and "Bloque" in issue.name
    assert sum(rows_by_code(rep)["1.1"].pct_vec[""]) == pytest.approx(1.0)


def test_weighted_row_skips_zero_tasks_without_pruning():
    """actualizar no quita del árbol las tareas con Peso = 0: el reparto las salta."""
    tasks = [T(1, 1, "1.1", "Paquete"), task(2, 2, "Con peso", JAN, 50), task(3, 2, "Cero", FEB, 0),
             task(4, 2, "Vacía", MAR)]
    from evm.wbs import build_tree
    build_tree(tasks)
    calendar = WorkCalendar()
    tree = RowTree(tasks[0], tasks[1:], lambda t: task_time(t, calendar))
    names = {t.uid: t.name for t in tasks}
    assert {names[uid]: pytest.approx(f) for uid, f in tree.full()[0].weights.items()} == {"Con peso": 0.5,
                                                                                          "Vacía": 0.5}


@pytest.mark.parametrize("field, number, expected", [(TEXT3, False, [0.0, None, 20.0]),
                                                     (NUMBER1, True, [None, None, 20.0])])
def test_readers_tell_zero_from_empty_only_in_text_fields(tmp_path, field, number, expected):
    tasks = [task_xml(0, 0, "Proyecto"),
             task_xml(1, 1, "Cero", wbs="1", ext={field: "0"}),
             task_xml(2, 1, "Vacía", wbs="2"),
             task_xml(3, 1, "Veinte", wbs="3", ext={field: "20"})]
    name = "Number1" if number else "Text3"
    path = write_xml(tmp_path / "p.xml", project_xml(tasks, definitions=[(field, name, "Peso")]))
    project = read_xml(path, cfg())
    assert [t.milestone_weight for t in project.tasks if t.outline_level == 1] == expected


def test_a_summary_whose_tasks_all_have_zero_does_not_count_either():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="% Completado"),
             task(3, 3, "A", JAN), task(4, 3, "Grupo", FEB), task(5, 4, "Cero 1", FEB, 0), task(6, 4, "Cero 2", FEB, 0),
             task(7, 2, "Otro", MAR, code="1.2"), task(8, 3, "Solo cero", MAR, 0)]
    rep = report(tasks)
    rows = rows_by_code(rep)
    assert [t.name for t in rows["1.1"].leaves] == ["A"]
    assert sum(rows["1.2"].pct_vec[""]) == pytest.approx(0.0)
    [issue] = ids(rep, "HIT-016")
    assert issue.code == "1.2" and issue.message.startswith("Todas las tareas de la fila tienen Peso = 0")


def test_actualizar_gives_the_same_earned_value_as_generar(tmp_path):
    from evm.update import run_update
    from evm.workbook import write_workbook
    from builders import set_bac

    def tasks(progress):
        return [T(1, 1, "1", "Fase"),
                T(2, 2, "1.1", "Ponderado", method="Ponderado"),
                task(3, 3, "Con peso", JAN, 60, percent_complete=progress, actual_start=JAN[0]),
                task(4, 3, "Cero", FEB, 0, percent_complete=100, actual_start=FEB[0], actual_finish=FEB[1]),
                task(5, 3, "Vacía", MAR, percent_complete=progress, actual_start=MAR[0]),
                T(6, 2, "1.2", "Completado", method="% Completado"),
                task(7, 3, "A", JAN, percent_complete=progress, actual_start=JAN[0]),
                task(8, 3, "Fuera", FEB, 0, percent_complete=100, actual_start=FEB[0], actual_finish=FEB[1])]
    c = cfg()
    book = tmp_path / "libro.xlsx"
    write_workbook(build_report(P(tasks(0)), c), c, book)
    set_bac(book)
    res = run_update(P(tasks(50.0), status_date=datetime(2026, 3, 31)), c, book, backup=False, dry_run=True)
    ev = {r.code: r.ev_cum for r in res.rows}
    assert ev["1.1"] == pytest.approx(0.50 * 1_000_000, abs=1)      # 60 × 50 % + 40 × 50 %; «Cero» no cuenta
    assert ev["1.2"] == pytest.approx(0.50 * 1_000_000, abs=1)      # solo «A»


ENTERPRISE = 188776001          # campo de tarea que no es Número1..20 ni Texto1..30 (p. ej. empresarial)


@pytest.mark.parametrize("values, expected", [
    (["0", "0", "0", "20"], [None, None, None, 20.0]),     # numérico: devuelve 0 en las tareas sin peso
    (["", "", "0", "20"], [None, None, 0.0, 20.0]),        # texto: un 0 escrito es Peso = 0
])
def test_weight_field_of_unknown_type_is_read_by_its_values(tmp_path, values, expected):
    tasks = [task_xml(0, 0, "Proyecto")] + [
        task_xml(i, 1, f"T{i}", wbs=str(i), ext={ENTERPRISE: v} if v else None) for i, v in enumerate(values, start=1)]
    path = write_xml(tmp_path / "p.xml", project_xml(tasks, definitions=[(ENTERPRISE, "Enterprise1", "Peso")]))
    project = read_xml(path, cfg())
    assert [t.milestone_weight for t in project.tasks if t.outline_level == 1] == expected


@pytest.mark.parametrize("cf_type, values, expected", [
    # Número: el XML omite los 0, pero un 0 escrito sigue siendo vacío (con pocos ceros, la regla de los valores
    # lo tomaría como Peso = 0)
    ("5", ["", "", "0", "20"], [None, None, None, 20.0]),
    ("0", ["", "0", "", "20"], [None, None, None, 20.0]),               # costo: numérico
    ("7", ["0", "0", "0", "20"], [0.0, 0.0, 0.0, 20.0]),                 # texto: el 0 es Peso = 0
])
def test_xml_enterprise_weight_field_type_comes_from_its_cftype(tmp_path, cf_type, values, expected):
    tasks = [task_xml(0, 0, "Proyecto")] + [
        task_xml(i, 1, f"T{i}", wbs=str(i), ext={ENTERPRISE: v} if v else None) for i, v in enumerate(values, start=1)]
    path = write_xml(tmp_path / "p.xml",
                     project_xml(tasks, definitions=[(ENTERPRISE, "Enterprise1", "Peso", cf_type)]))
    project = read_xml(path, cfg())
    assert [t.milestone_weight for t in project.tasks if t.outline_level == 1] == expected


def test_com_weight_field_kind():
    from evm.fields import TASK_NUMBER_IDS, TASK_TEXT_IDS
    from evm.reader_com import weight_kind
    assert weight_kind(("attr", "Number3")) is True and weight_kind(("attr", "Text7")) is False
    assert weight_kind(("fid", TASK_NUMBER_IDS[1])) is True and weight_kind(("fid", TASK_TEXT_IDS[7])) is False
    assert weight_kind(("fid", ENTERPRISE)) is None and weight_kind(None) is False
