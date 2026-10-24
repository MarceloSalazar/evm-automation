from datetime import date, datetime

import pytest

from evm.config import Config, load_config
from evm.model import ProjectData, Task
from evm.reader_xml import read_xml
from evm.wbs import KIND_LEAF, KIND_RESIDUAL, build_report


def test_read_example_xml(example_files):
    cfg = load_config(example_files["config"])
    p = read_xml(example_files["base"], cfg)
    assert p.name == "Programa de ejemplo EVM"
    assert p.currency_symbol == "USD"
    assert p.baseline_found
    by_name = {}
    for t in p.tasks:                       # primero aparecen las tareas de INTERNO
        by_name.setdefault(t.name, t)
    assert by_name["Licitación de ingeniería"].method == "0/100"          # alias "Método EVM"
    assert by_name["Hito: bases aprobadas"].milestone_weight == 30         # alias "Peso hito"
    assert by_name["Fundaciones"].baseline_start is not None and by_name["Fundaciones"].baseline_finish


def test_report_limits_level_and_planned_percent(example_files):
    cfg = load_config(example_files["config"])
    rep = build_report(read_xml(example_files["base"], cfg), cfg)
    assert max(r.level for r in rep.rows) == 3
    inputs = [r for r in rep.rows if r.is_input]
    assert len(inputs) == 24 and all(sum(r.pct_vec[""]) == pytest.approx(1.0) for r in inputs)
    assert sum(rep.total_vector) == pytest.approx(24)       # suma de los % planificados de las filas
    codes = [r.code for r in rep.rows]
    assert codes[0] == "1.0.0" and "2.1.2" in codes and "2.1.0" in codes
    assert "INTERNO" not in codes and not any(c.count(".") > 2 for c in codes)
    conceptual = next(r for r in rep.rows if r.name == "Diseño conceptual")
    assert conceptual.kind == KIND_LEAF and conceptual.hidden_tasks == 6
    assert rep.horizon.start == date(2026, 1, 1)
    assert not [i for i in rep.issues if i.severity != "Info"]


def _t(uid, level, name, start=None, finish=None, **kw):
    return Task(uid=uid, id=uid, name=name, outline_level=level, wbs=kw.pop("wbs", ""),
                baseline_start=start, baseline_finish=finish, has_baseline=start is not None, start=start,
                finish=finish, **kw)


def _project(tasks, **kw):
    return ProjectData(name="P", source="p.xml", backend="xml", tasks=tasks, **kw)


def test_loose_tasks_of_a_summary_form_the_other_tasks_row():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)
    tasks = [
        _t(1, 1, "Fase", s, f, wbs="1"),
        _t(2, 2, "Grupo", s, f, wbs="1.1"),
        _t(3, 3, "Paquete", s, f, wbs="1.1.1"),
        _t(4, 3, "Reunión sin código", s, f),
    ]
    rep = build_report(_project(tasks), Config())
    kinds = [(r.code, r.kind) for r in rep.rows]
    assert ("1.1.0-D", KIND_RESIDUAL) in kinds
    residual = next(r for r in rep.rows if r.kind == KIND_RESIDUAL)
    assert residual.name.startswith("Otras tareas") and [t.name for t in residual.leaves] == ["Reunión sin código"]


def test_planned_percent_before_horizon_goes_to_month_1_and_after_is_reported():
    tasks = [_t(1, 1, "A", datetime(2025, 11, 3, 8), datetime(2026, 2, 27, 17), wbs="1")]
    cfg = Config(start_month=date(2026, 1, 1), months=1)
    rep = build_report(_project(tasks), cfg)
    row = rep.rows[0]
    before = sum(v for k, v in row.pct[""].items() if k < (2026, 1))
    assert row.vector[0] == pytest.approx(before + row.pct[""][(2026, 1)])
    assert any("posterior" in i.message for i in rep.issues) and any("anterior" in i.message for i in rep.issues)


def test_inactive_tasks_excluded():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    tasks = [_t(1, 1, "A", s, f, wbs="1"), _t(2, 1, "B", s, f, wbs="2", active=False)]
    rep = build_report(_project(tasks), Config())
    assert [r.name for r in rep.rows] == ["A"]
    assert any("tarea inactiva" in i.message for i in rep.issues)
    rep2 = build_report(_project(tasks), Config(include_inactive=True))
    assert len(rep2.rows) == 2


def test_task_without_baseline_dates_uses_current_dates():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    t_new = _t(2, 1, "Nueva", None, None, wbs="2")
    t_new.start, t_new.finish = datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17)
    rep = build_report(_project([_t(1, 1, "A", s, f, wbs="1"), t_new]), Config())
    assert any("fechas de línea base" in i.message for i in rep.issues)
    new_row = next(r for r in rep.rows if r.name == "Nueva")
    assert new_row.pct[""] == {(2026, 2): pytest.approx(1.0)}


def test_xml_without_baseline_uses_current_schedule(tmp_path):
    xml = tmp_path / "sin_lb.xml"
    xml.write_text("""<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project"><Title>Sin LB</Title><Tasks>
<Task><UID>0</UID><ID>0</ID><Name>Sin LB</Name><OutlineLevel>0</OutlineLevel></Task>
<Task><UID>1</UID><ID>1</ID><Name>A</Name><WBS>1</WBS><OutlineLevel>1</OutlineLevel>
<Start>2026-01-05T08:00:00</Start><Finish>2026-01-30T17:00:00</Finish></Task>
<Task><UID>2</UID><ID>2</ID><IsNull>1</IsNull><OutlineLevel>0</OutlineLevel></Task>
</Tasks></Project>""", encoding="utf-8")
    p = read_xml(xml, Config())
    assert not p.baseline_found
    rep = build_report(p, Config())
    assert rep.rows[0].pct[""] == {(2026, 1): pytest.approx(1.0)}
    assert any("no tiene guardada la línea base" in i.message for i in rep.issues)
