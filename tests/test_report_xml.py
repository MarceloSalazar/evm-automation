from datetime import date, datetime

import pytest

from evm.config import Config, load_config
from evm.model import ProjectData, Task
from evm.reader_xml import read_xml, timephased_baseline_cost_type
from evm.wbs import KIND_LEAF, KIND_RESIDUAL, LinearPVSource, TimephasedPVSource, build_report


def xml_source(cfg):
    return TimephasedPVSource(lambda t: t.timephased_pv, "xml", cfg.distribution, cfg.holidays)


def test_timephased_type_numbers():
    assert timephased_baseline_cost_type(0) == 10
    assert timephased_baseline_cost_type(1) == 19
    assert timephased_baseline_cost_type(10) == 73


def test_read_example_xml(example_files):
    cfg = load_config(example_files["config"])
    p = read_xml(example_files["base"], cfg)
    assert p.name == "Programa de ejemplo EVM"
    assert p.currency_symbol == "USD"
    assert p.baseline_found
    # Montos del XML están multiplicados por 100
    assert p.summary_baseline_cost == pytest.approx(44_210_000)
    by_name = {t.name: t for t in p.tasks}
    assert by_name["Licitación de ingeniería"].method == "0/100"          # alias "Método EVM"
    assert by_name["Hito: bases aprobadas"].milestone_weight == 30         # alias "Peso hito"
    assert by_name["Licitación de ingeniería"].timephased_pv
    assert sum(by_name["Fundaciones"].timephased_pv.values()) == pytest.approx(5_400_000)


def test_report_limits_level_and_totals(example_files):
    cfg = load_config(example_files["config"])
    p = read_xml(example_files["base"], cfg)
    src = xml_source(cfg)
    rep = build_report(p, cfg, src)
    assert max(r.level for r in rep.rows) == 3
    assert rep.total_bac == pytest.approx(44_210_000)
    assert sum(rep.total_vector) == pytest.approx(44_210_000)
    assert src.timephased_count > 0 and src.linear_count == 0
    codes = [r.code for r in rep.rows]
    assert codes[0] == "1.0.0" and "2.1.2" in codes
    conceptual = next(r for r in rep.rows if r.name == "Diseño conceptual")
    assert conceptual.kind == KIND_LEAF and conceptual.hidden_tasks == 6
    assert conceptual.bac == pytest.approx(600_000)
    assert rep.horizon.start == date(2026, 1, 1)
    assert not [i for i in rep.issues if i.severity != "Info"]


def _t(uid, level, name, cost=0.0, start=None, finish=None, **kw):
    return Task(uid=uid, id=uid, name=name, outline_level=level, wbs=kw.pop("wbs", ""), baseline_cost=cost,
                baseline_start=start, baseline_finish=finish, has_baseline=start is not None, cost=cost, **kw)


def _project(tasks, **kw):
    return ProjectData(name="P", source="p.xml", backend="xml", tasks=tasks, **kw)


def test_summary_with_own_cost_creates_residual_row():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)
    tasks = [
        _t(1, 1, "Fase", 1500, s, f, wbs="1"),
        _t(2, 2, "Grupo", 1500, s, f, wbs="1.1"),       # 1000 propios + 500 de la hija
        _t(3, 3, "Paquete", 500, s, f, wbs="1.1.1"),
    ]
    tasks[1].children = []
    cfg = Config()
    rep = build_report(_project(tasks), cfg, LinearPVSource())
    kinds = [(r.code, r.kind) for r in rep.rows]
    assert ("1.1.0-D", KIND_RESIDUAL) in kinds
    residual = next(r for r in rep.rows if r.kind == KIND_RESIDUAL)
    assert residual.bac == pytest.approx(1000)
    assert rep.total_bac == pytest.approx(1500)
    assert any("costo propio" in i.message for i in rep.issues)


def test_summary_baseline_not_rolled_up_uses_children():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    tasks = [_t(1, 1, "Fase", 100, s, f, wbs="1"), _t(2, 2, "A", 300, s, f, wbs="1.1")]
    rep = build_report(_project(tasks), Config(), LinearPVSource())
    assert rep.total_bac == pytest.approx(300)
    assert any("menor que la suma" in i.message for i in rep.issues)


def test_pv_before_horizon_goes_to_month_1_and_after_is_reported():
    tasks = [_t(1, 1, "A", 1200, datetime(2025, 11, 3, 8), datetime(2026, 2, 27, 17), wbs="1")]
    cfg = Config(start_month=date(2026, 1, 1), months=1)
    rep = build_report(_project(tasks), cfg, LinearPVSource())
    row = rep.rows[0]
    assert row.vector[0] == pytest.approx(row.pv_before + row.pv[(2026, 1)])
    assert rep.pv_after > 0
    assert any("posterior" in i.message for i in rep.issues)


def test_inactive_tasks_excluded():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    tasks = [_t(1, 1, "A", 100, s, f, wbs="1"), _t(2, 1, "B", 50, s, f, wbs="2", active=False)]
    rep = build_report(_project(tasks), Config(), LinearPVSource())
    assert [r.name for r in rep.rows] == ["A"]
    rep2 = build_report(_project(tasks), Config(include_inactive=True), LinearPVSource())
    assert len(rep2.rows) == 2


def test_task_without_baseline_is_reported():
    s, f = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    t_new = _t(2, 1, "Nueva", 0, None, None, wbs="2")
    t_new.cost = 999
    tasks = [_t(1, 1, "A", 100, s, f, wbs="1"), t_new]
    rep = build_report(_project(tasks), Config(), LinearPVSource())
    assert any("no tienen línea base" in i.message for i in rep.issues)


def test_xml_without_baseline_uses_current_schedule(tmp_path):
    xml = tmp_path / "sin_lb.xml"
    xml.write_text("""<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project"><Title>Sin LB</Title><Tasks>
<Task><UID>0</UID><ID>0</ID><Name>Sin LB</Name><OutlineLevel>0</OutlineLevel><Cost>10000</Cost></Task>
<Task><UID>1</UID><ID>1</ID><Name>A</Name><WBS>1</WBS><OutlineLevel>1</OutlineLevel>
<Start>2026-01-05T08:00:00</Start><Finish>2026-01-30T17:00:00</Finish><Cost>10000</Cost></Task>
<Task><UID>2</UID><ID>2</ID><IsNull>1</IsNull><OutlineLevel>0</OutlineLevel></Task>
</Tasks></Project>""", encoding="utf-8")
    p = read_xml(xml, Config())
    assert not p.baseline_found
    rep = build_report(p, Config(), xml_source(Config()))
    assert rep.total_bac == pytest.approx(100.0)
    assert any("no tiene guardada la línea base" in i.message for i in rep.issues)
