"""Prueba del lector COM con objetos simulados (no requiere MS Project).

Los objetos imitan el comportamiento documentado de la API de MS Project:
fechas como datetime con zona UTC (pywin32), 'NA' en fechas vacías, filas en
blanco como None y TimeScaleData con valores '' en meses sin datos.
"""
from datetime import date, datetime, timezone

import pytest

from evm.config import Config
from evm.reader_com import PJ_TIMESCALE_MONTHS, TS_BASELINE_COST, ComSession
from evm.timeline import add_months
from evm.wbs import KIND_LEAF, build_report

pytestmark = pytest.mark.modo("costo")

UTC = timezone.utc


def d(y, m, day, h=8):
    return datetime(y, m, day, h, 0, tzinfo=UTC)


class FakeItem:
    def __init__(self, start, value):
        self.StartDate = start
        self.Value = value


class FakeTSV:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)

    def Item(self, i):
        return self._items[i - 1]


class FakeTask:
    def __init__(self, uid, level, name, wbs, monthly=None, bl=True, **kw):
        self.UniqueID = uid
        self.ID = uid
        self.Name = name
        self.OutlineLevel = level
        self.OutlineNumber = wbs
        self.WBS = wbs
        self.Summary = kw.get("summary", False)
        self.Milestone = kw.get("milestone", False)
        self.Active = True
        self.ExternalTask = False
        self.monthly = monthly or {}
        total = sum(self.monthly.values())
        self.BaselineCost = total if bl else 0.0
        keys = sorted(self.monthly) or [(2026, 1)]
        self.BaselineStart = d(*keys[0], 5) if bl else "NA"
        self.BaselineFinish = d(keys[-1][0], keys[-1][1], 20, 17) if bl else "NA"
        self.Start = d(*keys[0], 5)
        self.Finish = d(keys[-1][0], keys[-1][1], 20, 17)
        self.Cost = total
        self.FixedCost = 0.0
        self.FixedCostAccrual = 3
        self.PercentComplete = kw.get("pct", 0)
        self.PhysicalPercentComplete = 0
        self.ActualStart = "NA"
        self.ActualFinish = "NA"
        self.ActualCost = kw.get("ac", 0.0)
        self.ACWP = kw.get("ac", 0.0)
        self.BCWP = 0.0
        self.Text10 = kw.get("method", "")
        self.Number1 = kw.get("weight", 0.0)
        self.Text5 = kw.get("visible", "")
        self.calls = []

    def GetField(self, fid):
        return ""

    def TimeScaleData(self, start, end, kind, unit, count):
        assert start.tzinfo is not None and end.tzinfo is not None    # fechas con zona explícita
        assert unit == PJ_TIMESCALE_MONTHS and count == 1
        self.calls.append(kind)
        items = []
        cur = date(start.year, start.month, 1)
        while cur <= date(end.year, end.month, 1):
            v = self.monthly.get((cur.year, cur.month), "")
            items.append(FakeItem(datetime(cur.year, cur.month, 1, tzinfo=UTC), v))
            cur = add_months(cur, 1)
        return FakeTSV(items)


class FakeProject:
    def __init__(self, tasks, saved=True):
        self.Tasks = tasks
        self.FullName = r"C:\proyectos\programa.mpp"
        self.Name = "programa.mpp"
        self.CurrencyCode = "CLP"
        self.StatusDate = "NA"
        self.ProjectStart = d(2026, 1, 5)
        self._saved = saved
        total = sum(t.BaselineCost for t in tasks if t is not None and t.OutlineLevel == 1)
        self.ProjectSummaryTask = type("S", (), {"Name": "Programa Simulado", "BaselineCost": total,
                                                "Cost": total})()

    def BaselineSavedDate(self, n):
        return d(2025, 12, 15) if self._saved else "NA"


class FakeApp:
    def CustomFieldGetName(self, fid):
        return {188743750: "Método EVM", 188743767: "Peso hito", 188743743: "Visible"}.get(fid, "")

    def FieldNameToFieldConstant(self, name, kind):
        raise RuntimeError("no existe")


def make_session(tasks, saved=True):
    s = ComSession("programa.mpp")
    s.app = FakeApp()
    s.project = FakeProject(tasks, saved)
    return s


def program():
    wp1 = FakeTask(3, 3, "Paquete 1", "1.1.1", {(2026, 1): 100.0, (2026, 2): 300.0}, method="50/50")
    wp2 = FakeTask(4, 3, "Paquete 2", "1.1.2", {(2026, 3): 600.0}, method="0/100")
    sub = FakeTask(5, 4, "Actividad", "1.1.2.1", {(2026, 3): 600.0})
    grp = FakeTask(2, 2, "Grupo", "1.1", {(2026, 1): 100.0, (2026, 2): 300.0, (2026, 3): 600.0}, summary=True)
    top = FakeTask(1, 1, "Fase", "1", dict(grp.monthly), summary=True)
    return [top, None, grp, wp1, wp2, sub]


def test_read_project_and_monthly_pv():
    s = make_session(program())
    cfg = Config(method_field="Método EVM", milestone_weight_field="Peso hito")
    p = s.read_project(cfg)
    assert p.name == "Programa Simulado" and p.backend == "com" and p.baseline_found
    assert len(p.tasks) == 5                                   # la fila en blanco (None) se omite
    t3 = next(t for t in p.tasks if t.uid == 3)
    assert t3.method == "50/50"                                # alias resuelto con CustomFieldGetName
    assert t3.baseline_start == datetime(2026, 1, 5, 8)        # sin zona horaria
    assert t3.actual_start is None                             # 'NA' -> None
    src = s.pv_source(p, cfg)
    rep = build_report(p, cfg, src)
    leaf_rows = [r for r in rep.rows if r.kind == KIND_LEAF]
    assert [r.code for r in leaf_rows] == ["1.1.1", "1.1.2"]
    assert rep.total_bac == pytest.approx(1000)
    assert rep.rows[0].vector[:3] == pytest.approx([100, 300, 600])
    assert src.timephased_count == 2 and src.linear_count == 0
    assert TS_BASELINE_COST[0] in s.project.Tasks[3].calls


def test_visible_field_by_alias():
    tasks = program()
    tasks[-1].Text5 = "Sí"                                     # 1.1.2.1, de nivel mayor que el máximo
    tasks[3].Text5 = "No"                                      # 1.1.1
    s = make_session(tasks)
    p = s.read_project(Config())
    assert {t.uid: t.visible for t in p.tasks}[5] == "Sí" and not p.warnings
    rep = build_report(p, Config(), s.pv_source(p, Config()))
    assert [r.code for r in rep.rows if r.kind == KIND_LEAF] == ["1.1.2.1"]
    assert rep.total_bac == pytest.approx(1000)


def test_without_baseline_uses_current_cost():
    tasks = [FakeTask(1, 1, "A", "1", {(2026, 1): 50.0}, bl=False)]
    s = make_session(tasks, saved=False)
    p = s.read_project(Config())
    assert not p.baseline_found
    assert p.tasks[0].baseline_cost == 50.0
    rep = build_report(p, Config(), s.pv_source(p, Config()))
    assert rep.total_bac == pytest.approx(50.0)
    assert 5 in tasks[0].calls                                 # pjTaskTimescaledCost


def test_baseline_n_uses_matching_constant():
    tasks = program()
    for t in tasks:
        if t is not None:
            t.Baseline2Cost, t.Baseline2Start, t.Baseline2Finish = t.BaselineCost, t.BaselineStart, t.BaselineFinish
    s = make_session(tasks)
    cfg = Config(baseline=2)
    p = s.read_project(cfg)
    assert p.baseline_label == "Línea base 2"
    build_report(p, cfg, s.pv_source(p, cfg))
    assert TS_BASELINE_COST[2] in tasks[3].calls
