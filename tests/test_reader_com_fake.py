"""Prueba del lector COM con objetos simulados (no requiere MS Project).

Los objetos imitan el comportamiento documentado de la API de MS Project:
fechas como datetime con zona UTC (pywin32), 'NA' en fechas vacías y filas en
blanco como None.
"""
from datetime import date, datetime, timezone

import pytest

from evm.config import Config
from evm.reader_com import ComSession
from evm.wbs import KIND_LEAF, build_report


UTC = timezone.utc


def d(y, m, day, h=8):
    return datetime(y, m, day, h, 0, tzinfo=UTC)


class FakeTask:
    def __init__(self, uid, level, name, wbs, months=((2026, 1),), bl=True, **kw):
        """Tarea que dura desde el día 5 del primer mes hasta el día 20 del último mes de `months`."""
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
        keys = sorted(months)
        self.BaselineStart = d(*keys[0], 5) if bl else "NA"
        self.BaselineFinish = d(keys[-1][0], keys[-1][1], 20, 17) if bl else "NA"
        self.Start = d(*keys[0], 5)
        self.Finish = d(keys[-1][0], keys[-1][1], 20, 17)
        self.PercentComplete = kw.get("pct", 0)
        self.PhysicalPercentComplete = 0
        self.ActualStart = "NA"
        self.ActualFinish = "NA"
        self.Text10 = kw.get("method", "")
        self.Number1 = kw.get("weight", 0.0)
        self.Text5 = kw.get("visible", "")

    def GetField(self, fid):
        return ""


class FakeProject:
    def __init__(self, tasks, saved=True):
        self.Tasks = tasks
        self.FullName = r"C:\proyectos\programa.mpp"
        self.Name = "programa.mpp"
        self.CurrencyCode = "CLP"
        self.StatusDate = "NA"
        self.ProjectStart = d(2026, 1, 5)
        self._saved = saved
        self.ProjectSummaryTask = type("S", (), {"Name": "Programa Simulado"})()

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
    wp1 = FakeTask(3, 3, "Paquete 1", "1.1.1", ((2026, 1), (2026, 2)), method="50/50")
    wp2 = FakeTask(4, 3, "Paquete 2", "1.1.2", ((2026, 3),), method="0/100")
    sub = FakeTask(5, 4, "Actividad", "1.1.2.1", ((2026, 3),))
    grp = FakeTask(2, 2, "Grupo", "1.1", ((2026, 1), (2026, 3)), summary=True)
    top = FakeTask(1, 1, "Fase", "1", ((2026, 1), (2026, 3)), summary=True)
    return [top, None, grp, wp1, wp2, sub]


def test_read_project_and_planned_percent():
    s = make_session(program())
    cfg = Config(method_field="Método EVM", milestone_weight_field="Peso hito")
    p = s.read_project(cfg)
    assert p.name == "Programa Simulado" and p.backend == "com" and p.baseline_found
    assert len(p.tasks) == 5                                   # la fila en blanco (None) se omite
    t3 = next(t for t in p.tasks if t.uid == 3)
    assert t3.method == "50/50"                                # alias resuelto con CustomFieldGetName
    assert t3.baseline_start == datetime(2026, 1, 5, 8)        # sin zona horaria
    assert t3.actual_start is None                             # 'NA' -> None
    rep = build_report(p, cfg)
    leaf_rows = [r for r in rep.rows if r.kind == KIND_LEAF]
    assert [r.code for r in leaf_rows] == ["1.1.1", "1.1.2"]
    assert sum(leaf_rows[0].pct_vec[""]) == pytest.approx(1.0)
    assert leaf_rows[1].pct_vec[""][2] == pytest.approx(1.0)  # todo en marzo


def test_visible_field_by_alias():
    tasks = program()
    tasks[-1].Text5 = "Sí"                                     # 1.1.2.1, de nivel mayor que el máximo
    tasks[3].Text5 = "No"                                      # 1.1.1
    s = make_session(tasks)
    p = s.read_project(Config())
    assert {t.uid: t.visible for t in p.tasks}[5] == "Sí" and not p.warnings
    rep = build_report(p, Config())
    assert [r.code for r in rep.rows if r.kind == KIND_LEAF] == ["1.1.2.1"]
    assert not [r.code for r in rep.rows if r.code.endswith("-D")]    # 1.1.1 queda fuera del cálculo


def test_without_baseline_uses_current_dates():
    tasks = [FakeTask(1, 1, "A", "1", ((2026, 2),), bl=False)]
    s = make_session(tasks, saved=False)
    p = s.read_project(Config())
    assert not p.baseline_found
    assert p.tasks[0].baseline_start == datetime(2026, 2, 5, 8)
    rep = build_report(p, Config())
    assert rep.rows[0].pct_vec[""][0] == pytest.approx(1.0) and rep.horizon.start == date(2026, 2, 1)


def test_baseline_n_reads_its_dates():
    tasks = program()
    for t in tasks:
        if t is not None:
            t.Baseline2Start, t.Baseline2Finish = d(2026, 6, 1), d(2026, 6, 26, 17)
    s = make_session(tasks)
    cfg = Config(baseline=2)
    p = s.read_project(cfg)
    assert p.baseline_label == "Línea base 2"
    assert {t.baseline_start for t in p.tasks} == {datetime(2026, 6, 1, 8)}
