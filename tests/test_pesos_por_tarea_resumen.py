"""Reparto de una fila entre sus tareas (percent.RowTree, desde la 2.7.0): el mismo en todos los métodos; el peso de
una tarea es un % de la parte de su tarea resumen y no se propaga más arriba."""
from datetime import datetime

import pytest

from evm.config import Config
from evm.methods import parse_method
from evm.percent import earned_fraction
from evm.update import run_update
from evm.wbs import build_report
from evm.workbook import write_workbook

from builders import P, T, rows_by_code, set_bac

JAN = (datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17))
FEB = (datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17))
MAR = (datetime(2026, 3, 16, 17), datetime(2026, 3, 16, 17))       # fecha de un hito
FEB_M = (datetime(2026, 2, 16, 17), datetime(2026, 2, 16, 17))
NONE = (None, None)

METHODS = ["Ponderado", "% Completado", "% Físico completado", "Hitos ponderados", "50/50", "Avance lineal", ""]


def cfg(**kw) -> Config:
    c = Config(max_level=2, months=12, milestone_weight_field="Peso", **kw)
    c.validate()
    return c


def task(uid, level, code, name, period=NONE, weight=None, done=0.0, milestone=False, baseline=True, **kw):
    """Tarea de detalle (o hito) con fechas actuales `period` y la misma línea base, salvo baseline=False."""
    return T(uid, level, code, name, bs=period[0] if baseline else None, bf=period[1] if baseline else None,
             start=period[0], finish=period[1], milestone_weight=weight, percent_complete=done,
             physical_percent_complete=done, milestone=milestone, duration=0.0 if milestone else None,
             actual_start=period[0] if done > 0 else None, actual_finish=period[1] if done >= 100 else None, **kw)


def summary(uid, level, code, name, weight=None, period=NONE, **kw):
    """Tarea resumen, sin línea base propia salvo period."""
    return T(uid, level, code, name, bs=period[0], bf=period[1], start=period[0], finish=period[1],
             milestone_weight=weight, **kw)


def deep_milestone(method, weight=10):
    """Fila 1.2: tarea 1.2.1 (enero, terminada) y tarea resumen 1.2.3 con la tarea 1.2.3.1 (febrero, al 50 %) y el
    hito 1.2.3.4 (marzo, sin cumplir) con peso."""
    return [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method=method),
            task(3, 3, "1.2.1", "Tarea A", JAN, done=100),
            summary(4, 3, "1.2.3", "Resumen"), task(5, 4, "1.2.3.1", "Tarea B", FEB, done=50),
            task(6, 4, "1.2.3.4", "Hito", MAR, weight, milestone=True)]


def report(tasks, **kw):
    return build_report(P(tasks), cfg(**kw))


def row_of(rep, code="1.2"):
    return rows_by_code(rep)[code]


def ids(rep, code):
    return [i for i in rep.issues if i.id == code]


def ev_of(row, cut=(2026, 2)):
    role = next(iter(row.weights))
    return earned_fraction(parse_method(row.method), row.weights[role], sum(row.pct_vec[role][:2]), cut)[0]


@pytest.mark.parametrize("method", METHODS)
def test_a_weight_does_not_go_above_its_summary_task(method):
    """El hito de peso 10 en 1.2.3.4 vale el 10 % de 1.2.3 (5 % de la fila), con cualquier método."""
    row = row_of(report(deep_milestone(method)))
    assert row.pct_vec[""][:3] == pytest.approx([0.50, 0.45, 0.05])
    assert row.weights[""].weights == pytest.approx({3: 0.50, 5: 0.45, 6: 0.05})


def test_a_direct_milestone_weight_is_a_share_of_the_row():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, done=100), task(5, 3, "1.2.2", "Tarea B", FEB, done=50),
             task(6, 3, "1.2.3", "Hito", MAR, 10, milestone=True)]
    row = row_of(report(tasks))
    assert row.pct_vec[""][:3] == pytest.approx([0.45, 0.45, 0.10])
    assert ev_of(row) == pytest.approx(0.45 + 0.45 * 0.5)


@pytest.mark.parametrize("method", ["% Completado", "Hitos ponderados", "Ponderado"])
def test_task_and_summary_weights_count_in_every_method(method):
    """Peso 30 en la tarea 1.2.1: la tarea resumen 1.2.3 recibe el 70 % restante (por duración)."""
    tasks = deep_milestone(method, None)
    tasks[2].milestone_weight = 30
    rep = report(tasks)
    assert row_of(rep).pct_vec[""][:2] == pytest.approx([0.30, 0.70])
    assert not ids(rep, "HIT-001")
    tasks = deep_milestone(method, None)
    tasks[3].milestone_weight = 70                      # el peso en la tarea resumen
    assert row_of(report(tasks)).pct_vec[""][:2] == pytest.approx([0.30, 0.70])


def test_weight_on_the_row_task_is_not_used():
    tasks = deep_milestone("% Completado")
    tasks[1].milestone_weight = 40
    [issue] = ids(report(tasks), "HIT-001")
    assert "tarea de una fila" in issue.message


def test_milestones_summing_100_leave_the_task_out_with_a_warning_in_every_method():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, done=100),
             task(4, 3, "1.2.2", "Hito 1", FEB_M, 60, milestone=True, done=100),
             task(5, 3, "1.2.3", "Hito 2", MAR, 40, milestone=True)]
    rep = report(tasks)
    assert row_of(rep).pct[""] == pytest.approx({(2026, 2): 0.60, (2026, 3): 0.40})     # la tarea A, nada
    assert [i.name for i in ids(rep, "HIT-011")] == ["Tarea A"]


def test_weights_summing_100_in_different_summaries_are_not_reported():
    """Cada tarea resumen con sus hitos 50/50: no es un error de pesos (antes, HIT-007 «suman 200»)."""
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             summary(3, 3, "1.2.1", "Suministro 1", 50), task(4, 4, "1.2.1.1", "Pedido 1", JAN, 50, milestone=True),
             task(5, 4, "1.2.1.2", "Entrega 1", FEB_M, 50, milestone=True),
             summary(6, 3, "1.2.2", "Suministro 2", 50), task(7, 4, "1.2.2.1", "Pedido 2", FEB_M, 50, milestone=True),
             task(8, 4, "1.2.2.2", "Entrega 2", MAR, 50, milestone=True)]
    rep = report(tasks)
    assert row_of(rep).pct_vec[""][:3] == pytest.approx([0.25, 0.50, 0.25])
    assert not [i for i in rep.issues if i.id.startswith("HIT-")]


def test_weights_over_100_and_unused_weights_are_errors_in_every_method():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             task(3, 3, "1.2.1", "A", JAN, 70), task(4, 3, "1.2.2", "B", FEB, 50),
             summary(5, 3, "1.2.3", "Sin peso"), task(6, 4, "1.2.3.1", "Con peso", JAN, 100)]
    rep = report(tasks)
    [over] = ids(rep, "HIT-013")
    assert over.severity == "Error" and "suman 120" in over.message
    [lost] = ids(rep, "HIT-014")
    assert "Sin peso" in lost.name


def test_milestone_without_weight_is_an_error_only_in_ponderado():
    tasks = deep_milestone("Ponderado", None)
    tasks[2].milestone_weight = 30
    assert ids(report(tasks), "HIT-010")
    tasks = deep_milestone("% Completado", None)
    tasks[2].milestone_weight = 30
    assert not ids(report(tasks), "HIT-010") and not ids(report(tasks), "HIT-011")


def test_weighted_milestones_without_any_weight_split_pv_and_ev_equally_among_milestones():
    tasks = deep_milestone("Hitos ponderados", None)
    tasks.append(task(7, 3, "1.2.4", "Hito 2", FEB_M, milestone=True, done=100))
    rep = report(tasks)
    row = row_of(rep)
    assert row.weights[""].weights == pytest.approx({6: 0.5, 7: 0.5})
    assert sum(row.pct_vec[""]) == pytest.approx(1.0)
    assert ev_of(row) == pytest.approx(0.5)                 # el hito de febrero cumplido; las tareas no cuentan
    [issue] = ids(rep, "HIT-003")
    assert issue.severity == "Advertencia"


def test_ponderado_uses_the_block_of_a_task_split_after_the_baseline():
    """1.2.3 tiene línea base y peso 60; sus subtareas se agregaron después, sin línea base."""
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="Ponderado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, 40, done=100),
             summary(4, 3, "1.2.3", "Bloque", 60, period=(FEB[0], MAR[1])),
             task(5, 4, "1.2.3.1", "Tarea B", FEB, done=50, baseline=False),
             task(6, 4, "1.2.3.2", "Tarea C", (MAR[0], MAR[1]), baseline=False)]
    row = row_of(report(tasks))
    assert row.pct_vec[""][0] == pytest.approx(0.40) and sum(row.pct_vec[""][1:]) == pytest.approx(0.60)


def test_ponderado_follows_pv_current_dates():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="Ponderado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, 50), task(4, 3, "1.2.2", "Nueva", FEB, 50, baseline=False)]
    assert row_of(report(tasks)).pct_vec[""][:2] == pytest.approx([1.0, 0.0])
    assert row_of(report(tasks, pv_current_dates=True)).pct_vec[""][:2] == pytest.approx([0.5, 0.5])


def test_ponderado_without_weights_measures_each_task_with_its_own_method():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="Ponderado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, done=100), task(4, 3, "1.2.2", "Tarea B", FEB, done=50, method="0/100")]
    rep = report(tasks)
    assert ev_of(row_of(rep)) == pytest.approx(0.5)        # B, 0/100 sin terminar: 0
    assert ids(rep, "HIT-012") and not ids(rep, "HIT-010")


def test_a_task_without_baseline_leaves_its_share_to_its_siblings():
    """Reparto local: 1.2.3 conserva su 40 % del PV; dentro, el hito con línea base recibe la parte de la tarea
    nueva. En el % ganado cuentan todas."""
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, 60, done=100),
             summary(4, 3, "1.2.3", "Resumen", 40), task(5, 4, "1.2.3.1", "Nueva", FEB, 50, baseline=False),
             task(6, 4, "1.2.3.4", "Hito", MAR, 50, milestone=True)]
    rep = report(tasks)
    row = row_of(rep)
    assert row.pct_vec[""][:3] == pytest.approx([0.60, 0.0, 0.40])
    assert row.weights[""].weights == pytest.approx({3: 0.60, 5: 0.20, 6: 0.20})
    assert not ids(rep, "HIT-017")


def test_a_summary_without_weight_and_without_planned_duration_is_reported():
    """1.2.3 sin peso: en el PV su duración planificada es 0 (solo el hito tiene línea base)."""
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method="% Completado"),
             task(3, 3, "1.2.1", "Tarea A", JAN, done=100),
             summary(4, 3, "1.2.3", "Resumen"), task(5, 4, "1.2.3.1", "Nueva", FEB, 50, baseline=False),
             task(6, 4, "1.2.3.4", "Hito", MAR, 50, milestone=True)]
    rep = report(tasks)
    assert row_of(rep).pct_vec[""][0] == pytest.approx(1.0)
    [issue] = ids(rep, "HIT-017")
    assert "Resumen" in issue.name


def test_residual_row_weights_stay_in_their_summary():
    tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.1", "Paquete"), task(3, 3, "1.1.1", "X", JAN),
             task(4, 2, "", "Suelta", JAN, done=100), summary(5, 2, "", "Bloque"), task(6, 3, "", "Tarea B", FEB),
             task(7, 3, "", "Hito", MAR, 10, milestone=True)]
    row = row_of(report(tasks), "1.0-D")
    assert row.pct_vec[""][:3] == pytest.approx([0.50, 0.45, 0.05])


def test_multi_company_weight_is_a_share_of_each_company():
    """Un hito de peso 10 de la empresa A vale el 10 % del BAC de A (también con «Ponderado»)."""
    for method in ("Ponderado", "% Completado"):
        tasks = [summary(1, 1, "1", "Fase"), summary(2, 2, "1.2", "Paquete", method=method),
                 task(3, 3, "1.2.1", "Tarea A", JAN, company="EmpresaA"),
                 task(4, 3, "1.2.2", "Tarea B", FEB, company="EmpresaB"),
                 task(5, 3, "1.2.3", "Hito", MAR, 10, milestone=True, company="EmpresaA")]
        row = row_of(report(tasks, multi_enabled=True))
        assert row.pct_vec["A"][:3] == pytest.approx([0.90, 0.0, 0.10]), method
        assert row.pct_vec["B"][:3] == pytest.approx([0.0, 1.0, 0.0]), method


def test_update_uses_the_same_split_as_generate(tmp_path):
    """actualizar recalcula el % planificado con el mismo reparto: sin cambios en MS Project no hay diferencias con
    la línea base del libro, y el % ganado es el de generar."""
    c = cfg()
    tasks = deep_milestone("% Completado")
    book = tmp_path / "libro.xlsx"
    write_workbook(build_report(P(tasks), c), c, book)
    set_bac(book)
    res = run_update(P(deep_milestone("% Completado"), status_date=datetime(2026, 2, 28)), c, book, backup=False,
                     dry_run=True)
    row = next(r for r in res.rows if r.code == "1.2")
    assert row.ev_cum == pytest.approx((0.50 + 0.45 * 0.5) * 1_000_000, abs=1)
    assert not [i for i in res.warnings + row.notes if i.id == "ACT-047"]
