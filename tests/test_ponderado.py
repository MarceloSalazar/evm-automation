"""Método «Ponderado»: cada tarea de la fila con su peso y su propio método (percent.weighted_row, desde la 2.7.0)."""
from datetime import datetime

import pytest

from evm.config import Config
from evm.methods import WEIGHTED, parse_method
from evm.update import run_update
from evm.wbs import build_report
from evm.workbook import write_workbook

from builders import P, T, rows_by_code, set_bac

JAN = (datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17))
FEB = (datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17))
MAR = (datetime(2026, 3, 2, 8), datetime(2026, 3, 27, 17))
APR = datetime(2026, 4, 30, 17)


def cfg() -> Config:
    c = Config(max_level=2, months=12, milestone_weight_field="Peso")
    c.validate()
    return c


def ejemplo(progress: float = 60.0):
    """Ejemplo: especificación 70 (% completado), preparación 0 (0/100), V2 hito sin peso y aprobación hito 30."""
    return [T(1, 1, "1", "Fase", bs=JAN[0], bf=APR),
            T(2, 2, "1.1", "Paquete", bs=JAN[0], bf=APR, method="Ponderado"),
            T(3, 3, "", "Especificación", bs=JAN[0], bf=FEB[1], start=JAN[0], finish=FEB[1], method="% Completado",
              milestone_weight=70, percent_complete=progress, actual_start=JAN[0]),
            T(4, 3, "", "Preparación", bs=MAR[0], bf=MAR[1], start=MAR[0], finish=MAR[1], method="0/100",
              milestone_weight=0),
            T(5, 3, "", "V2", bs=MAR[1], bf=MAR[1], start=MAR[1], finish=MAR[1], milestone=True, duration=0.0),
            T(6, 3, "", "Aprobación", bs=APR, bf=APR, start=APR, finish=APR, milestone=True, duration=0.0,
              milestone_weight=30)]


def test_parse_ponderado():
    assert parse_method("Ponderado").kind == WEIGHTED and parse_method("Hitos ponderados").kind != WEIGHTED


def test_pv_by_weights_and_checks():
    rep = build_report(P(ejemplo()), cfg())
    row = rows_by_code(rep)["1.1"]
    pv = row.pct_vec[""]
    assert pv[0] + pv[1] == pytest.approx(0.70) and pv[2] == pytest.approx(0.0) and pv[3] == pytest.approx(0.30)
    ids = [(i.id, i.name) for i in rep.issues if i.id.startswith("HIT-01")]
    assert ("HIT-010", "V2 (Id 5)") in ids                        # hito sin peso: error
    assert ("HIT-011", "Preparación (Id 4)") in ids               # tarea sin peso: aviso
    assert not any(i.id == "HIT-001" for i in rep.issues)        # el peso de una tarea no es un error en «Ponderado»
    assert not any(i.id == "HIT-013" for i in rep.issues)        # 70 + 30 = 100


def test_ev_is_the_sum_of_each_part(tmp_path):
    c = cfg()
    book = tmp_path / "libro.xlsx"
    write_workbook(build_report(P(ejemplo()), c), c, book)
    set_bac(book)
    res = run_update(P(ejemplo(60.0), status_date=datetime(2026, 3, 31)), c, book, backup=False, dry_run=True)
    row = next(r for r in res.rows if r.code == "1.1")
    assert row.ev_cum == pytest.approx(0.42 * 1_000_000, abs=1)         # 70 × 60 %; el hito de 30 sin cumplir
    tasks = ejemplo(100.0)
    tasks[5].percent_complete, tasks[5].actual_finish = 100.0, APR
    res = run_update(P(tasks, status_date=datetime(2026, 4, 30)), c, book, backup=False, dry_run=True)
    assert next(r for r in res.rows if r.code == "1.1").ev_cum == pytest.approx(1_000_000, abs=1)


def test_groups_split_by_duration_or_relative_weights():
    def tasks(wx=0.0, wy=0.0):
        return [T(1, 1, "1", "Fase", bs=JAN[0], bf=APR),
                T(2, 2, "1.1", "Paquete", bs=JAN[0], bf=APR, method="Ponderado"),
                T(3, 3, "", "Ingeniería", bs=JAN[0], bf=MAR[1], milestone_weight=70),
                T(4, 4, "", "X", bs=JAN[0], bf=FEB[1], start=JAN[0], finish=FEB[1], milestone_weight=wx),
                T(5, 4, "", "Y", bs=MAR[0], bf=MAR[1], start=MAR[0], finish=MAR[1], milestone_weight=wy),
                T(6, 3, "", "Aprobación", bs=APR, bf=APR, start=APR, finish=APR, milestone=True, duration=0.0,
                  milestone_weight=30)]
    by_duration = rows_by_code(build_report(P(tasks()), cfg()))["1.1"].pct_vec[""]
    assert by_duration[0] + by_duration[1] == pytest.approx(0.70 * 2 / 3, abs=0.01)     # X: 8 semanas de 12
    assert by_duration[2] == pytest.approx(0.70 / 3, abs=0.01)
    relative = rows_by_code(build_report(P(tasks(3, 1)), cfg()))["1.1"].pct_vec[""]
    assert relative[0] + relative[1] == pytest.approx(0.525) and relative[2] == pytest.approx(0.175)


def test_rework_without_baseline_counts_only_in_ev():
    from evm import percent as pct
    tasks = ejemplo(100.0)
    tasks[5].percent_complete, tasks[5].actual_finish = 100.0, APR
    tasks.append(T(7, 3, "", "Retrabajo", bs=None, bf=None, start=APR, finish=APR, milestone_weight=20))
    rep = build_report(P(tasks), cfg())
    row = rows_by_code(rep)["1.1"]
    assert sum(row.pct_vec[""]) == pytest.approx(1.0) and row.pct_vec[""][3] == pytest.approx(0.30)   # PV igual
    assert pct.weighted_earned(row.weighted[""], None) == pytest.approx(100 / 120)


def test_without_weights_falls_back():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="Ponderado"), T(3, 3, "", "A"), T(4, 3, "", "B")]
    rep = build_report(P(tasks), cfg())
    assert any(i.id == "HIT-012" for i in rep.issues)
    assert sum(rows_by_code(rep)["1.1"].pct_vec[""]) == pytest.approx(1.0)
