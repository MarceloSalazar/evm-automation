"""Método «Ponderado»: cada tarea de la fila con su peso y su propio método (percent.RowTree, desde la 2.7.0)."""
from datetime import datetime

import pytest

from evm.config import Config
from evm.messages import SEV_ERROR, SEV_WARN
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
              milestone_weight=None),
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
    def tasks(wx=None, wy=None):
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
    assert pct.earned_fraction(parse_method(row.method), row.weights[""], 0.0)[0] == pytest.approx(100 / 120)


def test_without_weights_falls_back():
    tasks = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "Paquete", method="Ponderado"), T(3, 3, "", "A"), T(4, 3, "", "B")]
    rep = build_report(P(tasks), cfg())
    assert any(i.id == "HIT-012" for i in rep.issues)
    assert sum(rows_by_code(rep)["1.1"].pct_vec[""]) == pytest.approx(1.0)


def test_ev_fill_includes_weighted_rows():
    """generar --ev-rellenar: las filas «Ponderado» también tienen el % ganado de cada mes (de cada empresa)."""
    from evm.history import earned_history
    c = Config(max_level=2, months=12, milestone_weight_field="Peso", multi_enabled=True)
    c.validate()
    tasks = ejemplo(60.0)
    for t in tasks:
        t.company = "EmpresaA"
    rep = build_report(P(tasks, status_date=datetime(2026, 3, 31)), c)
    history = earned_history(rep, c)
    assert history.loaded
    row = rows_by_code(rep)["1.1"]
    values = row.ev_hist["A"]
    assert len(values) == 3 and values[0] > 0 and values[1] > 0               # enero y febrero, en línea recta
    assert sum(values) == pytest.approx(0.42)                                # 70 × 60 %: el acumulado a la fecha de estado


# ----------------------------------------------------------------------------
# Pesos en cada tarea resumen (desde la 2.7.0): % de la tarea resumen; el resto, por duración
# ----------------------------------------------------------------------------
def paquete(*children):
    return [T(1, 1, "1", "Fase", bs=JAN[0], bf=APR), T(2, 2, "1.1", "Paquete", bs=JAN[0], bf=APR, method="Ponderado"),
            *children]


def hito(uid, name, level=3, weight=None, **kw):
    return T(uid, level, "", name, bs=APR, bf=APR, start=APR, finish=APR, milestone=True, duration=0.0,
             milestone_weight=weight, **kw)


def pv_of(tasks):
    rep = build_report(P(tasks), cfg())
    return rep, rows_by_code(rep)["1.1"].pct_vec[""]


def found(rep, code):
    return [i for i in rep.issues if i.id == code]


def test_the_rest_goes_to_the_tasks_without_weight_by_duration():
    """Hito de 10 y dos tareas sin peso (8 y 4 semanas): el 90 restante se reparte 60 y 30."""
    rep, pv = pv_of(paquete(T(3, 3, "", "A", bs=JAN[0], bf=FEB[1]), T(4, 3, "", "B", bs=MAR[0], bf=MAR[1]),
                            hito(5, "Hito", weight=10)))
    assert pv[0] + pv[1] == pytest.approx(0.60) and pv[2] == pytest.approx(0.30) and pv[3] == pytest.approx(0.10)
    assert not found(rep, "HIT-011") and not found(rep, "HIT-013") and not found(rep, "HIT-010")


def test_more_than_100_is_an_error_and_tasks_without_weight_do_not_count():
    rep, pv = pv_of(paquete(T(3, 3, "", "A", bs=JAN[0], bf=FEB[1], milestone_weight=80),
                            T(4, 3, "", "B", bs=MAR[0], bf=MAR[1], milestone_weight=40),
                            T(5, 3, "", "C", bs=MAR[0], bf=MAR[1])))
    assert pv[0] + pv[1] == pytest.approx(2 / 3) and pv[2] == pytest.approx(1 / 3)
    [issue] = found(rep, "HIT-013")
    assert (issue.code, issue.severity) == ("1.1", SEV_ERROR)
    assert issue.message.startswith("Los pesos de las subtareas de la fila suman 120, más de 100")
    assert [i.name for i in found(rep, "HIT-011")] == ["C (Id 5)"]


def test_less_than_100_without_tasks_for_the_rest_is_scaled_with_a_warning():
    rep, pv = pv_of(paquete(T(3, 3, "", "A", bs=JAN[0], bf=FEB[1], milestone_weight=30),
                            T(4, 3, "", "B", bs=MAR[0], bf=MAR[1], milestone_weight=10)))
    assert pv[0] + pv[1] == pytest.approx(0.75) and pv[2] == pytest.approx(0.25)
    [issue] = found(rep, "HIT-013")
    assert issue.severity == SEV_WARN and "suman 40 y no hay subtareas sin peso" in issue.message


def test_weights_inside_a_summary_without_weight_are_of_its_part():
    """Hito de 20; el resumen S (sin peso) recibe el 80 restante y lo reparte con sus propios pesos: X 50 % y el resto
    para Y."""
    rep, pv = pv_of(paquete(hito(3, "Hito", weight=20), T(4, 3, "", "S", bs=JAN[0], bf=MAR[1]),
                            T(5, 4, "", "X", bs=JAN[0], bf=FEB[1], milestone_weight=50),
                            T(6, 4, "", "Y", bs=MAR[0], bf=MAR[1])))
    assert pv[0] + pv[1] == pytest.approx(0.40) and pv[2] == pytest.approx(0.40) and pv[3] == pytest.approx(0.20)
    assert not found(rep, "HIT-013")


def test_summary_without_weight_next_to_weights_that_sum_100_does_not_count():
    """Como un bloque subcontratado sin peso junto a tareas que suman 100: no cuenta (ni sus pesos internos)."""
    from evm import percent as pct
    tasks = paquete(T(3, 3, "", "A", bs=JAN[0], bf=FEB[1], milestone_weight=70, percent_complete=100.0),
                    T(4, 3, "", "B", bs=MAR[0], bf=MAR[1], milestone_weight=30),
                    T(5, 3, "", "Bloque", bs=None, bf=None),
                    T(6, 4, "", "Tarea del bloque", bs=None, bf=None, start=MAR[0], finish=MAR[1],
                      percent_complete=100.0),
                    hito(7, "Hito del bloque", level=4, weight=10, percent_complete=100.0))
    rep = build_report(P(tasks), cfg())
    row = rows_by_code(rep)["1.1"]
    assert pct.earned_fraction(parse_method(row.method), row.weights[""], 0.0)[0] == pytest.approx(0.70)
    [issue] = found(rep, "HIT-014")                              # sus pesos internos no se usan: error
    assert (issue.name, issue.severity) == ("Bloque (Id 5)", SEV_ERROR)
    assert "sus tareas hermanas con peso suman 100 o más" in issue.message and "El peso de su subtarea" in issue.message
    assert not found(rep, "HIT-013") and not found(rep, "HIT-011")


def test_summary_without_weight_with_only_weighted_milestones_gets_nothing_by_duration():
    """Fila sin pesos en su primer nivel: el resumen que solo tiene hitos (duración 0) no recibe nada y sus pesos no
    se usan (HIT-014); una tarea sin peso junto a pesos que suman 100 se informa con HIT-011."""
    rep, pv = pv_of(paquete(T(3, 3, "", "Trabajo", bs=JAN[0], bf=MAR[1]),
                            T(4, 3, "", "Compras", bs=APR, bf=APR),
                            hito(5, "Compra 1", level=4, weight=50), hito(6, "Compra 2", level=4, weight=50),
                            T(7, 3, "", "Grupo", bs=JAN[0], bf=FEB[1]),
                            T(8, 4, "", "Con peso", bs=JAN[0], bf=JAN[1], milestone_weight=100),
                            T(9, 4, "", "Sin peso", bs=FEB[0], bf=FEB[1])))
    assert pv[3] == pytest.approx(0.0) and sum(pv) == pytest.approx(1.0)
    [lost] = found(rep, "HIT-014")
    assert lost.name == "Compras (Id 4)" and "la suya es 0" in lost.message and "sus 2 subtareas" in lost.message
    [unused] = found(rep, "HIT-011")
    assert unused.name == "Sin peso (Id 9)" and "suman 100 o más" in unused.message


def test_milestone_without_weight_gets_nothing_of_the_rest():
    rep, pv = pv_of(paquete(T(3, 3, "", "A", bs=JAN[0], bf=MAR[1]), hito(4, "Con peso", weight=10),
                            hito(5, "Sin peso")))
    assert sum(pv) == pytest.approx(1.0) and pv[3] == pytest.approx(0.10)
    assert [(i.name, i.severity) for i in found(rep, "HIT-010")] == [("Sin peso (Id 5)", SEV_ERROR)]
