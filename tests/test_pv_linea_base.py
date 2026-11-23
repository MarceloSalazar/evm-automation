"""% planificado solo con las tareas con línea base (wbs.pv_selection, desde la 2.7.0) y generar --pv-fechas-actuales."""
from datetime import datetime

import pytest
from openpyxl import load_workbook

from evm.config import Config
from evm.update import run_update
from evm.wbs import build_report
from evm.workbook import read_meta, write_workbook

from builders import P, T, rows_by_code, set_bac

JAN, FEB, MAR = (datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)), (datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17)), \
    (datetime(2026, 3, 2, 8), datetime(2026, 3, 27, 17))
APR = (datetime(2026, 4, 1, 8), datetime(2026, 4, 30, 17))


def cfg(**kw) -> Config:
    c = Config(max_level=2, months=12, **kw)
    c.validate()
    return c


def tasks():
    """1.1: dos tareas con línea base y un retrabajo sin línea base; 1.2: una tarea con línea base desglosada después
    en subtareas sin línea base; 1.3: fila sin ninguna línea base."""
    def none(uid, level, code, name, dates):
        return T(uid, level, code, name, bs=None, bf=None, start=dates[0], finish=dates[1])
    return [T(1, 1, "1", "Fase", bs=JAN[0], bf=APR[1]),
            T(2, 2, "1.1", "Paquete", bs=JAN[0], bf=FEB[1]),
            T(3, 3, "", "A", bs=JAN[0], bf=JAN[1], start=JAN[0], finish=JAN[1]),
            T(4, 3, "", "B", bs=FEB[0], bf=FEB[1], start=FEB[0], finish=FEB[1]),
            none(5, 3, "", "Retrabajo", (MAR[0], APR[1])),
            T(6, 2, "1.2", "Desglosado", bs=JAN[0], bf=MAR[1]),
            none(7, 3, "", "D", (JAN[0], FEB[1])),
            none(8, 3, "", "E", (FEB[0], APR[1])),
            none(9, 2, "1.3", "Sin línea base", (MAR[0], APR[1]))]


def test_pv_only_with_baselined_work():
    rep = build_report(P(tasks()), cfg())
    rows = rows_by_code(rep)
    r1, r2, r3 = rows["1.1"], rows["1.2"], rows["1.3"]
    # Retrabajo: fuera del PV (fechas Plan de A y B), dentro del % ganado
    assert [t.name for t in r1.pv_leaves] == ["A", "B"] and [t.name for t in r1.pv_out] == ["Retrabajo"]
    assert (r1.start, r1.finish) == (JAN[0], FEB[1])
    assert sum(r1.pct_vec[""][:2]) == pytest.approx(1.0)
    assert [tt.task.name for tt in r1.weights[""].members] == ["A", "B", "Retrabajo"]
    # Desglose posterior: el PV usa la línea base de la tarea desglosada, como un bloque
    assert [t.name for t in r2.pv_leaves] == ["Desglosado"] and not r2.pv_out
    assert (r2.start, r2.finish) == (JAN[0], MAR[1])
    assert [tt.task.name for tt in r2.weights[""].members] == ["D", "E"]
    # Sin ninguna línea base en la fila: sin % planificado ni fechas del plan (su BAC debe ser 0)
    assert r3.pv_leaves == [] and (r3.start, r3.finish) == (None, None) and not any(r3.pct_vec[""])
    assert [tt.task.name for tt in r3.weights[""].members] == ["Sin línea base"]
    ids = {i.id for i in rep.issues}
    assert "FEC-019" in ids and "FEC-010" not in ids and "HIT-005" not in ids
    fec019 = [i for i in rep.issues if i.id == "FEC-019"]
    assert fec019[0].message.startswith("2 tareas sin línea base no forman parte del % planificado de su fila "
                                        "(1.1, 1.3)")
    assert fec019[0].message.endswith("1 fila no tiene ninguna tarea con línea base (1.3): sin % planificado; su BAC "
                                      "debe ser 0.")
    assert fec019[1].name.startswith("Retrabajo")
    # La fila resumen no cuenta las filas sin % planificado
    assert (rows["1.0"].start, rows["1.0"].finish) == (JAN[0], MAR[1])


def test_pv_fechas_actuales_keeps_the_previous_criterion():
    rep = build_report(P(tasks()), cfg(pv_current_dates=True))
    rows = rows_by_code(rep)
    assert not rows["1.1"].pv_leaves and rows["1.1"].finish == APR[1]          # el retrabajo, con sus fechas actuales
    assert (rows["1.2"].start, rows["1.2"].finish) == (JAN[0], APR[1])
    assert not any(i.id == "FEC-019" for i in rep.issues)


def test_adding_a_subtask_does_not_change_the_pv():
    """El PV de una fila no cambia al agregar una subtarea sin línea base, también al regenerar el libro."""
    before = [t for t in tasks() if t.name != "Retrabajo"]
    pv_before = rows_by_code(build_report(P(before), cfg()))["1.1"].pct_vec[""]
    pv_after = rows_by_code(build_report(P(tasks()), cfg()))["1.1"].pct_vec[""]
    assert pv_after == pytest.approx(pv_before)
    old = rows_by_code(build_report(P(tasks()), cfg(pv_current_dates=True)))["1.1"].pct_vec[""]
    assert old != pytest.approx(pv_before)


def test_book_records_the_criterion_and_actualizar_compares_with_it(tmp_path):
    c = cfg()
    book = tmp_path / "libro.xlsx"
    base = [t for t in tasks() if t.name != "Retrabajo"]
    write_workbook(build_report(P(base), c), c, book)
    set_bac(book)
    assert read_meta(load_workbook(book))["pv_fechas_actuales"] is False
    # Retrabajo sin línea base: el % ganado lo incluye, el PV no cambia (sin ACT-047)
    res = run_update(P(tasks(), status_date=datetime(2026, 3, 31)), c, book, backup=False, dry_run=True)
    assert not any(w.id == "ACT-047" for w in res.warnings)
    # Una línea base guardada de nuevo sí cambia el PV: ACT-047
    moved = tasks()
    moved[3].baseline_start, moved[3].baseline_finish = MAR
    res = run_update(P(moved, status_date=datetime(2026, 3, 31)), c, book, backup=False, dry_run=True)
    assert any(w.id == "ACT-047" and "1.1" in w.message for w in res.warnings)
