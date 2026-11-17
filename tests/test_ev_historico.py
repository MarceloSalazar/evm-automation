"""% ganado de cada mes hasta la fecha de estado al generar el libro (history.earned_history)."""
from datetime import date, datetime

import pytest
from openpyxl import load_workbook

from evm.cli import main
from evm.config import Config
from evm.history import earned_history
from evm.update import read_book, run_update
from evm.wbs import build_report
from evm.workbook import FIRST_MONTH_COL, STATUS_CELL, write_workbook

from builders import P, T, set_bac

STATUS = datetime(2026, 3, 31)
JAN, FEB, MAR = 0, 1, 2


def cfg(**kw) -> Config:
    c = Config(max_level=2, start_month=date(2026, 1, 1), months=12, distribution="calendario", **kw)
    c.validate()
    return c


def tasks():
    """Un paquete por método, con avance a la fecha de estado 31/03/2026."""
    def done(t, s, f):
        t.actual_start, t.actual_finish, t.percent_complete = s, f, 100.0
        return t
    bs, bf = datetime(2026, 1, 1, 8), datetime(2026, 12, 31, 17)
    out = [T(1, 1, "1", "Fase"),
           done(T(2, 2, "1.1", "Cero cien", method="0/100"), datetime(2026, 1, 5, 8), datetime(2026, 2, 13, 17)),
           T(3, 2, "1.2", "Cincuenta", method="50/50", percent_complete=40.0, actual_start=datetime(2026, 2, 2, 8)),
           T(4, 2, "1.3", "Completado", method="% Completado", percent_complete=60.0,
             actual_start=datetime(2026, 1, 1, 0)),
           T(5, 2, "1.4", "Hitos", method="Hitos ponderados"),
           done(T(6, 3, "1.4.1", "H1", milestone=True, milestone_weight=30.0, bs=datetime(2026, 1, 20, 8),
                  bf=datetime(2026, 1, 20, 8)), datetime(2026, 1, 20, 17), datetime(2026, 1, 20, 17)),
           T(7, 3, "1.4.2", "H2", milestone=True, milestone_weight=70.0, bs=datetime(2026, 5, 20, 8),
             bf=datetime(2026, 5, 20, 8)),
           T(8, 2, "1.5", "Gestion", method="Avance lineal", bs=bs, bf=bf, start=bs, finish=bf),
           T(9, 2, "1.6", "Sin comenzar", method="% Completado")]
    return out


def pct_ev(book, sheet="% EV") -> dict[str, list]:
    ws = load_workbook(book)[sheet]
    return {ws.cell(r, 1).value: [ws.cell(r, FIRST_MONTH_COL + i).value for i in range(4)]
            for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "TOTAL")}


def generate(tmp_path, project, c=None):
    c = c or cfg()
    report = build_report(project, c)
    history = earned_history(report, c)
    report.issues.extend(history.issues)
    book = tmp_path / "libro.xlsx"
    write_workbook(report, c, book, history=history)
    return book, report, history


def test_each_method_earns_in_the_months_of_its_actual_dates(tmp_path):
    book, report, history = generate(tmp_path, P(tasks(), status_date=STATUS))
    assert (history.status, history.index) == (date(2026, 3, 31), MAR)
    assert any(i.id == "LIB-016" and "3 meses (01/2026 a 03/2026)" in i.message for i in report.issues)
    ev = pct_ev(book)
    assert ev["1.1"] == pytest.approx([0, 1, 0, None])              # 0/100: el mes del fin real
    assert ev["1.2"] == pytest.approx([0, 0.5, 0, None])            # 50/50: el mes del comienzo real
    assert ev["1.4"] == pytest.approx([0.3, 0, 0, None])            # hito cumplido en enero (peso 30)
    assert ev["1.6"] == pytest.approx([0, 0, 0, None])              # sin avance: 0 (mes cargado)
    # % completado: 60 % repartido en línea recta (días corridos) del comienzo real a la fecha de estado
    jan, feb, mar = ev["1.3"][:3]
    assert jan + feb + mar == pytest.approx(0.6)
    assert jan == pytest.approx(0.6 * 31 / 90, abs=0.002) and feb == pytest.approx(0.6 * 28 / 90, abs=0.002)
    # Avance lineal: el PV de cada mes
    pv = pct_ev(book, "% PV")["1.5"]
    assert ev["1.5"][:3] == pytest.approx(pv[:3])
    assert load_workbook(book)["Resumen"][STATUS_CELL].value == datetime(2026, 3, 31)


def test_actualizar_with_the_same_file_keeps_the_values(tmp_path):
    """El mes de la fecha de estado es el que calcularía 'actualizar': volver a cargarlo no cambia nada."""
    project = P(tasks(), status_date=STATUS)
    book, _report, _h = generate(tmp_path, project)
    set_bac(book)
    before = pct_ev(book)
    res = run_update(P(tasks(), status_date=STATUS), cfg(), book, backup=False)
    assert not any(w.id in ("ACT-007", "ACT-021", "ACT-046") for w in res.warnings)
    after = pct_ev(book)
    for code, values in before.items():
        assert after[code][:3] == pytest.approx(values[:3], abs=1e-9), code
    assert read_book(book).status_date == date(2026, 3, 31)


def test_companies_have_their_own_history(tmp_path):
    t = [T(1, 1, "1", "Fase"),
         T(2, 2, "1.1", "Paquete", method="% Completado"),
         T(3, 3, "1.1.1", "De A", company="EmpresaA", percent_complete=100.0, actual_start=datetime(2026, 1, 5, 8),
           actual_finish=datetime(2026, 1, 30, 17)),
         T(4, 3, "1.1.2", "De B", company="EmpresaB", percent_complete=100.0, actual_start=datetime(2026, 3, 2, 8),
           actual_finish=datetime(2026, 3, 20, 17))]
    book, _report, history = generate(tmp_path, P(t, status_date=STATUS), cfg(multi_enabled=True))
    assert history.loaded
    assert pct_ev(book, "% EV-EmpresaA")["1.1"][:3] == pytest.approx([1, 0, 0])
    assert pct_ev(book, "% EV-EmpresaB")["1.1"][:3] == pytest.approx([0, 0, 1])


def test_without_status_date_or_progress_nothing_is_loaded(tmp_path):
    book, report, history = generate(tmp_path, P(tasks()))           # avance sin fecha de estado
    assert not history.loaded and [i.id for i in history.issues] == ["LIB-017"]
    assert all(v is None for code, values in pct_ev(book).items() if code != "1.0" for v in values)
    assert load_workbook(book)["Resumen"][STATUS_CELL].value is None
    quiet = [T(1, 1, "1", "Fase"), T(2, 2, "1.1", "A")]               # sin avance: sin observación
    _book, report, history = generate(tmp_path, P(quiet, status_date=STATUS))
    assert not history.loaded and not history.issues


def test_generar_loads_history_but_not_with_migrar_desde(tmp_path, example_files):
    book = tmp_path / "jun.xlsx"
    assert main(["generar", "--origen", str(example_files["jun"]), "--config", str(example_files["config"]),
                 "--salida", str(book)]) == 0
    b = read_book(book)
    assert b.status_date == date(2027, 6, 30)
    idx = b.horizon.index(date(2027, 6, 30))
    loaded = [i for i in range(b.horizon.months) if any(b.pct_ev[""][s.row][i] is not None for s in b.input_rows)]
    assert loaded == list(range(idx + 1))
    # Julio se carga con actualizar, como siempre
    from evm.config import load_config
    from evm.reader_xml import read_xml
    conf = load_config(example_files["config"])
    res = run_update(read_xml(example_files["jul"], conf), conf, book, backup=False)
    assert res.month_index == idx + 1
    # Con --migrar-desde se conserva el % ganado del libro anterior (no se reconstruye)
    out = tmp_path / "con_historial.xlsx"
    assert main(["generar", "--origen", str(example_files["jun"]), "--config", str(example_files["config"]),
                 "--salida", str(out), "--migrar-desde", str(book)]) == 0
    m = read_book(out)
    assert [m.pct_ev[""][s.row][idx + 1] for s in m.input_rows] == [b2 for b2 in
                                                                    (read_book(book).pct_ev[""][s.row][idx + 1]
                                                                     for s in read_book(book).input_rows)]


def test_option_ev_historico(tmp_path, example_files):
    """[exportacion] ev_historico: activado por defecto (también en config.ejemplo.toml); con false, generar no carga
    el % ganado."""
    from pathlib import Path

    from evm.config import ConfigError, load_config
    assert Config().ev_history and load_config(Path(__file__).parents[1] / "config.ejemplo.toml").ev_history
    bad = tmp_path / "mal.toml"
    bad.write_text('[exportacion]\nev_historico = "no"\n', encoding="utf-8")
    with pytest.raises(ConfigError, match="'ev_historico' debe ser true o false"):
        load_config(bad)
    off = tmp_path / "sin_historial.toml"
    off.write_text(example_files["config"].read_text(encoding="utf-8").replace(
        "[exportacion]\n", "[exportacion]\nev_historico = false\n"), encoding="utf-8")
    book = tmp_path / "libro.xlsx"
    assert main(["generar", "--origen", str(example_files["jun"]), "--config", str(off), "--salida", str(book)]) == 0
    b = read_book(book)
    assert b.status_date is None
    assert all(v is None for s in b.input_rows for v in b.pct_ev[""][s.row])
