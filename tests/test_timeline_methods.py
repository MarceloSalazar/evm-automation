from datetime import date, datetime

import pytest

from evm.config import ConfigError, parse_date, parse_month
from evm.indicators import compute, earned_schedule
from evm.methods import FIXED, LOE, MILESTONES, PCT, PHYSICAL, UNDEFINED, UNITS, UNKNOWN, ev_fraction, parse_method
from evm.model import Task
from evm.timeline import Horizon, add_months, distribute_linear, round_preserving_total, split_segment
from evm.wbs import build_tree, pad_code


def test_working_calendar_and_holidays():
    from evm.workcalendar import WorkCalendar
    cal = WorkCalendar()
    jan = cal.minutes_by_month(datetime(2026, 1, 1), datetime(2026, 2, 1))
    assert jan[(2026, 1)] == 22 * 8 * 60                                 # 22 días hábiles de 8 h
    cal.add_holidays({date(2026, 1, 1)})
    assert cal.minutes_by_month(datetime(2026, 1, 1), datetime(2026, 2, 1))[(2026, 1)] == 21 * 8 * 60
    assert not cal.is_working_day(date(2026, 1, 3))                      # sábado


def test_parse_date_rejects_invalid_dates():
    with pytest.raises(ConfigError):
        parse_date("2026-02-30")


def test_add_months_and_horizon():
    assert add_months(date(2026, 1, 15), -1) == date(2025, 12, 1)
    assert add_months(date(2026, 11, 1), 3) == date(2027, 2, 1)
    h = Horizon(date(2026, 1, 1), 120)
    assert h.index(date(2026, 1, 31)) == 0
    assert h.index((2035, 12)) == 119
    assert h.end == date(2035, 12, 31)
    vec, before, after = h.to_vector({(2025, 12): 5.0, (2026, 2): 7.0, (2036, 1): 3.0})
    assert vec[1] == 7.0 and before == 5.0 and after == 3.0


def test_distribute_linear_by_working_days():
    d = distribute_linear(1000, datetime(2026, 1, 26, 8), datetime(2026, 2, 6, 17))
    assert d == pytest.approx({(2026, 1): 500.0, (2026, 2): 500.0})
    assert sum(distribute_linear(1234.56, datetime(2026, 3, 2), datetime(2027, 8, 20)).values()) == pytest.approx(1234.56)
    assert distribute_linear(50, datetime(2026, 5, 4, 8), datetime(2026, 5, 4, 17)) == {(2026, 5): 50.0}
    # fin a las 00:00 => el trabajo termina el día anterior
    d = distribute_linear(100, datetime(2026, 1, 26, 8), datetime(2026, 2, 2, 0, 0))
    assert d == pytest.approx({(2026, 1): 100.0})


def test_split_segment_working_hours():
    d = split_segment(1000, datetime(2026, 1, 26, 8), datetime(2026, 2, 6, 17))
    assert d == pytest.approx({(2026, 1): 500.0, (2026, 2): 500.0})
    assert split_segment(10, datetime(2026, 1, 5, 8), datetime(2026, 1, 9, 17)) == {(2026, 1): 10.0}
    # tramo sin horas laborables (fin de semana) => por tiempo transcurrido
    d = split_segment(10, datetime(2026, 1, 31, 0), datetime(2026, 2, 1, 12))
    assert sum(d.values()) == pytest.approx(10)


def test_round_preserving_total():
    vals = [1 / 3] * 3
    r = round_preserving_total(vals, 2)
    assert sum(r) == pytest.approx(1.0)


def test_pad_code():
    assert pad_code("1", 3) == "1.0.0"
    assert pad_code("1.2", 3) == "1.2.0"
    assert pad_code("1.2.3", 3) == "1.2.3"
    assert pad_code("A-01", 3) == "A-01"                  # los códigos no numéricos no se rellenan
    assert pad_code("INTERNO", 3) == "INTERNO"
    assert pad_code("", 3) == ""


@pytest.mark.parametrize("text,kind,start", [
    ("0/100", FIXED, 0.0), ("50/50", FIXED, 0.5), ("20-80", FIXED, 0.2), ("Hitos ponderados", MILESTONES, 0),
    ("% Completado", PCT, 0), ("porcentaje completado", PCT, 0), ("% Físico completado", PHYSICAL, 0),
    ("Unidades completadas", UNITS, 0), ("Avance lineal (LOE)", LOE, 0), ("avance linear", LOE, 0),
    ("Por definir", UNDEFINED, 0), ("", UNDEFINED, 0), ("Esfuerzo prorrateado (AE)", UNKNOWN, 0),
    ("30/50", UNKNOWN, 0),
])
def test_parse_method(text, kind, start):
    m = parse_method(text)
    assert m.kind == kind
    if kind == FIXED:
        assert m.start_fraction == start


def _tree():
    wp = Task(uid=1, id=1, name="WP", outline_level=3, baseline_cost=300)
    a = Task(uid=2, id=2, name="A", outline_level=4, baseline_cost=100, percent_complete=100,
             actual_start=datetime(2026, 1, 5), actual_finish=datetime(2026, 1, 9))
    b = Task(uid=3, id=3, name="B", outline_level=4, baseline_cost=200, percent_complete=25,
             physical_percent_complete=50, actual_start=datetime(2026, 1, 12))
    m1 = Task(uid=4, id=4, name="H1", outline_level=4, milestone=True, milestone_weight=30, percent_complete=100)
    m2 = Task(uid=5, id=5, name="H2", outline_level=4, milestone=True, milestone_weight=70)
    build_tree([wp, a, b, m1, m2])
    wp.percent_complete = 40
    return wp


def test_ev_fraction_methods():
    wp = _tree()
    assert ev_fraction(parse_method("0/100"), wp, 0.5)[0] == 0.0
    assert ev_fraction(parse_method("50/50"), wp, 0.5)[0] == 0.5
    # ponderado por costo: (100*100 + 200*25) / 300 = 50 %
    assert ev_fraction(parse_method("% Completado"), wp, 0.5)[0] == pytest.approx(0.5)
    # físico: (100*0 + 200*50) / 300
    assert ev_fraction(parse_method("% Físico completado"), wp, 0.5)[0] == pytest.approx(100 / 300)
    assert ev_fraction(parse_method("Hitos ponderados"), wp, 0.5)[0] == pytest.approx(0.3)
    assert ev_fraction(parse_method("Avance lineal (LOE)"), wp, 0.42)[0] == pytest.approx(0.42)
    frac, note = ev_fraction(parse_method("Por definir"), wp, 0.5)
    assert frac == pytest.approx(0.5) and note


def test_ev_fraction_complete_and_not_started():
    t = Task(uid=1, id=1, name="T", outline_level=3, baseline_cost=10, percent_complete=100)
    assert ev_fraction(parse_method("0/100"), t, 0)[0] == 1.0
    t = Task(uid=1, id=1, name="T", outline_level=3, baseline_cost=10)
    assert ev_fraction(parse_method("20/80"), t, 0)[0] == 0.0


def test_earned_schedule_lipke_example():
    # PV acumulado mensual; EV = 450 cae entre el mes 4 (400) y el 5 (500) => ES = 4.5
    cum = [100, 200, 300, 400, 500, 600]
    assert earned_schedule(cum, 450, 600) == pytest.approx(4.5)
    assert earned_schedule(cum, 600, 600) == pytest.approx(6.0)
    assert earned_schedule(cum, 50, 600) == pytest.approx(0.5)


def test_compute_indicators():
    ind = compute([100] * 6, [80, 80, 0, 0, 0, 0], [100, 100, 0, 0, 0, 0], 600, 1)
    assert ind.pv == 200 and ind.ev == 160 and ind.ac == 200
    assert ind.spi == pytest.approx(0.8) and ind.cpi == pytest.approx(0.8)
    assert ind.eac_cpi == pytest.approx(750)
    assert ind.es == pytest.approx(1.6) and ind.at == 2


def test_parse_month_and_date():
    assert parse_month("2026-03") == date(2026, 3, 1)
    assert parse_date("31/10/2026") == date(2026, 10, 31)
    with pytest.raises(ConfigError):
        parse_month("2026/13")
