"""Robustez del cálculo: estructura EDT y PV, métodos de medición, multi-empresa, reparto en el
tiempo, indicadores y configuración.

Incluye los casos de los problemas detectados en la revisión del código, ya corregidos, para que no
vuelvan a aparecer.
"""
import random
from datetime import date, datetime, timedelta
from itertools import accumulate

import pytest

from evm.companies import ac_by_company, ev_by_company
from evm.config import Config, ConfigError, load_config
from evm.indicators import compute, earned_schedule
from evm.methods import FIXED, UNKNOWN, ev_fraction, parse_method, weighted_percent
from evm.model import Task
from evm.timeline import distribute_linear, lump, round_preserving_total, split_segment
from evm.wbs import KIND_RESIDUAL, LinearPVSource, build_tree

from builders import (F, S, T, assert_report_invariants, by_name, ev_after_update, expected_bac, issues_with,
                      random_tree, report, rows_by_code, seeds, timephased_source, tree)


# ----------------------------------------------------------------------------
# EDT y PV: invariantes sobre árboles aleatorios
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("multi", [False, True])
@pytest.mark.parametrize("seed", seeds(10))
def test_report_invariants_random_trees(seed, multi):
    rng = random.Random(seed)
    tasks = random_tree(rng, multi)
    cfg = Config(max_level=rng.choice([1, 2, 3, 4]), months=rng.choice([1, 6, 24, 120]),
                 start_month=rng.choice([None, date(2025, 6, 1), date(2027, 1, 1)]),
                 distribution=rng.choice(["habiles", "calendario"]), multi_enabled=multi)
    bac = expected_bac(tasks, cfg.max_level)
    rep = report(tasks, cfg, LinearPVSource(cfg.distribution))
    assert_report_invariants(rep)
    assert rep.total_bac == pytest.approx(bac, abs=1e-6)


# ----------------------------------------------------------------------------
# EDT y PV: casos puntuales
# ----------------------------------------------------------------------------
def test_inactive_tasks_cost_is_reported():
    """El costo excluido incluye las tareas inactivas de nivel superior y las que dependen de otra tarea."""
    rep = report([T(1, 1, "1", "A", 100, active=False), T(2, 1, "2", "B", 50)])
    assert rep.total_bac == pytest.approx(50)
    [issue] = issues_with(rep, "inactiva")
    assert issue.message == ("Se excluyó 1 tarea inactiva (costo de línea base 100.00, que no forma parte del "
                             "BAC).")
    rep = report([T(1, 1, "1", "A", 100, active=False), T(2, 1, "2", "B", 70),
                  T(3, 2, "2.1", "B1", 50), T(4, 2, "2.2", "B2", 20, active=False)])
    assert rep.total_bac == pytest.approx(50)
    assert issues_with(rep, "Se excluyeron 2 tareas inactivas (costo de línea base 120.00")


def test_horizon_start_ignores_omitted_tasks():
    """Sin mes de inicio, el mes 1 es el primer mes con PV del libro (las tareas omitidas no cuentan)."""
    tasks = [T(1, 1, "", "Hito suelto", 999, bs=datetime(2024, 3, 4, 8), bf=datetime(2024, 3, 8, 17)),
             T(2, 1, "1", "Fase", 100)]
    rep = report(tasks)
    assert rep.total_bac == pytest.approx(100) and issues_with(rep, "Se omitió 1 tarea")
    assert rep.horizon.start == date(2026, 1, 1)
    assert rep.total_vector[0] > 0 and not rep.pv_before


def test_plan_prefix_variants():
    for code in ("INTERNO.1.1", "interno-1.1", "INTERNO 1.1", "INTERNO_1.1", "INTERNO/1.1"):
        rep = report([T(1, 1, "INTERNO", "Plan", 100), T(2, 2, code, "X", 100)], Config(plan="INTERNO"))
        assert [r.code for r in rep.rows] == ["1.1.0"], code
    for code in ("INTERNO1.1", "INTERNOS.1"):                          # sin separador: no es el prefijo
        rep = report([T(1, 1, "INTERNO", "Plan", 100), T(2, 2, code, "X", 100)], Config(plan="INTERNO"))
        assert [r.code for r in rep.rows] == [code]
    # Planificación con código numérico: "10" no depende de "1"
    rep = report([T(1, 1, "1", "Plan", 100), T(2, 2, "1.1", "X", 60), T(3, 2, "1.2", "Y", 40),
                  T(4, 1, "10", "Z", 5)], Config(plan="1"))
    assert [(r.code, r.bac) for r in rep.rows] == [("1.0.0", 60), ("2.0.0", 40)]


@pytest.mark.parametrize("level, codes", [
    (1, ["1"]),
    (10, ["1.0.0.0.0.0.0.0.0.0", "1.1.0.0.0.0.0.0.0.0", "1.1.1.0.0.0.0.0.0.0", "1.1.2.0.0.0.0.0.0.0",
          "1.1.2.1.0.0.0.0.0.0"]),
])
def test_max_level_bounds(level, codes):
    tasks = tree("""
        1 F 300
          1.1 G 300
            1.1.1 P 100
            1.1.2 Q 200
              1.1.2.1 R 200
    """)
    rep = report(tasks, Config(max_level=level))
    assert [r.code for r in rep.rows] == codes and rep.total_bac == pytest.approx(300)
    assert_report_invariants(rep)


def test_top_level_uncoded_container_is_omitted_but_children_exported():
    rep = report(tree("""
        - Contenedor 350
          1 A 100
          2 B 200
          - Suelto 50
    """))
    assert [r.code for r in rep.rows] == ["1.0.0", "2.0.0"] and rep.total_bac == pytest.approx(300)
    assert issues_with(rep, "Se omitieron 2 tareas") and "50.00" in issues_with(rep, "Se omitieron")[0].message


def test_cost_without_dates_and_empty_report_are_reported():
    rep = report([T(1, 1, "1", "Hito", 100, bs=None, bf=None, start=None, finish=None, milestone=True)])
    assert rep.total_bac == 0 and issues_with(rep, "no tiene fechas")
    rep = report([T(1, 1, "", "Sin código", 100)])
    assert not rep.rows and issues_with(rep, "Ninguna tarea tiene código EDT")


@pytest.mark.parametrize("kw, current, expected", [
    ({"fixed_cost": 300, "fixed_cost_accrual": "inicio"}, 300, {(2026, 1): 300}),
    ({"fixed_cost": 300, "fixed_cost_accrual": "fin"}, 300, {(2026, 3): 300}),
    ({"fixed_cost": 300, "fixed_cost_accrual": "inicio"}, 500, {(2026, 1): 100, (2026, 2): 100, (2026, 3): 100}),
    ({"milestone": True}, 300, {(2026, 3): 300}),                       # hito: mes de fin
    ({"bs": None, "bf": None}, 300, {(2026, 1): 100, (2026, 2): 100, (2026, 3): 100}),  # fechas actuales
])
def test_linear_leaf_accrual_milestone_and_current_dates(kw, current, expected):
    """Costo fijo al comienzo / al final solo si la tarea no tiene costo de recursos."""
    task = T(1, 1, "1", "A", 300, current_cost=current, **kw)
    assert LinearPVSource().leaf(task) == pytest.approx(expected)


# ----------------------------------------------------------------------------
# PV con datos escalonados de MS Project
# ----------------------------------------------------------------------------
def test_timephased_mismatch_is_scaled_and_reported_with_padded_code():
    tasks = [T(1, 1, "1", "Fase", 300),
             T(2, 2, "1.1", "A", 100, timephased_pv={(2026, 1): 60.0, (2026, 2): 60.0}),   # suma 120
             T(3, 2, "1.2", "B", 200)]                                                       # sin datos: lineal
    rep = report(tasks, source=timephased_source())
    assert rows_by_code(rep)["1.1.0"].pv == pytest.approx({(2026, 1): 50.0, (2026, 2): 50.0})
    [issue] = issues_with(rep, "suma 120.00")
    assert issue.code == "1.1.0"
    assert issues_with(rep, "se repartió en forma lineal", "Info")
    assert "sin datos mensuales: reparto lineal" in rep.pv_source


def test_timephased_without_any_data_warns_linear_distribution():
    rep = report([T(1, 1, "1", "A", 100)], source=timephased_source())
    assert issues_with(rep, "MS Project no entregó la distribución mensual", "Advertencia")
    assert rep.pv_source.startswith("Distribución lineal")


def test_timephased_with_opposite_sign_falls_back_to_linear():
    task = T(1, 1, "1", "A", 100, timephased_pv={(2026, 1): 150.0, (2026, 2): -200.0})    # suma -50
    rep = report([task], source=timephased_source())
    assert rep.rows[0].pv == pytest.approx({(2026, 1): 100 / 3, (2026, 2): 100 / 3, (2026, 3): 100 / 3})
    [issue] = issues_with(rep, "signo opuesto", "Advertencia")
    assert issue.code == "1.0.0" and "suma -50.00" in issue.message and "forma lineal" in issue.message


def test_timephased_summary_residual_uses_rollup_difference_or_linear():
    fase = T(1, 1, "1", "Fase", 400, timephased_pv={(2026, 1): 250.0, (2026, 2): 150.0})
    hija = T(2, 2, "1.1", "A", 300, timephased_pv={(2026, 1): 150.0, (2026, 2): 150.0})
    residual = next(r for r in report([fase, hija], source=timephased_source()).rows if r.kind == KIND_RESIDUAL)
    assert residual.pv == pytest.approx({(2026, 1): 100.0})            # distribución de la resumen - hijos
    fase.timephased_pv = {(2026, 1): 100.0, (2026, 2): 300.0}           # diferencia negativa en enero
    residual = next(r for r in report([fase, hija], source=timephased_source()).rows if r.kind == KIND_RESIDUAL)
    assert residual.pv == pytest.approx({(2026, 1): 100 / 3, (2026, 2): 100 / 3, (2026, 3): 100 / 3})


def test_inactive_child_with_timephased_data_is_subtracted_from_summary():
    fase = T(1, 1, "1", "Fase", 500, timephased_pv={(2026, 1): 300.0, (2026, 2): 200.0})
    tasks = [fase, T(2, 2, "1.1", "A", 100, timephased_pv={(2026, 1): 100.0}),
             T(3, 2, "1.2", "B", 200, active=False, timephased_pv={(2026, 2): 200.0})]
    residual = next(r for r in report(tasks, source=timephased_source()).rows if r.kind == KIND_RESIDUAL)
    assert residual.pv == pytest.approx({(2026, 1): 200.0})


# ----------------------------------------------------------------------------
# Métodos de medición
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("text, kind, start", [
    ("0 / 100", FIXED, 0.0), ("50%/50%", FIXED, 0.5), ("100/0", FIXED, 1.0), ("050/050", FIXED, 0.5),
    ("1/99", FIXED, 0.01), ("60/30", UNKNOWN, None), ("0/0", UNKNOWN, None), ("cualquier texto", UNKNOWN, None),
])
def test_parse_method_fixed_formula_edges(text, kind, start):
    m = parse_method(text)
    assert m.kind == kind and (start is None or m.start_fraction == start)


def milestone_package(weights, done=(0,)):
    wp = Task(uid=1, id=1, name="WP", outline_level=3, baseline_cost=300, percent_complete=30)
    hitos = [Task(uid=2 + i, id=2 + i, name=f"H{i}", outline_level=4, milestone=True, milestone_weight=w,
                  percent_complete=100 if i in done else 0) for i, w in enumerate(weights)]
    build_tree([wp, *hitos])
    return wp


def test_weighted_milestones_without_weights_count_equally():
    for weights in ((None, None, None), (0, 0, 0)):
        frac, note = ev_fraction(parse_method("Hitos ponderados"), milestone_package(weights), 0)
        assert frac == pytest.approx(1 / 3) and "por igual" in note


def test_weighted_milestones_negative_weight_counts_as_zero():
    """Un peso negativo no es válido: el hito pesa 0 y se informa."""
    method = parse_method("Hitos ponderados")
    frac, note = ev_fraction(method, milestone_package((100, -50)), 0)
    assert frac == pytest.approx(1.0) and "negativo" in note
    frac, note = ev_fraction(method, milestone_package((100, -50), done=(1,)), 0)
    assert frac == 0.0 and "negativo" in note


def test_weighted_milestones_all_negative_weights_count_equally():
    """Sin ningún peso válido, los hitos pesan igual (como sin pesos) y se informa."""
    frac, note = ev_fraction(parse_method("Hitos ponderados"), milestone_package((-10, -50)), 0)
    assert frac == pytest.approx(0.5)
    assert "sin hitos" not in note and "por igual" in note and "negativo" in note


def test_weighted_percent_without_costs():
    wp = Task(uid=1, id=1, name="WP", outline_level=3, percent_complete=40)
    a = Task(uid=2, id=2, name="A", outline_level=4, percent_complete=100, physical_percent_complete=100)
    b = Task(uid=3, id=3, name="B", outline_level=4)
    build_tree([wp, a, b])
    assert weighted_percent(wp) == pytest.approx(0.4)                  # % completado de la tarea en MS Project
    assert weighted_percent(wp, physical=True) == pytest.approx(0.5)   # promedio de las subtareas


@pytest.mark.parametrize("seed", seeds(20))
def test_ev_fraction_always_in_unit_interval(seed):
    """Con pesos de hito no negativos, la fracción del BAC ganada está siempre entre 0 y 1."""
    rng = random.Random(seed)
    wp = Task(uid=1, id=1, name="WP", outline_level=3, baseline_cost=rng.choice([0, 300]),
              percent_complete=rng.uniform(-10, 120), bcwp=rng.choice([None, rng.uniform(-50, 500)]))
    kids = [Task(uid=i, id=i, name=f"K{i}", outline_level=4, baseline_cost=rng.uniform(-20, 200),
                 milestone=rng.random() < 0.4, milestone_weight=rng.choice([None, 0, rng.uniform(0, 50)]),
                 percent_complete=rng.uniform(-10, 120), physical_percent_complete=rng.uniform(-10, 120))
            for i in range(2, rng.randint(2, 7))]
    build_tree([wp, *kids])
    for text in ("0/100", "30/70", "Hitos ponderados", "% Completado", "% Físico completado", "Unidades",
                 "Avance lineal (LOE)", "CPTR de MS Project", "Por definir", "otro"):
        frac, _ = ev_fraction(parse_method(text), wp, rng.uniform(-0.5, 1.5), bac=rng.choice([0.0, 300.0]))
        assert 0.0 <= frac <= 1.0, text


# ----------------------------------------------------------------------------
# EV de las filas de otros costos (-D)
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("multi", [False, True])
def test_residual_row_ev_uses_only_its_own_tasks(tmp_path, multi):
    tasks = tree("""
        1 Fase 150
          1.1 A 100
          - Suelto 50
    """, company="EmpresaA")
    t = by_name(tasks)
    t["Fase"].percent_complete, t["Fase"].acwp = 33, 50.0
    t["A"].acwp = 0.0
    t["Suelto"].percent_complete, t["Suelto"].acwp, t["Suelto"].actual_start = 100, 50.0, S
    ev = ev_after_update(tasks, Config(months=12, multi_enabled=multi), tmp_path, date(2026, 3, 31))
    assert ev["1.1.0"] == pytest.approx(0.0)
    assert ev["1.0.0-D"] == pytest.approx(50.0)                        # hoy, una empresa: 16.67


# ----------------------------------------------------------------------------
# Multi-empresa
# ----------------------------------------------------------------------------
def company_package():
    wp = Task(uid=1, id=1, name="WP", outline_level=1, baseline_cost=300, percent_complete=40, bcwp=90.0,
              actual_start=S)
    a = Task(uid=2, id=2, name="A", outline_level=2, baseline_cost=100, percent_complete=100,
             physical_percent_complete=80, bcwp=100.0)
    b = Task(uid=3, id=3, name="B", outline_level=2, baseline_cost=200, percent_complete=10,
             physical_percent_complete=20, bcwp=-10.0)
    m = Task(uid=4, id=4, name="H", outline_level=2, milestone=True, percent_complete=100, bcwp=0.0)
    build_tree([wp, a, b, m])
    return [wp, a, b, m]


@pytest.mark.parametrize("text", ["0/100", "30/70", "% Completado", "% Físico completado", "Unidades completadas",
                                  "Hitos ponderados", "Avance lineal (LOE)", "CPTR de MS Project"])
def test_multi_all_company_a_matches_single_mode(text):
    tasks = company_package()
    method = parse_method(text)
    frac, _ = ev_fraction(method, tasks[0], 0.25, bac=300, bcwp=tasks[0].bcwp)
    ev = ev_by_company(method, tasks, {t.uid: 1.0 for t in tasks}, (300.0, 0.0), (75.0, 0.0), frac)
    assert ev == pytest.approx((300 * frac, 0.0))


def test_ev_by_company_cptr_fallbacks():
    tasks = company_package()
    wp, a, b, m = tasks
    shares = {wp.uid: 0.5, a.uid: 1.0, b.uid: 0.0, m.uid: 1.0}
    cptr = parse_method("CPTR de MS Project")
    assert ev_by_company(cptr, tasks, shares, (100.0, 200.0), (0.0, 0.0), 0.3) == pytest.approx((100.0, 0.0))
    # CPTR calculado con otra línea base: fracción = CPTR / costo de esa línea base de cada empresa
    wp.ev_baseline_cost, a.ev_baseline_cost, b.ev_baseline_cost = 400.0, 200.0, 200.0
    assert ev_by_company(cptr, tasks, shares, (100.0, 200.0), (0.0, 0.0), 0.3, True) == pytest.approx((50.0, 0.0))
    # Sin CPTR en alguna tarea de la fila: % completado de las tareas de cada empresa
    m.bcwp = None
    assert ev_by_company(cptr, tasks, shares, (100.0, 200.0), (0.0, 0.0), 0.3) == pytest.approx((100.0, 20.0))
    # Sin costos para ponderar: avance del paquete completo
    for t in tasks:
        t.baseline_cost = 0.0
    assert ev_by_company(parse_method("% Completado"), tasks, shares, (100.0, 200.0), (0.0, 0.0), 0.3) == \
        pytest.approx((30.0, 60.0))


def test_ac_by_company_uses_actual_cost_without_acwp():
    tasks = company_package()
    wp, a, b, _ = tasks
    wp.actual_cost, a.actual_cost, b.actual_cost = 70.0, 40.0, 20.0     # sin CRTR: costo real
    assert ac_by_company(tasks, {wp.uid: 0.5, a.uid: 1.0, b.uid: 0.0}) == pytest.approx((45.0, 25.0))


# ----------------------------------------------------------------------------
# Reparto en el tiempo y redondeo
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("seed", seeds(10))
def test_round_preserving_total_property(seed):
    rng = random.Random(seed)
    for _ in range(50):
        values = [rng.uniform(-1e6, 1e6) * rng.choice([0, 1e-3, 1]) for _ in range(rng.choice([1, 2, 12, 120]))]
        out = round_preserving_total(values, 2)
        assert len(out) == len(values)
        assert sum(out) == pytest.approx(round(sum(values), 2), abs=1e-6)
        assert all(round(v, 2) == v for v in out)
        assert all(abs(o - v) <= 0.005 * len(values) + 1e-9 for o, v in zip(out, values))


@pytest.mark.parametrize("mode", ["habiles", "calendario"])
@pytest.mark.parametrize("seed", seeds(5))
def test_distribute_linear_preserves_amount(seed, mode):
    rng = random.Random(seed)
    for _ in range(20):
        start = datetime(2025, 1, 1) + timedelta(minutes=rng.randint(0, 2_000_000))
        finish = start + timedelta(minutes=rng.randint(-50_000, 1_000_000))   # también fin anterior al inicio
        amount = rng.uniform(-1e5, 1e6)
        dist = distribute_linear(amount, start, finish, mode)
        assert sum(dist.values()) == pytest.approx(amount)
        lo, hi = min(start, finish), max(start, finish)
        assert all((lo.year, lo.month) <= k <= (hi.year, hi.month) for k in dist)


def test_distribute_linear_and_lump_edges():
    assert distribute_linear(100, None, datetime(2026, 3, 10)) == {(2026, 3): 100.0}
    assert distribute_linear(100, date(2026, 3, 10), None) == {(2026, 3): 100.0}
    assert distribute_linear(100, None, None) == {} and distribute_linear(0, S, F) == {}
    # Fin de semana entre dos meses: sin horas laborables, se reparte por tiempo transcurrido
    weekend = distribute_linear(100, datetime(2026, 1, 31), datetime(2026, 2, 1, 23))
    assert weekend == pytest.approx({(2026, 1): 100 * 24 / 47, (2026, 2): 100 * 23 / 47})
    assert lump(5, None) == {} and lump(0, S) == {}
    assert split_segment(0, S, F) == {} and round_preserving_total([]) == []


# ----------------------------------------------------------------------------
# Indicadores
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("seed", seeds(10))
def test_earned_schedule_monotonic_in_ev(seed):
    """ES no decrece al aumentar el EV, está dentro del horizonte y solo es None si la curva no llega."""
    rng = random.Random(seed)
    for _ in range(50):
        n = rng.randint(1, 24)
        cum = list(accumulate(rng.choice([0, 0, rng.uniform(0, 1000)]) for _ in range(n)))
        bac = cum[-1] + rng.choice([0, rng.uniform(0, 500)])
        at = rng.randint(1, n)
        previous = 0.0
        for ev in sorted([rng.uniform(0, bac) for _ in range(15)] + [0.0, bac]):
            es = earned_schedule(cum, ev, bac, at)
            if es is None:
                assert ev > cum[-1] - 0.006 or ev >= bac - 0.005
                continue
            assert previous - 1e-9 <= es <= n
            previous = es


def test_earned_schedule_negative_ev_is_zero():
    """EV acumulado negativo (correcciones cargadas a mano): ES = 0, sin error."""
    assert earned_schedule([0.0, 100.0], -1.0, 100.0, at=1) == 0.0
    assert earned_schedule([50.0, 100.0], -1.0, 100.0, at=2) == 0.0
    ind = compute([0.0, 100.0], [-1.0, 0.0], [0.0, 0.0], 100.0, 0)
    assert (ind.ev, ind.es, ind.sv_t) == (-1.0, 0.0, -1.0)


def test_compute_before_first_month():
    ind = compute([100.0, 100.0], [0.0, 0.0], [0.0, 0.0], 200.0, -1)
    assert (ind.pv, ind.ev, ind.ac, ind.at) == (0.0, 0.0, 0.0, 0) and ind.spi is None and ind.spi_t is None


# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("text, message", [
    ('[exportacion]\nincluir_inactivas = "si"\n', "true o false"),
    ('[horizonte]\nmeses = "12"\n', "numérico"),
    ('[semaforo]\nrojo = true\n', "numérico"),
    ('[campos]\nedt = 5\n', "texto"),
    ('[exportacion]\nferiados = "2026-01-01"\n', "lista de fechas"),
    ('[exportacion]\nferiados = [5]\n', "Fecha inválida"),
    ('[metodos]\nlista = ["0/100", 5]\n', "lista de textos"),
    ('[horizonte\n', "sintaxis"),
    ('[horizonte]\nmes_inicio = "2026-13"\n', "Mes inválido"),
    ('[semaforo]\nrojo = 0.97\namarillo = 0.95\n', "umbrales"),
    ('[exportacion]\nlinea_base = 11\n', "linea_base"),
    ('[exportacion]\ndecimales = 5\n', "decimales"),
    ('[exportacion]\ndistribucion = "lineal"\n', "distribucion"),
    ('[exportacion]\nnivel_maximo = 11\n', "nivel_maximo"),
    ('[horizonte]\nmeses = 601\n', "meses"),
    ('campos = "x"\n', r"\[campos\] debe ser una sección"),
    ('exportacion = 5\n', r"\[exportacion\] debe ser una sección"),
    ('[horizonte]\nmeses = 12.7\n', "'meses' debe ser un número entero"),
    ('[exportacion]\nnivel_maximo = 2.9\n', "'nivel_maximo' debe ser un número entero"),
    ('[exportacion]\ndecimales = 1.5\n', "'decimales' debe ser un número entero"),
])
def test_load_config_rejects_invalid_values(tmp_path, text, message):
    path = tmp_path / "config.toml"
    path.write_text(text, encoding="utf-8")
    with pytest.raises(ConfigError, match=message):
        load_config(path)


def test_load_config_accepts_toml_dates_and_normalizes(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text('[horizonte]\nmeses = 12.0\n[exportacion]\nferiados = [2026-01-01, "06/01/2026", '
                    '2026-05-01T00:00:00]\nlinea_base = ""\ndistribucion = "CALENDARIO"\nnivel_maximo = 10\n'
                    'metodo_por_defecto = "50/50"\n[metodos]\nlista = ["0/100", "  "]\n', encoding="utf-8")
    cfg = load_config(path)
    assert cfg.holidays == {date(2026, 1, 1), date(2026, 1, 6), date(2026, 5, 1)}
    assert (cfg.months, cfg.max_level, cfg.baseline, cfg.distribution) == (12, 10, None, "calendario")
    assert type(cfg.months) is int                                     # 12.0 se acepta como entero
    assert cfg.methods == ["50/50", "0/100"]                           # el método por defecto se agrega a la lista
    with pytest.raises(ConfigError, match="No se encontró"):
        load_config(tmp_path / "no_existe.toml")
