"""Robustez del cálculo: estructura EDT y % planificado, métodos de medición, multi-empresa, redondeo,
indicadores y configuración.

Incluye los casos de los problemas detectados en la revisión del código, ya corregidos, para que no
vuelvan a aparecer.
"""
import random
from datetime import date, datetime
from itertools import accumulate

import pytest

from evm.config import Config, ConfigError, load_config
from evm.indicators import compute, earned_schedule
from evm.methods import FIXED, UNKNOWN, parse_method
from evm.model import Task
from evm.percent import RowWeights, TaskTime, company_weights, earned_fraction, full_weights, row_weights
from evm.timeline import round_preserving_total

from builders import (S, T, assert_report_invariants, by_name, earned_after_update, issues_with, random_tree,
                      report, seeds, tree)


def times_of(tasks, minutes=480.0):
    """Tareas de una fila con la misma duración, todas en enero de 2026 (los hitos, sin duración)."""
    return [TaskTime(t, 0.0 if t.milestone else minutes, {(2026, 1): 1.0}, True) for t in tasks]


# ----------------------------------------------------------------------------
# EDT y % planificado: invariantes sobre árboles aleatorios
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("multi", [False, True])
@pytest.mark.parametrize("seed", seeds(10))
def test_report_invariants_random_trees(seed, multi):
    rng = random.Random(seed)
    tasks = random_tree(rng, multi)
    cfg = Config(max_level=rng.choice([1, 2, 3, 4]), months=rng.choice([1, 6, 24, 120]),
                 start_month=rng.choice([None, date(2025, 6, 1), date(2027, 1, 1)]),
                 distribution=rng.choice(["habiles", "calendario"]), multi_enabled=multi)
    assert_report_invariants(report(tasks, cfg))


# ----------------------------------------------------------------------------
# EDT y % planificado: casos puntuales
# ----------------------------------------------------------------------------
def test_inactive_tasks_are_reported():
    """Las tareas excluidas incluyen las inactivas de nivel superior y las que dependen de otra tarea."""
    rep = report([T(1, 1, "1", "A", active=False), T(2, 1, "2", "B")])
    assert [r.code for r in rep.rows] == ["2.0.0"]
    [issue] = issues_with(rep, "inactiva")
    assert issue.message == "Se excluyó 1 tarea inactiva: no forma parte del % planificado."
    rep = report([T(1, 1, "1", "A", active=False), T(2, 1, "2", "B"),
                  T(3, 2, "2.1", "B1"), T(4, 2, "2.2", "B2", active=False)])
    assert [r.code for r in rep.rows] == ["2.0.0", "2.1.0"]
    assert issues_with(rep, "Se excluyeron 2 tareas inactivas")


def test_horizon_start_ignores_omitted_tasks():
    """Sin mes de inicio, el mes 1 es el primer mes con % planificado del libro (las tareas omitidas no cuentan)."""
    tasks = [T(1, 1, "", "Hito suelto", bs=datetime(2024, 3, 4, 8), bf=datetime(2024, 3, 8, 17)),
             T(2, 1, "1", "Fase")]
    rep = report(tasks)
    assert issues_with(rep, "Se omitió 1 tarea")
    assert rep.horizon.start == date(2026, 1, 1)
    assert rep.total_vector[0] > 0 and not issues_with(rep, "anterior")


def test_plan_prefix_variants():
    for code in ("INTERNO.1.1", "interno-1.1", "INTERNO 1.1", "INTERNO_1.1", "INTERNO/1.1"):
        rep = report([T(1, 1, "INTERNO", "Plan"), T(2, 2, code, "X")], Config(plan="INTERNO"))
        assert [r.code for r in rep.rows] == ["1.1.0"], code
    for code in ("INTERNO1.1", "INTERNOS.1"):                          # sin separador: no es el prefijo
        rep = report([T(1, 1, "INTERNO", "Plan"), T(2, 2, code, "X")], Config(plan="INTERNO"))
        assert [r.code for r in rep.rows] == [code]
    # Planificación con código numérico: "10" no depende de "1"
    rep = report([T(1, 1, "1", "Plan"), T(2, 2, "1.1", "X"), T(3, 2, "1.2", "Y"), T(4, 1, "10", "Z")],
                 Config(plan="1"))
    assert [(r.code, r.name) for r in rep.rows] == [("1.0.0", "X"), ("2.0.0", "Y")]


@pytest.mark.parametrize("level, codes", [
    (1, ["1"]),
    (10, ["1.0.0.0.0.0.0.0.0.0", "1.1.0.0.0.0.0.0.0.0", "1.1.1.0.0.0.0.0.0.0", "1.1.2.0.0.0.0.0.0.0",
          "1.1.2.1.0.0.0.0.0.0"]),
])
def test_max_level_bounds(level, codes):
    tasks = tree("""
        1 F
          1.1 G
            1.1.1 P
            1.1.2 Q
              1.1.2.1 R
    """)
    rep = report(tasks, Config(max_level=level))
    assert [r.code for r in rep.rows] == codes
    assert {t.name for r in rep.rows if r.is_input for t in r.leaves} == {"P", "R"}
    assert_report_invariants(rep)


def test_top_level_uncoded_container_is_omitted_but_children_exported():
    rep = report(tree("""
        - Contenedor
          1 A
          2 B
          - Suelto
    """))
    assert [r.code for r in rep.rows] == ["1.0.0", "2.0.0"]
    assert issues_with(rep, "Se omitieron 2 tareas")


def test_task_without_dates_and_empty_report_are_reported():
    rep = report([T(1, 1, "1", "Hito", bs=None, bf=None, start=None, finish=None, milestone=True)])
    assert sum(rep.total_vector) == 0 and issues_with(rep, "Ninguna tarea de la fila tiene fechas")
    rep = report([T(1, 1, "", "Sin código")])
    assert not rep.rows and issues_with(rep, "Ninguna tarea tiene código EDT")


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


def milestone_row(weights, done=(0,)):
    """Fila con un hito por peso; los hitos de `done` están cumplidos."""
    hitos = [Task(uid=2 + i, id=2 + i, name=f"H{i}", outline_level=4, milestone=True, milestone_weight=w,
                  percent_complete=100 if i in done else 0) for i, w in enumerate(weights)]
    return RowWeights({}, times_of(hitos))


def test_weighted_milestones_without_weights_count_equally():
    for weights in ((None, None, None), (0, 0, 0)):
        frac, note = earned_fraction(parse_method("Hitos ponderados"), milestone_row(weights), 0)
        assert frac == pytest.approx(1 / 3) and "por igual" in note.message


def test_weighted_milestones_negative_weight_counts_as_zero():
    """Un peso negativo no es válido: el hito pesa 0 y se informa."""
    method = parse_method("Hitos ponderados")
    frac, note = earned_fraction(method, milestone_row((100, -50)), 0)
    assert frac == pytest.approx(1.0) and "negativo" in note.message
    frac, note = earned_fraction(method, milestone_row((100, -50), done=(1,)), 0)
    assert frac == 0.0 and "negativo" in note.message


def test_weighted_milestones_all_negative_weights_count_equally():
    """Sin ningún peso válido, los hitos pesan igual (como sin pesos) y se informa."""
    frac, note = earned_fraction(parse_method("Hitos ponderados"), milestone_row((-10, -50)), 0)
    assert frac == pytest.approx(0.5)
    assert "sin hitos" not in note.message and "por igual" in note.message and "negativo" in note.message


def test_progress_without_weights_is_simple_average():
    """Tareas sin fechas (sin pesos): el avance de la fila es el promedio simple de sus tareas."""
    a = Task(uid=2, id=2, name="A", outline_level=4, percent_complete=100, physical_percent_complete=80)
    b = Task(uid=3, id=3, name="B", outline_level=4, percent_complete=20)
    rw = RowWeights({}, [TaskTime(t, 0.0, {}, True) for t in (a, b)])
    assert earned_fraction(parse_method("% Completado"), rw, 0)[0] == pytest.approx(0.6)
    assert earned_fraction(parse_method("% Físico completado"), rw, 0)[0] == pytest.approx(0.4)


@pytest.mark.parametrize("seed", seeds(20))
def test_earned_fraction_always_in_unit_interval(seed):
    """Con cualquier avance (también fuera de 0 a 100), el % ganado está siempre entre 0 y 1."""
    rng = random.Random(seed)
    kids = [Task(uid=i, id=i, name=f"K{i}", outline_level=4, milestone=rng.random() < 0.4,
                 milestone_weight=rng.choice([None, 0, rng.uniform(-10, 80)]),
                 percent_complete=rng.uniform(-10, 120), physical_percent_complete=rng.uniform(-10, 120))
            for i in range(2, rng.randint(2, 7))]
    times = [TaskTime(t, 0.0 if t.milestone else rng.choice([0.0, rng.uniform(1, 5000)]),
                      {(2026, 1): 1.0} if rng.random() < 0.9 else {}, True) for t in kids]
    rw = row_weights(times) or RowWeights({}, [])
    for text in ("0/100", "30/70", "Hitos ponderados", "% Completado", "% Físico completado", "Unidades",
                 "Avance lineal (LOE)", "Por definir", "otro"):
        frac, _ = earned_fraction(parse_method(text), rw, rng.uniform(-0.5, 1.5))
        assert 0.0 <= frac <= 1.0, text


# ----------------------------------------------------------------------------
# % ganado de las filas de otras tareas (-D)
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("multi", [False, True])
def test_residual_row_ev_uses_only_its_own_tasks(tmp_path, multi):
    tasks = tree("""
        1 Fase
          1.1 A
          - Suelto
    """, company="EmpresaA")
    t = by_name(tasks)
    t["Fase"].percent_complete = 33
    t["Suelto"].percent_complete, t["Suelto"].actual_start = 100, S
    earned = earned_after_update(tasks, Config(months=12, multi_enabled=multi), tmp_path, date(2026, 3, 31))
    assert earned["1.1.0"] == pytest.approx(0.0)
    assert earned["1.0.0-D"] == pytest.approx(1.0)


# ----------------------------------------------------------------------------
# Multi-empresa
# ----------------------------------------------------------------------------
def company_row():
    a = Task(uid=2, id=2, name="A", outline_level=2, percent_complete=100, physical_percent_complete=80,
             actual_start=S)
    b = Task(uid=3, id=3, name="B", outline_level=2, percent_complete=10, physical_percent_complete=20,
             actual_start=S)
    m = Task(uid=4, id=4, name="H", outline_level=2, milestone=True, milestone_weight=20, percent_complete=100)
    return times_of([a, b, m])


@pytest.mark.parametrize("text", ["0/100", "30/70", "% Completado", "% Físico completado", "Unidades completadas",
                                  "Hitos ponderados", "Avance lineal (LOE)"])
def test_multi_all_company_a_matches_single_mode(text):
    times = company_row()
    method = parse_method(text)
    single = earned_fraction(method, full_weights(times), 0.25)
    assert earned_fraction(method, company_weights(times, lambda t: 1.0, "EmpresaA"), 0.25) == \
        pytest.approx(single)


def test_company_weights_by_share_and_fallback():
    times = company_row()
    share = {2: 1.0, 3: 0.5, 4: 0.0}
    rw = company_weights(times, lambda t: share[t.uid], "EmpresaA")
    # el hito no es de la empresa; las tareas pesan por duración × parte de la empresa (A completa, B a la mitad)
    assert rw.weights == pytest.approx({2: 2 / 3, 3: 1 / 3})
    assert not rw.fallback
    # Una empresa sin tareas en la fila: pesos de la fila completa (fallback)
    none = company_weights(times, lambda t: 0.0, "EmpresaB")
    assert none.fallback and none.weights == pytest.approx(full_weights(times).weights)


# ----------------------------------------------------------------------------
# Redondeo
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
    assert round_preserving_total([]) == []


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
