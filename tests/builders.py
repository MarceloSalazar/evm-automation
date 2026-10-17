"""Constructores compartidos para las pruebas: tareas, proyectos, informes y árboles aleatorios.

Uso típico:

    tasks = tree('''
        1 Fase 150
          1.1 A 100
          - Suelto 50
    ''')
    rep = report(tasks, Config(max_level=2))
    assert_report_invariants(rep)
"""
from __future__ import annotations

import os
import random
import textwrap
from datetime import date, datetime, timedelta

import pytest

from evm.config import Config
from evm.model import ProjectData, Task
from evm.wbs import (KIND_RESIDUAL, KIND_SUMMARY, LinearPVSource, Report, TimephasedPVSource, build_report,
                     build_tree, code_level, tolerance)

# Fechas de línea base por defecto: tres meses completos (enero a marzo de 2026)
S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)


def seeds(default: int) -> range:
    """Semillas de las pruebas aleatorias; la variable de entorno EVM_SEMILLAS permite ampliar la cantidad."""
    return range(int(os.environ.get("EVM_SEMILLAS", default)))


# ----------------------------------------------------------------------------
# Tareas y proyectos
# ----------------------------------------------------------------------------
def T(uid: int, level: int, wbs: str, name: str, cost: float = 0.0, *, bs: datetime | None = S,
      bf: datetime | None = F, start: datetime | None = S, finish: datetime | None = F,
      current_cost: float | None = None, **kw) -> Task:
    """Tarea con línea base (costo `cost`, fechas `bs`-`bf`) y programa actual (`start`-`finish`).

    El costo actual es igual al de la línea base salvo que se indique `current_cost`.
    """
    return Task(uid=uid, id=uid, name=name, outline_level=level, wbs=wbs, baseline_cost=cost,
                baseline_start=bs, baseline_finish=bf, has_baseline=bs is not None, start=start, finish=finish,
                cost=cost if current_cost is None else current_cost, **kw)


def P(tasks: list[Task], **kw) -> ProjectData:
    return ProjectData(name="Programa", source="p.xml", backend="xml", tasks=tasks, **kw)


def tree(text: str, **common) -> list[Task]:
    """Tareas a partir de un esquema con sangría de 2 espacios por nivel.

    Cada línea es "<código EDT> <nombre> <costo>"; el código "-" indica una tarea sin código EDT
    y el nombre puede tener espacios. Los argumentos adicionales se aplican a todas las tareas.
    """
    tasks = []
    for line in textwrap.dedent(text).strip("\n").splitlines():
        if not line.strip():
            continue
        code, *name, cost = line.split()
        level = (len(line) - len(line.lstrip(" "))) // 2 + 1
        tasks.append(T(len(tasks) + 1, level, "" if code == "-" else code, " ".join(name), float(cost), **common))
    return tasks


def by_name(tasks: list[Task]) -> dict[str, Task]:
    return {t.name: t for t in tasks}


# ----------------------------------------------------------------------------
# Informes
# ----------------------------------------------------------------------------
def timephased_source() -> TimephasedPVSource:
    """Fuente de PV que usa `Task.timephased_pv` (como el lector XML)."""
    return TimephasedPVSource(lambda t: t.timephased_pv, "Datos escalonados de la línea base")


def report(tasks: list[Task], cfg: Config | None = None, source=None, **project_kw) -> Report:
    return build_report(P(tasks, **project_kw), cfg or Config(), source or LinearPVSource())


def rows_by_code(rep: Report) -> dict:
    return {r.code: r for r in rep.rows}


def issues_with(rep: Report, text: str, severity: str | None = None) -> list:
    """Observaciones de la hoja Validación que contienen `text` (opcionalmente, de una severidad)."""
    return [i for i in rep.issues if text in i.message and (severity is None or i.severity == severity)]


def ev_after_update(tasks: list[Task], cfg: Config, folder, status: date) -> dict[str, float]:
    """Genera el libro, simula `actualizar` a la fecha de corte y devuelve el EV acumulado por código."""
    from evm.update import run_update
    from evm.workbook import write_workbook
    book = folder / "libro.xlsx"
    write_workbook(build_report(P(tasks), cfg, LinearPVSource()), cfg, book)
    res = run_update(P(tasks), cfg, book, status, backup=False, dry_run=True)
    return {r.code: r.ev_cum for r in res.rows}


# ----------------------------------------------------------------------------
# Árboles aleatorios e invariantes
# ----------------------------------------------------------------------------
def random_tree(rng: random.Random, multi: bool = False) -> list[Task]:
    """EDT aleatoria con costos no negativos.

    Hasta 6 niveles de esquema y 1 a 3 hijos por tarea; 20 % de tareas sin código EDT, 5 %
    inactivas y 10 % hitos. Las tareas resumen suman a sus hijos (también los inactivos) y a
    veces tienen costo propio o una línea base menor que la suma (no acumulada).
    """
    tasks: list[Task] = []

    def add(level: int, parent_code: str) -> None:
        for i in range(1, (rng.randint(1, 3) if level <= 5 else 0) + 1):
            code = f"{parent_code}.{i}" if parent_code else str(i)
            start = datetime(2025, 1, 1, 8) + timedelta(days=rng.randint(0, 900))
            tasks.append(T(len(tasks) + 1, level, "" if rng.random() < 0.2 else code, f"T{len(tasks) + 1}",
                           round(rng.uniform(0, 1000), 2) if rng.random() < 0.85 else 0.0,
                           bs=start, bf=start + timedelta(days=rng.randint(0, 400)),
                           milestone=rng.random() < 0.1, active=rng.random() < 0.95,
                           company=rng.choice(["", "EmpresaA", "EmpresaB", "Ambas", "Otra"]) if multi else "",
                           split_a=rng.choice([None, 0, 30, 150])))
            add(level + 1, code)

    add(1, "")
    build_tree(tasks)

    def roll(t: Task) -> float:
        if not t.children:
            return t.baseline_cost
        total = sum(roll(c) for c in t.children)
        r = rng.random()
        t.baseline_cost = t.cost = total + (rng.uniform(1, 300) if r < 0.15 else 0) - (rng.uniform(1, 50) if r > 0.95 else 0)
        return t.baseline_cost

    for t in tasks:
        if t.parent is None:
            roll(t)
    return tasks


def expected_bac(tasks: list[Task], max_level: int) -> float:
    """BAC esperado, calculado de forma independiente del informe (sin planificación).

    Cuenta el costo propio de cada tarea activa que tiene código EDT de nivel <= max_level o
    depende de una; el costo propio de una tarea resumen es lo que su línea base supera a la
    suma de sus hijos (incluidos los inactivos, que no se cuentan).
    """
    roots = build_tree(tasks)
    inactive: dict[int, float] = {}

    def active(nodes: list[Task], parent: Task | None) -> list[Task]:
        kept = []
        for n in nodes:
            if not n.active:
                if parent is not None:
                    inactive[parent.uid] = inactive.get(parent.uid, 0.0) + n.baseline_cost
                continue
            n.children = active(n.children, n)
            kept.append(n)
        return kept

    roots = active(roots, None)
    own: dict[int, float] = {}

    def effective(t: Task) -> float:
        kids = sum(effective(c) for c in t.children)
        if not t.children and t.uid not in inactive:
            own[t.uid] = t.baseline_cost
            return t.baseline_cost
        residual = t.baseline_cost - kids - inactive.get(t.uid, 0.0)
        own[t.uid] = residual if residual > tolerance(t.baseline_cost) else 0.0
        return kids + own[t.uid]

    def counted(t: Task, covered: bool) -> float:
        covered = covered or (bool(t.wbs) and code_level(t.wbs) <= max_level)
        return (own[t.uid] if covered else 0.0) + sum(counted(c, covered) for c in t.children)

    for r in roots:
        effective(r)
    return sum(counted(r, False) for r in roots)


def assert_report_invariants(rep: Report, non_negative: bool = True) -> None:
    """Invariantes del informe: meses + PV posterior = BAC, subtotales = suma de hijos, una fila por tarea.

    Con costos no negativos y multi-empresa, además 0 <= BAC de la empresa A <= BAC.
    """
    months = rep.horizon.months
    for row in rep.rows:
        assert sum(row.vector) + row.pv_after == pytest.approx(row.bac, abs=1e-6), row.code
        if row.kind == KIND_SUMMARY:
            assert row.children, row.code
            assert row.bac == pytest.approx(sum(c.bac for c in row.children), abs=1e-6), row.code
            assert row.vector == pytest.approx([sum(c.vector[i] for c in row.children) for i in range(months)],
                                               abs=1e-6), row.code
        else:
            assert not row.children, row.code
        if row.parent is not None:
            assert row.level == row.parent.level + 1, row.code
        if rep.multi and non_negative:
            assert -1e-6 <= row.bac_a <= row.bac + 1e-6, row.code
    assert sum(rep.total_vector) + rep.pv_after == pytest.approx(rep.total_bac, abs=1e-6)
    assert rep.total_bac == pytest.approx(sum(r.bac for r in rep.roots), abs=1e-6)
    uids = [r.task.uid for r in rep.rows if r.kind != KIND_RESIDUAL]
    assert len(uids) == len(set(uids)), "una tarea no debe tener dos filas"
