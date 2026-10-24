"""Constructores compartidos para las pruebas: tareas, proyectos, informes y árboles aleatorios.

Uso típico:

    tasks = tree('''
        1 Fase
          1.1 A
          - Suelto
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
from evm.wbs import KIND_RESIDUAL, KIND_SUMMARY, Report, build_report, build_tree

# Fechas de línea base por defecto: tres meses completos (enero a marzo de 2026)
S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)


def seeds(default: int) -> range:
    """Semillas de las pruebas aleatorias; la variable de entorno EVM_SEMILLAS permite ampliar la cantidad."""
    return range(int(os.environ.get("EVM_SEMILLAS", default)))


# ----------------------------------------------------------------------------
# Tareas y proyectos
# ----------------------------------------------------------------------------
def T(uid: int, level: int, wbs: str, name: str, *, bs: datetime | None = S, bf: datetime | None = F,
      start: datetime | None = S, finish: datetime | None = F, **kw) -> Task:
    """Tarea con fechas de línea base (`bs`-`bf`) y del programa actual (`start`-`finish`)."""
    return Task(uid=uid, id=uid, name=name, outline_level=level, wbs=wbs, baseline_start=bs, baseline_finish=bf,
                has_baseline=bs is not None, start=start, finish=finish, **kw)


def P(tasks: list[Task], **kw) -> ProjectData:
    return ProjectData(name="Programa", source="p.xml", backend="xml", tasks=tasks, **kw)


def tree(text: str, **common) -> list[Task]:
    """Tareas a partir de un esquema con sangría de 2 espacios por nivel.

    Cada línea es "<código EDT> <nombre>"; el código "-" indica una tarea sin código EDT y el nombre puede tener
    espacios. Los argumentos adicionales se aplican a todas las tareas.
    """
    tasks = []
    for line in textwrap.dedent(text).strip("\n").splitlines():
        if not line.strip():
            continue
        code, *name = line.split()
        level = (len(line) - len(line.lstrip(" "))) // 2 + 1
        tasks.append(T(len(tasks) + 1, level, "" if code == "-" else code, " ".join(name), **common))
    return tasks


def by_name(tasks: list[Task]) -> dict[str, Task]:
    return {t.name: t for t in tasks}


# ----------------------------------------------------------------------------
# Informes
# ----------------------------------------------------------------------------
def report(tasks: list[Task], cfg: Config | None = None, **project_kw) -> Report:
    return build_report(P(tasks, **project_kw), cfg or Config())


def rows_by_code(rep: Report) -> dict:
    return {r.code: r for r in rep.rows}


def issues_with(rep: Report, text: str, severity: str | None = None) -> list:
    """Observaciones de la hoja Validación que contienen `text` (opcionalmente, de una severidad)."""
    return [i for i in rep.issues if text in i.message and (severity is None or i.severity == severity)]


BAC = 1_000_000.0      # BAC de cada fila en earned_after_update: el EV acumulado / BAC es el % ganado


def set_bac(path, value: float = BAC) -> dict[str, int]:
    """Ingresa el mismo BAC en todas las filas de ingreso (hoja PV o las de cada empresa) y devuelve las filas del
    libro por código."""
    from openpyxl import load_workbook

    from evm.workbook import read_meta, read_structure
    wb = load_workbook(path)
    meta = read_meta(wb)
    sheets = ([f"PV-{meta['empresa_a']}", f"PV-{meta['empresa_b']}"] if meta.get("multiempresa") else ["PV"])
    rows = {}
    for s in read_structure(wb):
        rows[s.code] = s.row
        if s.is_input:
            for sheet in sheets:
                wb[sheet].cell(row=s.row, column=4, value=value)
    wb.save(path)
    return rows


def earned_after_update(tasks: list[Task], cfg: Config, folder, status: date) -> dict[str, float]:
    """Genera el libro (BAC 1.000.000 en cada fila y empresa), simula `actualizar` a la fecha de corte y devuelve el
    % ganado acumulado de cada fila (en multi-empresa, el promedio de las dos empresas)."""
    from evm.update import run_update
    from evm.workbook import write_workbook
    book = folder / "libro.xlsx"
    write_workbook(build_report(P(tasks), cfg), cfg, book)
    set_bac(book)
    res = run_update(P(tasks), cfg, book, status, backup=False, dry_run=True)
    roles = 2 if cfg.multi_enabled else 1
    return {r.code: r.ev_cum / (BAC * roles) for r in res.rows}


# ----------------------------------------------------------------------------
# Árboles aleatorios e invariantes
# ----------------------------------------------------------------------------
def random_tree(rng: random.Random, multi: bool = False) -> list[Task]:
    """EDT aleatoria.

    Hasta 6 niveles de esquema y 1 a 3 hijos por tarea; 20 % de tareas sin código EDT, 5 % inactivas, 10 % hitos
    (la mitad con peso), 5 % sin fechas de línea base y Visible vacío, Sí o No.
    """
    tasks: list[Task] = []

    def add(level: int, parent_code: str) -> None:
        for i in range(1, (rng.randint(1, 3) if level <= 5 else 0) + 1):
            code = f"{parent_code}.{i}" if parent_code else str(i)
            start = datetime(2025, 1, 1, 8) + timedelta(days=rng.randint(0, 900))
            milestone = rng.random() < 0.1
            no_baseline = rng.random() < 0.05
            tasks.append(T(len(tasks) + 1, level, "" if rng.random() < 0.2 else code, f"T{len(tasks) + 1}",
                           bs=None if no_baseline else start,
                           bf=None if no_baseline else start + timedelta(days=0 if milestone else rng.randint(0, 400)),
                           milestone=milestone,
                           milestone_weight=rng.choice([None, 10.0, 30.0]) if milestone else None,
                           active=rng.random() < 0.95,
                           visible=rng.choice(["", "", "", "", "Sí", "No"]),
                           company=rng.choice(["", "EmpresaA", "EmpresaB", "Ambas", "Otra"]) if multi else "",
                           split_a=rng.choice([None, 0, 30, 150])))
            add(level + 1, code)

    add(1, "")
    build_tree(tasks)
    return tasks


def assert_report_invariants(rep: Report) -> None:
    """Invariantes del informe: el % planificado de cada fila de ingreso (y empresa) está entre 0 y 100 %, y es 100 %
    si la fila tiene alguna tarea con peso y todo su % cae dentro del horizonte; los subtotales suman a sus hijos;
    una fila por tarea."""
    months = rep.horizon.months
    for row in rep.rows:
        if row.is_input:
            for role, vec in row.pct_vec.items():
                assert min(vec, default=0.0) >= -1e-9, (row.code, role)
                planned = sum(row.pct[role].values())
                assert planned == pytest.approx(1.0) or planned == pytest.approx(0.0), (row.code, role, planned)
                # lo anterior al mes 1 se acumula en el mes 1; lo posterior al horizonte no aparece
                after = sum(v for k, v in row.pct[role].items() if rep.horizon.index(k) >= months)
                assert sum(vec) == pytest.approx(planned - after, abs=1e-9), (row.code, role)
        if row.kind == KIND_SUMMARY:
            assert row.children, row.code
            assert row.vector == pytest.approx([sum(c.vector[i] for c in row.children) for i in range(months)],
                                               abs=1e-9), row.code
        else:
            assert not row.children, row.code
        if row.parent is not None:
            assert row.level == row.parent.level + 1, row.code
    assert rep.total_vector == pytest.approx([sum(r.vector[i] for r in rep.roots) for i in range(months)], abs=1e-9)
    uids = [r.task.uid for r in rep.rows if r.kind != KIND_RESIDUAL]
    assert len(uids) == len(set(uids)), "una tarea no debe tener dos filas"
