"""Multi-empresa: parte de cada empresa en cada tarea (para su % planificado y su % ganado).

Cada tarea tiene una proporción de la empresa A entre 0 y 1 (la de B es el
complemento), según la columna "Empresa" de MS Project:

- Empresa A: 1. Empresa B: 0.
- "Ambas": el valor de la columna "% Reparto" (% de la empresa A, mayor que 0 y
  hasta 100) o el reparto por defecto si está vacío, es 0 o está fuera de rango.
  En MS Project un campo numérico vacío vale 0, por eso 0 equivale a vacío.
- Sin valor: se hereda de la tarea superior (empresa y % Reparto). Si ninguna
  tarea superior tiene empresa, se trata como "Ambas" con el reparto por defecto.
- Valor no reconocido: se trata como "Ambas" con el reparto por defecto.
"""
from __future__ import annotations

from evm.config import Config
from evm.fields import same_name
from evm.model import Task

def _split_value(task: Task, cfg: Config) -> tuple[float, bool]:
    """Proporción de A para una tarea "Ambas" y si el valor de % Reparto era inválido."""
    v = task.split_a
    if v is None or v == 0:
        return cfg.split_default / 100.0, False
    if not 0 <= v <= 100:
        return cfg.split_default / 100.0, True
    return v / 100.0, False


def explicit_share(task: Task, cfg: Config) -> tuple[float | None, str]:
    """Proporción de A definida en la propia tarea (None si no tiene empresa) y una marca:
    '' (normal), 'desconocida' (valor no reconocido) o 'reparto' (% Reparto fuera de rango)."""
    value = (task.company or "").strip()
    if not value:
        return None, ""
    if same_name(value, cfg.company_a):
        return 1.0, ""
    if same_name(value, cfg.company_b):
        return 0.0, ""
    if same_name(value, cfg.company_both):
        share, bad = _split_value(task, cfg)
        return share, "reparto" if bad else ""
    return cfg.split_default / 100.0, "desconocida"


def resolve_shares(roots: list[Task], cfg: Config) -> tuple[dict[int, float], dict[str, list[Task]]]:
    """Proporción de la empresa A en cada tarea de las ramas indicadas.

    Las raíces pueden heredar la empresa de sus tareas superiores (p. ej. la
    tarea de la planificación). Devuelve también las tareas observadas por tipo:
    'sin_empresa', 'desconocida' y 'reparto'.
    """
    shares: dict[int, float] = {}
    notes: dict[str, list[Task]] = {"sin_empresa": [], "desconocida": [], "reparto": []}

    def inherited_from_ancestors(t: Task) -> float | None:
        p = t.parent
        while p is not None:
            share, _ = explicit_share(p, cfg)
            if share is not None:
                return share
            p = p.parent
        return None

    def visit(t: Task, inherited: float | None) -> None:
        share, mark = explicit_share(t, cfg)
        if mark:
            notes[mark].append(t)
        if share is None:
            if inherited is None:
                notes["sin_empresa"].append(t)
                share = cfg.split_default / 100.0
                child_inherited = None
            else:
                share = inherited
                child_inherited = inherited
        else:
            child_inherited = share
        shares[t.uid] = share
        for c in t.children:
            visit(c, child_inherited)

    for r in roots:
        visit(r, inherited_from_ancestors(r))
    return shares, notes
