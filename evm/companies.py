"""Multi-empresa: reparto del presupuesto, el avance y el costo real entre dos empresas.

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
from evm.methods import FIXED, LOE, MILESTONES, PHYSICAL, PROJECT_BCWP, UNITS, Method
from evm.model import Task

# Claves de las series por empresa (hojas EV-A, AC-B, etc.): A y B por rol, no por nombre.
ROLE_A, ROLE_B = "A", "B"


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


# ----------------------------------------------------------------------------
# Valor ganado y costo real por empresa
# ----------------------------------------------------------------------------
def _own(task: Task, value) -> float | None:
    """Valor propio de una tarea (sin sus subtareas) para un campo acumulado por MS Project."""
    total = value(task)
    if total is None:
        return None
    if not task.children:
        return total
    kids = [value(c) for c in task.children]
    if any(k is None for k in kids):
        return None
    return total - sum(kids)


def _weighted(leaves: list[Task], weight, pct) -> float | None:
    total = sum(weight(t) for t in leaves)
    if total <= 0:
        return None
    return max(0.0, min(1.0, sum(weight(t) * pct(t) for t in leaves) / total / 100.0))


def ev_by_company(method: Method, contributors: list[Task], shares: dict[int, float],
                  bac: tuple[float, float], pv_cum: tuple[float, float], total_fraction: float,
                  ev_baseline_mismatch: bool = False) -> tuple[float, float]:
    """EV acumulado de las empresas A y B en una fila.

    contributors: tareas cuyo costo forma la fila (rama del paquete, o costo propio
    y tareas sin fila de una tarea resumen). Para cada empresa se usa su BAC en la
    fila y su avance, calculado con sus tareas ponderadas por el BAC que le
    corresponde en cada una:
    - % completado, % físico, unidades: avance de las tareas de cada empresa.
    - CPTR de MS Project: CPTR de las tareas de cada empresa.
    - Avance lineal (LOE): PV acumulado de cada empresa.
    - 0/100, X/Y e hitos ponderados: el avance del paquete completo (total_fraction).
    """
    out = []
    for role, bac_c, pv_c in ((ROLE_A, bac[0], pv_cum[0]), (ROLE_B, bac[1], pv_cum[1])):
        if bac_c <= 0:
            out.append(0.0)
            continue

        def s(t: Task, role=role) -> float:
            a = shares.get(t.uid, 0.5)
            return a if role == ROLE_A else 1.0 - a

        kind = method.kind
        if kind in (FIXED, MILESTONES):
            out.append(bac_c * total_fraction)
            continue
        if kind == LOE:
            out.append(max(0.0, min(bac_c, pv_c)))
            continue
        leaves = [t for t in contributors if not t.children]
        if kind == PROJECT_BCWP:
            own_bcwp = [(t, _own(t, lambda x: x.bcwp)) for t in contributors]
            if all(v is not None for _, v in own_bcwp):
                value = sum(v * s(t) for t, v in own_bcwp)
                if ev_baseline_mismatch:
                    base = sum((_own(t, lambda x: x.ev_baseline_cost or 0.0) or 0.0) * s(t) for t in contributors)
                    frac = max(0.0, min(1.0, value / base)) if base > 0 else total_fraction
                    out.append(bac_c * frac)
                else:
                    out.append(max(0.0, min(bac_c, value)))
                continue
            kind = None   # sin CPTR: % completado
        physical = kind in (PHYSICAL, UNITS)
        pct = (lambda t: t.physical_percent_complete) if physical else (lambda t: t.percent_complete)
        frac = _weighted(leaves, lambda t: max(t.baseline_cost, 0.0) * s(t), pct)
        out.append(bac_c * (total_fraction if frac is None else frac))
    return out[0], out[1]


def ac_by_company(contributors: list[Task], shares: dict[int, float]) -> tuple[float, float]:
    """Costo real (CRTR) de MS Project repartido entre A y B según las tareas de la fila."""
    def acwp(t: Task) -> float:
        return t.acwp if t.acwp is not None else t.actual_cost

    a = b = 0.0
    for t in contributors:
        own = _own(t, acwp) or 0.0
        share = shares.get(t.uid, 0.5)
        a += own * share
        b += own * (1.0 - share)
    return a, b
