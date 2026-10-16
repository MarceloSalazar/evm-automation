"""Métodos de medición del valor ganado (EV) y su cálculo."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from evm.model import Task

FIXED = "formula_fija"        # 0/100, 50/50, 25/75, 20/80 ...
PCT = "pct_completado"        # % completado (ponderado por costo)
PHYSICAL = "pct_fisico"       # % completado físico
UNITS = "unidades"            # unidades completadas (se registra como % completado físico)
MILESTONES = "hitos"          # hitos ponderados
LOE = "loe"                   # avance lineal / nivel de esfuerzo: EV = PV
PROJECT_BCWP = "cptr"         # valor acumulado (CPTR/BCWP) calculado por MS Project
UNDEFINED = "por_definir"
UNKNOWN = "desconocido"


@dataclass(frozen=True)
class Method:
    kind: str
    start_fraction: float = 0.0   # solo para fórmula fija (p. ej. 0.5 en 50/50)
    text: str = ""


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().lower()


def parse_method(text: str) -> Method:
    """Interpreta el texto de la columna C (o del campo de MS Project)."""
    raw = (text or "").strip()
    t = _normalize(raw)
    if t in ("", "-", "por definir", "a definir", "pendiente", "n/a"):
        return Method(UNDEFINED, text=raw)
    m = re.fullmatch(r"(\d{1,3})\s*%?\s*[/\-–]\s*(\d{1,3})\s*%?", t)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if a + b == 100:
            return Method(FIXED, a / 100.0, raw)
        return Method(UNKNOWN, text=raw)
    if "cptr" in t or "bcwp" in t or "ms project" in t or "valor acumulado" in t:
        return Method(PROJECT_BCWP, text=raw)
    if "hito" in t or "milestone" in t:
        return Method(MILESTONES, text=raw)
    if "fisic" in t or "physical" in t:
        return Method(PHYSICAL, text=raw)
    if "unidad" in t or "unit" in t:
        return Method(UNITS, text=raw)
    if "prorrate" in t or "apportion" in t:
        return Method(UNKNOWN, text=raw)
    if "lineal" in t or "linear" in t or "loe" in t or "esfuerzo" in t or "level of effort" in t:
        return Method(LOE, text=raw)
    if "%" in t or "porcentaje" in t or "percent" in t or "completado" in t or "avance" in t:
        return Method(PCT, text=raw)
    return Method(UNKNOWN, text=raw)


def _leaves(task: Task) -> list[Task]:
    if not task.children:
        return [task]
    out: list[Task] = []
    for c in task.children:
        out.extend(_leaves(c))
    return out


def _descendants(task: Task):
    for c in task.children:
        yield c
        yield from _descendants(c)


def is_started(task: Task) -> bool:
    return task.actual_start is not None or task.percent_complete > 0 or task.physical_percent_complete > 0


def is_complete(task: Task) -> bool:
    return task.percent_complete >= 100 or (task.actual_finish is not None and task.percent_complete >= 99.5)


def weighted_percent(task: Task, physical: bool = False) -> float:
    """% de avance ponderado por el costo de línea base de las tareas de detalle (0..1)."""
    leaves = _leaves(task)
    pct = (lambda t: t.physical_percent_complete) if physical else (lambda t: t.percent_complete)

    def clamp(v: float) -> float:
        return max(0.0, min(1.0, v / 100.0))

    if len(leaves) == 1:
        return clamp(pct(leaves[0]))
    total = sum(max(t.baseline_cost, 0.0) for t in leaves)
    if total <= 0:
        # Sin costos para ponderar: % completado de MS Project de la tarea (acumulado por
        # duración); el % físico no se acumula en las tareas resumen, se promedia.
        return clamp(sum(pct(t) for t in leaves) / len(leaves)) if physical else clamp(pct(task))
    return clamp(sum(max(t.baseline_cost, 0.0) * pct(t) for t in leaves) / total)


def ev_fraction(method: Method, task: Task, pv_fraction: float, bac: float | None = None,
                bcwp: float | None = None, bcwp_base: float | None = None) -> tuple[float, str]:
    """Fracción del BAC de la fila ganada a la fecha de corte y una nota explicativa.

    pv_fraction: PV acumulado / BAC a la fecha de corte (para avance lineal/LOE).
    bac, bcwp:   BAC de la fila y CPTR (BCWP) de MS Project que le corresponde
                 (para el método "CPTR de MS Project").
    bcwp_base:   costo de la línea base con que MS Project calculó el CPTR, si es
                 distinta de la línea base del informe (fracción = CPTR / bcwp_base).
    """
    if method.kind == FIXED:
        if is_complete(task):
            return 1.0, ""
        if is_started(task):
            return method.start_fraction, ""
        return 0.0, ""
    if method.kind == PCT:
        return weighted_percent(task), ""
    if method.kind in (PHYSICAL, UNITS):
        return weighted_percent(task, physical=True), ""
    if method.kind == PROJECT_BCWP:
        value = bcwp if bcwp is not None else task.bcwp
        base = bcwp_base if bcwp_base is not None else (bac if bac is not None else task.baseline_cost)
        if value is not None and base and base > 0:
            note = "el CPTR de MS Project supera el presupuesto: se limitó al BAC" if value > base + 0.5 else ""
            return max(0.0, min(1.0, value / base)), note
        if value is not None and not base:
            return 0.0, ""
        return weighted_percent(task), "MS Project no informó CPTR: se usó % completado"
    if method.kind == LOE:
        return max(0.0, min(1.0, pv_fraction)), ""
    if method.kind == MILESTONES:
        milestones = [d for d in _descendants(task) if d.milestone]
        if milestones:
            weights = [(d, d.milestone_weight) for d in milestones]
            if all(w is None or w == 0 for _, w in weights):
                weights = [(d, 1.0) for d in milestones]
                note = "hitos sin peso definido: se ponderaron por igual"
            else:
                weights = [(d, float(w or 0.0)) for d, w in weights]
                note = ""
            total = sum(w for _, w in weights)
            if total > 0:
                done = sum(w for d, w in weights if is_complete(d))
                return done / total, note
        if task.physical_percent_complete > 0 or any(t.physical_percent_complete > 0 for t in _leaves(task)):
            return weighted_percent(task, physical=True), "sin hitos: se usó % físico completado"
        return weighted_percent(task), "sin hitos ni % físico completado: se usó % completado"
    # Por definir / no reconocido
    note = "método no definido" if method.kind == UNDEFINED else f"método no reconocido '{method.text}'"
    return weighted_percent(task), note + ": se usó % completado"
