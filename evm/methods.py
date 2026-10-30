"""Métodos de medición del valor ganado (EV): interpretación del texto de la columna C."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from evm.model import Task

FIXED = "formula_fija"        # 0/100, 50/50, 25/75, 20/80 ...
PCT = "pct_completado"        # % completado
PHYSICAL = "pct_fisico"       # % completado físico
UNITS = "unidades"            # unidades completadas (se registra como % completado físico)
MILESTONES = "hitos"          # hitos ponderados
LOE = "loe"                   # avance lineal / nivel de esfuerzo: EV = PV
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


def is_undefined(text: str) -> bool:
    """Método sin definir (vacío, «Por definir», «Pendiente»...): MET-005."""
    return parse_method(text).kind == UNDEFINED


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


def is_started(task: Task) -> bool:
    return task.actual_start is not None or task.percent_complete > 0 or task.physical_percent_complete > 0


def is_complete(task: Task) -> bool:
    return task.percent_complete >= 100 or (task.actual_finish is not None and task.percent_complete >= 99.5)
