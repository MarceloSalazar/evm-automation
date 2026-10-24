"""Horizonte mensual y utilidades de meses."""
from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date

from evm.model import MonthKey


def first_of_month(d: date) -> date:
    return date(d.year, d.month, 1)


def add_months(d: date, n: int) -> date:
    """Primer día del mes que está n meses después del mes de d."""
    y, m = divmod(d.month - 1 + n, 12)
    return date(d.year + y, m + 1, 1)


def last_of_month(d: date) -> date:
    return date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


def month_key(d: date) -> MonthKey:
    return (d.year, d.month)


@dataclass(frozen=True)
class Horizon:
    """Horizonte de meses: el mes 1 empieza en `start` (primer día del mes)."""

    start: date
    months: int

    def index(self, value: date | MonthKey) -> int:
        """Índice base 0 del mes (puede quedar fuera de 0..months-1)."""
        y, m = value if isinstance(value, tuple) else (value.year, value.month)
        return (y - self.start.year) * 12 + (m - self.start.month)

    def month_start(self, i: int) -> date:
        return add_months(self.start, i)

    def month_end(self, i: int) -> date:
        return last_of_month(self.month_start(i))

    def key(self, i: int) -> MonthKey:
        return month_key(self.month_start(i))

    @property
    def end(self) -> date:
        return self.month_end(self.months - 1)

    def to_vector(self, values: dict[MonthKey, float]) -> tuple[list[float], float, float]:
        """Convierte {mes: valor} en (vector de `months` valores, previo, posterior)."""
        vec = [0.0] * self.months
        before = after = 0.0
        for k, v in values.items():
            i = self.index(k)
            if i < 0:
                before += v
            elif i >= self.months:
                after += v
            else:
                vec[i] += v
        return vec, before, after


def add_into(target: dict[MonthKey, float], values: dict[MonthKey, float], factor: float = 1.0) -> None:
    for k, v in values.items():
        target[k] = target.get(k, 0.0) + v * factor


def round_preserving_total(values: list[float], decimals: int = 2) -> list[float]:
    """Redondea cada valor sin alterar el total (el residuo va al mayor valor)."""
    if not values:
        return values
    rounded = [round(v, decimals) for v in values]
    diff = round(round(sum(values), decimals) - sum(rounded), decimals)
    if diff:
        i = max(range(len(values)), key=lambda j: abs(values[j]))
        rounded[i] = round(rounded[i] + diff, decimals)
    return rounded
