"""Indicadores de Valor Ganado (mismas fórmulas que la hoja 'Resumen')."""
from __future__ import annotations

from dataclasses import dataclass
from itertools import accumulate

from evm.fields import fmt_number

EPS = 0.005


def _div(a: float | None, b: float | None) -> float | None:
    if a is None or b is None or abs(b) < 1e-12:
        return None
    return a / b


def planned_duration(cum_pv: list[float], bac: float) -> float | None:
    """Meses del horizonte hasta completar el BAC (None si no se completa dentro del horizonte)."""
    if not cum_pv or cum_pv[-1] < bac - EPS:
        return None
    return float(sum(1 for v in cum_pv if v < bac - EPS) + 1)


def earned_schedule(cum_pv: list[float], ev: float, bac: float, at: int | None = None) -> float | None:
    """Cronograma ganado (ES) en meses, interpolando linealmente en la curva PV.

    Si el EV coincide con un tramo plano de la curva PV (meses sin PV), cualquier
    momento del tramo es válido: se toma el más cercano al tiempo transcurrido AT.
    Devuelve None si la curva PV no alcanza el EV dentro del horizonte y 0 si el EV es negativo
    (correcciones cargadas a mano).
    """
    if ev < -EPS:
        return 0.0
    if ev >= bac - EPS:
        return planned_duration(cum_pv, bac)
    lo = sum(1 for v in cum_pv if v < ev - EPS)     # meses con PV acumulado < EV
    hi = sum(1 for v in cum_pv if v <= ev + EPS)    # meses con PV acumulado <= EV (con tolerancia)
    if hi > lo:
        lower = lo + 1 if ev > 0 else 0
        return float(min(max(at, lower), hi)) if at is not None else float(lower)
    c = hi
    if c >= len(cum_pv):
        return None
    pv_c = cum_pv[c - 1] if c > 0 else 0.0
    pv_next = cum_pv[c]
    if pv_next == pv_c:         # solo con una curva PV no creciente (PV mensual negativo)
        return float(c)
    return c + (ev - pv_c) / (pv_next - pv_c)


@dataclass
class Indicators:
    status_index: int          # índice base 0 del mes de corte
    bac: float
    pv: float
    ev: float
    ac: float
    sv: float
    cv: float
    spi: float | None
    cpi: float | None
    eac_cpi: float | None
    eac_remaining_budget: float
    eac_cpi_spi: float | None
    etc: float | None
    vac: float | None
    tcpi: float | None
    pct_planned: float | None
    pct_complete: float | None
    pct_spent: float | None
    planned_duration: float | None
    es: float | None
    at: int
    sv_t: float | None
    spi_t: float | None
    estimated_duration: float | None
    # Tendencia: CPI y SPI de los últimos 3 meses e indicadores acumulados del mes anterior
    cpi_3m: float | None = None
    spi_3m: float | None = None
    cpi_prev: float | None = None
    spi_prev: float | None = None
    eac_prev: float | None = None


def compute(pv_month: list[float], ev_month: list[float], ac_month: list[float], bac: float,
            status_index: int, loaded: list[bool] | None = None) -> Indicators:
    """Calcula los indicadores acumulados hasta el mes `status_index` (base 0) inclusive.

    loaded: meses con EV cargado (None = todos). El mes anterior es el último mes con datos antes del de corte, y
    los últimos 3 meses empiezan en el último mes con datos hasta 3 meses antes (como en la hoja Resumen).
    """
    cum_pv = list(accumulate(pv_month))
    k = status_index
    pv = cum_pv[k] if 0 <= k < len(cum_pv) else (cum_pv[-1] if k >= len(cum_pv) else 0.0)
    ev = sum(ev_month[: k + 1]) if k >= 0 else 0.0
    ac = sum(ac_month[: k + 1]) if k >= 0 else 0.0
    spi = _div(ev, pv)
    cpi = _div(ev, ac)
    eac_cpi = _div(bac, cpi)
    eac_cpi_spi = None
    if cpi and spi:
        eac_cpi_spi = ac + (bac - ev) / (cpi * spi)
    at = k + 1
    es = earned_schedule(cum_pv, ev, bac, at)
    spi_t = _div(es, at)
    pd = planned_duration(cum_pv, bac)

    def last_loaded(back: int) -> int | None:
        """Último mes con datos (índice) hasta `back` meses antes del mes de corte."""
        for i in range(k - back, -1, -1):
            if loaded is None or (i < len(loaded) and loaded[i]):
                return i
        return None

    def until(series: list[float], i: int | None) -> float | None:
        return None if i is None else sum(series[: i + 1])

    p, b = last_loaded(1), last_loaded(3)
    pv1, ev1, ac1 = until(pv_month, p), until(ev_month, p), until(ac_month, p)
    pv3, ev3, ac3 = until(pv_month, b), until(ev_month, b), until(ac_month, b)
    cpi_prev = _div(ev1, ac1)
    return Indicators(
        status_index=k, bac=bac, pv=pv, ev=ev, ac=ac, sv=ev - pv, cv=ev - ac, spi=spi, cpi=cpi,
        eac_cpi=eac_cpi, eac_remaining_budget=ac + bac - ev, eac_cpi_spi=eac_cpi_spi,
        etc=(eac_cpi - ac) if eac_cpi is not None else None,
        vac=(bac - eac_cpi) if eac_cpi is not None else None,
        tcpi=_div(bac - ev, bac - ac), pct_planned=_div(pv, bac), pct_complete=_div(ev, bac),
        pct_spent=_div(ac, bac), planned_duration=pd, es=es, at=at, sv_t=(es - at) if es is not None else None,
        spi_t=spi_t,
        estimated_duration=_div(pd, spi_t),
        cpi_3m=None if b is None else _div(ev - ev3, ac - ac3),
        spi_3m=None if b is None else _div(ev - ev3, pv - pv3), cpi_prev=cpi_prev,
        spi_prev=_div(ev1, pv1), eac_prev=_div(bac, cpi_prev),
    )


def format_table(ind: Indicators, currency: str = "") -> str:
    def m(v):
        return "-" if v is None else fmt_number(v, 0)

    def r(v):
        return "-" if v is None else fmt_number(v)

    def p(v):
        return "-" if v is None else f"{fmt_number(v * 100, 1)} %"

    cur = f" ({currency})" if currency else ""
    lines = [
        f"BAC{cur:<10}{m(ind.bac):>20}",
        f"PV acumulado{'':<3}{m(ind.pv):>20}   % planificado {p(ind.pct_planned)}",
        f"EV acumulado{'':<3}{m(ind.ev):>20}   % completado  {p(ind.pct_complete)}",
        f"AC acumulado{'':<3}{m(ind.ac):>20}   % gastado     {p(ind.pct_spent)}",
        f"SV = EV - PV{'':<3}{m(ind.sv):>20}   SPI = {r(ind.spi)}",
        f"CV = EV - AC{'':<3}{m(ind.cv):>20}   CPI = {r(ind.cpi)}",
        f"EAC = BAC/CPI{'':<2}{m(ind.eac_cpi):>20}   VAC = {m(ind.vac)}",
        f"TCPI{'':<11}{r(ind.tcpi):>20}",
        f"ES (meses){'':<5}{r(ind.es):>20}   AT = {ind.at}  SV(t) = {r(ind.sv_t)}  SPI(t) = {r(ind.spi_t)}",
        f"Últimos 3 meses{'':<16}CPI = {r(ind.cpi_3m)}  SPI = {r(ind.spi_3m)}",
        f"Mes anterior{'':<19}CPI = {r(ind.cpi_prev)}  SPI = {r(ind.spi_prev)}  EAC = {m(ind.eac_prev)}"
        + (f"  (variación del EAC: {_signed(ind.eac_cpi - ind.eac_prev)})"
           if ind.eac_cpi is not None and ind.eac_prev is not None else ""),
    ]
    return "\n".join(lines)


def _signed(v: float) -> str:
    return ("+" if v > 0 else "") + fmt_number(v, 0)
