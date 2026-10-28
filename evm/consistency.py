"""Controles de coherencia del libro en el cierre mensual (norma EIA-748-E).

- EV y AC del mismo período (guías 14 y 15): paquetes con avance y sin costo, con costo y sin avance, o con costo
  después de llegar al 100 %. Se revisan los valores acumulados hasta el último mes con AC cargado: después de
  'actualizar', el AC del mes de corte se suele cargar a mano y revisar con 'analizar'.
- Meses sin informe (guía 17): meses sin % ganado entre dos meses cargados.

Los usan 'actualizar' (con el % ganado que acaba de calcular) y 'analizar', con los valores que quedan en el libro.
"""
from __future__ import annotations

from datetime import date

from evm.fields import fmt_number
from evm.messages import Issue
from evm.methods import FIXED, MILESTONES, parse_method

EPS = 1e-9
EPS_MONEY = 0.005
DONE = 1 - 1e-6                         # % ganado acumulado a partir del que el paquete está terminado
SHOWN = 6                               # meses que se detallan en una observación


def _months(dates: list[date]) -> str:
    """'03/2027', '03/2027 y 04/2027', '03/2027, 04/2027 y 2 más'."""
    names = [f"{d:%m/%Y}" for d in dates[:SHOWN]]
    if len(dates) > SHOWN:
        return ", ".join(names) + f" y {len(dates) - SHOWN} más"
    return names[0] if len(names) == 1 else ", ".join(names[:-1]) + " y " + names[-1]


def sentence(prefix: str, text: str) -> str:
    """Texto de una observación: con el prefijo de la empresa ('EmpresaA: ...') o con mayúscula inicial."""
    return prefix + text if prefix else text[:1].upper() + text[1:]


def _pct(v: float) -> str:
    return f"{fmt_number(v * 100, 1, trim=True)} %"


def ev_ac_checks(book, idx: int) -> list[Issue]:
    """EV y AC de cada fila de ingreso (y empresa) hasta el último mes con AC cargado, como máximo el mes idx."""
    decimals = int(book.meta.get("decimales") or 0)
    roles = book.roles()
    loaded = [i for i in range(idx + 1)
              if any(book.ac_of(role)[s.row][i] is not None for s in book.input_rows for role, _c in roles)]
    cutoff = book.horizon.month_start(idx)
    if not loaded:
        return [Issue("ACT-042", message="Ningún mes tiene AC cargado: no se revisó la coherencia del EV y el AC. "
                                         "Después de cargar el AC en la hoja AC, se debe ejecutar analizar.")]
    until = loaded[-1]
    month = book.horizon.month_start(until)
    out: list[Issue] = []
    if until < idx:
        out.append(Issue("ACT-042", message=f"El AC de {cutoff:%m/%Y} no está cargado: la coherencia del EV y el AC "
                                            f"se revisó hasta {month:%m/%Y}. Después de cargarlo, se debe ejecutar "
                                            "analizar."))
    for s in book.input_rows:
        method = parse_method(book.method[s.row])
        # 0/100 e hitos ponderados: el EV llega al terminar o al cumplirse cada hito
        stepwise = (method.kind == FIXED and method.start_fraction <= 0) or method.kind == MILESTONES
        for role, company in roles:
            prefix = f"{company}: " if company else ""
            pct = [v or 0.0 for v in book.pct_ev[role][s.row][: until + 1]]
            ac = [v or 0.0 for v in book.ac_of(role)[s.row][: until + 1]]
            bac = book.bac_role[role][s.row]
            earned, spent = sum(pct), sum(ac)
            if bac * earned > EPS_MONEY and abs(spent) <= EPS_MONEY:
                out.append(Issue("ACT-039", s.code, s.name, sentence(
                    prefix, f"EV acumulado {fmt_number(bac * earned, decimals)} ({_pct(earned)} ganado) y sin AC a "
                            f"{month:%m/%Y}: el CPI no se puede calcular.")))
            elif spent > EPS_MONEY and earned <= EPS and not stepwise:
                out.append(Issue("ACT-040", s.code, s.name, sentence(
                    prefix, f"AC acumulado {fmt_number(spent, decimals)} y sin % ganado a {month:%m/%Y}.")))
            total = 0.0
            for i, v in enumerate(pct):
                total += v
                if total >= DONE:
                    late = [j for j in range(i + 1, until + 1) if abs(ac[j]) > EPS_MONEY]
                    if late:
                        out.append(Issue("ACT-041", s.code, s.name, sentence(
                            prefix, f"llegó al 100 % en {book.horizon.month_start(i):%m/%Y} y tiene AC en "
                                    f"{_months([book.horizon.month_start(j) for j in late])} "
                                    f"({fmt_number(sum(ac[j] for j in late), decimals)}).")))
                    break
    return out


def _loaded_ev(book, idx: int) -> list[int]:
    """Meses hasta idx con % ganado cargado en alguna fila de ingreso (también 0)."""
    return [i for i in range(idx + 1)
            if any(book.pct_ev[role][s.row][i] is not None for s in book.input_rows for role, _c in book.roles())]


def month_gaps(book, idx: int, latest_only: bool) -> list[Issue]:
    """Meses sin % ganado: entre el último mes cargado y el mes de corte (latest_only, 'actualizar') o entre el
    primer mes cargado y el mes de corte ('analizar')."""
    loaded = _loaded_ev(book, idx)
    if not loaded:
        return []
    before = [i for i in loaded if i < idx]
    if not before:
        return []
    start = before[-1] + 1 if latest_only else loaded[0]
    gaps = [i for i in range(start, idx) if i not in set(loaded)]
    if not gaps:
        return []
    dates = [book.horizon.month_start(i) for i in gaps]
    after = next((i for i in loaded if i > gaps[-1]), idx)
    where = (f"su avance quedó en {book.horizon.month_start(after):%m/%Y}" if latest_only or len(gaps) == 1
             else "su avance quedó en el mes cargado siguiente")
    return [Issue("ACT-046", message=f"{_months(dates)} no {'tiene' if len(gaps) == 1 else 'tienen'} % ganado "
                                     f"cargado: {where}. La norma EIA-748-E pide la información al menos cada mes: se "
                                     "recomienda ejecutar actualizar al cierre de cada mes.")]
