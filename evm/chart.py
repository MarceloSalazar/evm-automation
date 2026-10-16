"""Comando 'curva': imagen PNG de la curva S e indicadores a partir del libro."""
from __future__ import annotations

from datetime import date
from itertools import accumulate
from pathlib import Path

from evm.indicators import Indicators, compute
from evm.update import BookData, book_totals, first_loaded_month


def last_data_index(values: list[float]) -> int:
    nz = [i for i, v in enumerate(values) if abs(v) > 0.005]
    return nz[-1] if nz else 0


def indicators_for(book: BookData, status: date) -> Indicators:
    pv, ev, ac, bac = book_totals(book)
    idx = book.horizon.index(status)
    if not 0 <= idx < book.horizon.months:
        raise ValueError(f"La fecha de corte {status:%d/%m/%Y} está fuera del horizonte del libro.")
    return compute(pv, ev, ac, bac, idx)


def plot_s_curve(book: BookData, status: date | None, out_path: str | Path, title: str = "") -> Path:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.dates as mdates
        import matplotlib.pyplot as plt
        from matplotlib.ticker import FuncFormatter
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Para generar la imagen instale matplotlib:  pip install matplotlib") from exc

    pv, ev, ac, bac = book_totals(book)
    n = book.horizon.months
    idx = book.horizon.index(status) if status else -1
    end = min(n, max(last_data_index(pv), idx) + 1 + 6)
    end = max(end, min(n, 12))
    # Cada valor acumulado se grafica al cierre de su mes; la curva parte de 0
    # al inicio del mes 1.
    x0 = book.horizon.start
    dates = [x0] + [book.horizon.month_end(i) for i in range(end)]
    cum_pv = [0.0] + list(accumulate(pv))[:end]
    few = end <= 24
    marker = "o" if few else None

    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=120)
    ax.plot(dates, cum_pv, color="#1F4E79", linewidth=2.2, marker=marker, markersize=3,
            label="PV acumulado (plan)")
    ax.axhline(bac, color="#7F7F7F", linestyle="--", linewidth=1.2, label=f"BAC {bac:,.0f}")
    subtitle = ""
    if status and 0 <= idx < n:
        k = min(idx, end - 1)
        # Cada curva real comienza en el primer mes con datos cargados
        for series, rows, color, label in ((ev, book.ev, "#70AD47", "EV acumulado (ganado)"),
                                           (ac, book.ac, "#C00000", "AC acumulado (costo real)")):
            first = first_loaded_month(rows)
            if first is None or first > k:
                continue
            cum = list(accumulate(series))
            xs = ([x0] if first == 0 else []) + [book.horizon.month_end(i) for i in range(first, k)] + [status]
            ys = ([0.0] if first == 0 else []) + cum[first: k + 1]
            ax.plot(xs, ys, color=color, linewidth=2.2, marker="o", markersize=3, label=label)
        ax.axvline(status, color="#404040", linestyle=":", linewidth=1)
        ax.text(status, 0.02, f" Corte {status:%d/%m/%Y}", transform=ax.get_xaxis_transform(),
                fontsize=9, ha="left", va="bottom", color="#404040")
        ind = compute(pv, ev, ac, bac, idx)
        spi = "-" if ind.spi is None else f"{ind.spi:.2f}"
        cpi = "-" if ind.cpi is None else f"{ind.cpi:.2f}"
        subtitle = (f"PV {ind.pv:,.0f}  |  EV {ind.ev:,.0f}  |  AC {ind.ac:,.0f}  |  SPI {spi}  |  CPI {cpi}"
                    f"  |  SPI(t) {'-' if ind.spi_t is None else f'{ind.spi_t:.2f}'}")
    name = title or str(book.meta.get("proyecto") or "")
    ax.set_title(f"Curva S - {name}" + (f"\n{subtitle}" if subtitle else ""), fontsize=12)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=6, maxticks=16))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%Y"))
    currency = str(book.meta.get("moneda") or "")
    ax.set_ylabel(f"Monto acumulado {currency}".strip())
    ax.set_ylim(bottom=0)
    ax.set_xlim(left=x0)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left")
    fig.autofmt_xdate()
    fig.tight_layout()
    out_path = Path(out_path)
    fig.savefig(out_path)
    plt.close(fig)
    return out_path
