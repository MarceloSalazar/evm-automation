"""Barra de progreso de la consola: tarea en curso, % de la tarea y del total, tiempo transcurrido y tiempo restante
estimado (ETA).

    [██████░░░░░░░░░░░░░░]  31 %  Transcurrido 0:19  ·  ETA 0:42  ·  Lectura de MS Project: 1.234 de 3.000 tareas (41 %)

Cada comando declara sus etapas con un peso relativo (el tiempo que suele llevar cada una) y las partes del programa
informan su avance con stage() y step(), sin saber si hay una barra activa. El % total combina las etapas
terminadas y el avance de la etapa en curso.

ETA = lo que falta de la etapa en curso + el peso de las etapas siguientes × el ritmo medido (segundos por unidad de
peso). Una etapa que informa su avance (tareas, filas, hojas) estima lo que le falta a su propio ritmo; una etapa
sin conteo (apertura de MS Project, lectura o guardado del libro) cuenta hacia atrás su duración esperada, y entre
dos avances informados también se cuenta hacia atrás. Así, una etapa más lenta o más rápida que su peso no hace
subir y después bajar el ETA de todo el comando.

Un hilo vuelve a dibujar la barra cada REFRESH segundos: el tiempo transcurrido y el ETA avanzan también en las
etapas sin conteo. La barra se escribe en la salida de errores y solo si es una consola interactiva: con la salida
redirigida (archivos de registro, pruebas, instalador) no se muestra nada. Con -v, al terminar se registra la duración
de cada etapa.
"""
from __future__ import annotations

import logging
import shutil
import sys
import threading
import time
from collections.abc import Callable
from contextlib import contextmanager

from evm.fields import fmt_number

log = logging.getLogger(__name__)

BAR_WIDTH = 20
REDRAW = 0.1            # segundos mínimos entre dos dibujos de la barra
REFRESH = 0.5           # segundos entre dos dibujos del hilo de refresco
MIN_ETA_FRACTION = 0.1  # el ETA se muestra cuando se midió este % del total (etapas terminadas o con conteo)
MIN_ETA_SECONDS = 1.0
STAGE_TRUST = 0.1       # desde este % de la etapa en curso, lo que le falta se estima solo con su propio ritmo


def _interactive(stream) -> bool:
    try:
        return bool(stream.isatty())
    except (AttributeError, ValueError):
        return False


def clock_text(seconds: float) -> str:
    """'0:42', '12:05', '1:02:03'."""
    s = int(round(seconds))
    h, rest = divmod(s, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


class Progress:
    """Progreso de un comando con etapas de peso conocido."""

    def __init__(self, stages: dict[str, float], stream=None, enabled: bool | None = None,
                 clock: Callable[[], float] = time.monotonic):
        self.stream = stream if stream is not None else sys.stderr
        self.enabled = _interactive(self.stream) if enabled is None else enabled
        total = sum(stages.values()) or 1.0
        self.weights = {k: v / total for k, v in stages.items()}
        self.clock = clock
        self.started = clock()
        self.done = 0.0                 # peso de las etapas terminadas
        self.done_seconds = 0.0         # duración de las etapas terminadas (con peso)
        self.closed: dict[str, float] = {}  # etapa terminada -> su duración
        self.key: str | None = None
        self.label = ""
        self.fraction = 0.0             # avance de la etapa en curso (0 a 1)
        self.count: tuple[int, int, str] | None = None
        self.stage_started = self.started
        self.reported_at = self.started # último avance informado de la etapa en curso (o su comienzo)
        self.width = 0                  # caracteres de la última línea dibujada
        self.drawn_at = float("-inf")
        self.full, self.empty = self._bar_chars()
        self.lock = threading.RLock()   # el hilo de refresco y el programa dibujan la misma línea
        self._stop = threading.Event()
        self._refresher: threading.Thread | None = None

    def _bar_chars(self) -> tuple[str, str]:
        try:
            "█░".encode(getattr(self.stream, "encoding", None) or "ascii")
            return "█", "░"
        except (LookupError, UnicodeEncodeError):
            return "#", "-"

    # -- avance -----------------------------------------------------------------
    def stage(self, key: str, label: str | None = None) -> None:
        """Comienza una etapa; la anterior se da por terminada."""
        with self.lock:
            self._close()
            self.key, self.label = key, label or key
            self.fraction, self.count = 0.0, None
            self.stage_started = self.reported_at = self.clock()
            self.draw(force=True)

    def step(self, done: int, total: int, unit: str = "", fraction: float | None = None) -> None:
        """Avance de la etapa en curso: `done` de `total` unidades (tareas, filas, hojas...).

        fraction: avance de la etapa (0 a 1) cuando las unidades no llevan el mismo tiempo; por defecto, done / total.
        """
        if fraction is None:
            fraction = done / total if total > 0 else 1.0
        with self.lock:
            self.fraction = min(1.0, max(0.0, fraction))
            self.count = (done, total, unit)
            self.reported_at = self.clock()
        self.draw()

    def _close(self) -> None:
        if self.key is not None and self.key not in self.closed:
            seconds = self.clock() - self.stage_started
            self.closed[self.key] = seconds
            weight = self.weights.get(self.key, 0.0)
            if weight:
                self.done += weight
                self.done_seconds += seconds
        self.key = None

    def total(self) -> float:
        """% del total (0 a 1)."""
        current = self.weights.get(self.key, 0.0) * self.fraction if self.key is not None else 0.0
        return min(1.0, self.done + current)

    def elapsed(self) -> float:
        return self.clock() - self.started

    def eta(self) -> float | None:
        """Segundos estimados hasta terminar (None mientras no hay datos suficientes)."""
        current = self.weights.get(self.key, 0.0) if self.key is not None else 0.0
        f = self.fraction
        spent = self.reported_at - self.stage_started   # tiempo de la etapa hasta su último avance informado
        seconds, weight = self.done_seconds, self.done
        if f > 0 and current:
            seconds, weight = seconds + spent, weight + current * f
        if weight < MIN_ETA_FRACTION or self.elapsed() < MIN_ETA_SECONDS:
            return None
        rate = seconds / weight                         # segundos por unidad de peso, medidos
        remaining = rate * current * (1 - f)            # lo que falta de la etapa según su peso
        if f > 0:                                       # y a su propio ritmo, cuando ya avanzó lo suficiente
            trust = min(1.0, f / STAGE_TRUST)
            remaining = trust * spent * (1 - f) / f + (1 - trust) * remaining
        remaining = max(0.0, remaining - (self.clock() - self.reported_at))  # cuenta hacia atrás hasta otro avance
        pending = sum(w for k, w in self.weights.items() if k != self.key and k not in self.closed)
        return remaining + rate * pending

    # -- dibujo -----------------------------------------------------------------
    def line(self) -> str:
        p = self.total()
        filled = int(round(p * BAR_WIDTH))
        bar = self.full * filled + self.empty * (BAR_WIDTH - filled)
        eta = self.eta()
        # Los tiempos van antes de la etapa: en una consola angosta se recorta el texto de la etapa
        text = (f"[{bar}] {int(p * 100):>3} %  Transcurrido {clock_text(self.elapsed())}  ·  "
                f"ETA {'calculando...' if eta is None else clock_text(eta)}  ·  {self.label}")
        if self.count is not None:
            done, total, unit = self.count
            text += (f": {fmt_number(done, 0)} de {fmt_number(total, 0)}{' ' + unit if unit else ''} "
                     f"({int(self.fraction * 100)} %)")
        else:
            text += f" ({clock_text(self.clock() - self.stage_started)})"
        return text

    def draw(self, force: bool = False) -> None:
        if not self.enabled or self.key is None:
            return
        now = self.clock()
        if not force and now - self.drawn_at < REDRAW:
            return
        with self.lock:
            if self.key is None:
                return
            columns = shutil.get_terminal_size((100, 20)).columns - 1
            text = self.line()[:max(10, columns)]
            self.stream.write("\r" + text + " " * max(0, self.width - len(text)))
            self.stream.flush()
            self.width, self.drawn_at = len(text), now

    def clear(self) -> None:
        """Borra la barra (antes de escribir un mensaje o al terminar)."""
        with self.lock:
            if self.enabled and self.width:
                self.stream.write("\r" + " " * self.width + "\r")
                self.stream.flush()
                self.width = 0

    def start(self) -> None:
        """Vuelve a dibujar la barra cada REFRESH segundos, en un hilo aparte, hasta finish()."""
        if self.enabled and self._refresher is None:
            self._refresher = threading.Thread(target=self._refresh, name="barra de progreso", daemon=True)
            self._refresher.start()

    def _refresh(self) -> None:
        while not self._stop.wait(REFRESH):
            self.draw(force=True)

    def finish(self) -> None:
        self._stop.set()
        if self._refresher is not None:
            self._refresher.join(timeout=4 * REFRESH)
        with self.lock:
            self._close()
            self.clear()
        if self.closed:
            log.debug("Duración de las etapas: %s (total %s s)",
                      ", ".join(f"{k} {fmt_number(s, 1)} s" for k, s in self.closed.items()),
                      fmt_number(self.elapsed(), 1))


# ----------------------------------------------------------------------------
# Progreso del comando en curso (las partes del programa no necesitan recibirlo)
# ----------------------------------------------------------------------------
_active: Progress | None = None
_shown = False      # se mostró una barra desde la última consulta de bar_was_shown()


@contextmanager
def tracking(stages: dict[str, float], **kwargs):
    """Activa la barra durante el bloque y la borra al terminar (también si hay un error)."""
    global _active, _shown
    previous, _active = _active, Progress(stages, **kwargs)
    _active.start()
    try:
        yield _active
    finally:
        _active.finish()
        _shown = _shown or _active.enabled
        _active = previous


def bar_was_shown() -> bool:
    """True si se mostró una barra desde la consulta anterior (el comando informa entonces su tiempo total)."""
    global _shown
    shown, _shown = _shown, False
    return shown


def stage(key: str, label: str | None = None) -> None:
    if _active is not None:
        _active.stage(key, label)


def step(done: int, total: int, unit: str = "", fraction: float | None = None) -> None:
    if _active is not None:
        _active.step(done, total, unit, fraction)


def say(text: str) -> None:
    """Escribe un mensaje en la salida estándar sin mezclarlo con la barra."""
    if _active is None:
        print(text, flush=True)
        return
    with _active.lock:
        _active.clear()
        print(text, flush=True)
        _active.draw(force=True)
