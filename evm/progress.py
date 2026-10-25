"""Barra de progreso de la consola: tarea en curso, % de la tarea y del total, y tiempo restante estimado (ETA).

    [██████░░░░░░░░░░░░░░]  31 %  Lectura de MS Project: 1.234 de 3.000 tareas (41 %)  ·  ETA 0:42

Cada comando declara sus etapas con un peso relativo (el tiempo que suele llevar cada una) y las partes del programa
informan su avance con stage() y step(), sin saber si hay una barra activa. El % total combina las etapas
terminadas y el avance de la etapa en curso; el ETA proyecta el tiempo transcurrido con ese %.

La barra se escribe en la salida de errores y solo si es una consola interactiva: con la salida redirigida
(archivos de registro, pruebas, instalador) no se muestra nada.
"""
from __future__ import annotations

import shutil
import sys
import time
from collections.abc import Callable
from contextlib import contextmanager

from evm.fields import fmt_number

BAR_WIDTH = 20
REDRAW = 0.1            # segundos mínimos entre dos dibujos de la barra
MIN_ETA_FRACTION = 0.03 # el ETA se muestra a partir de este % del total
MIN_ETA_SECONDS = 1.0


def _interactive(stream) -> bool:
    try:
        return bool(stream.isatty())
    except (AttributeError, ValueError):
        return False


def _clock_text(seconds: float) -> str:
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
        self.closed: set[str] = set()
        self.key: str | None = None
        self.label = ""
        self.fraction = 0.0             # avance de la etapa en curso (0 a 1)
        self.count: tuple[int, int, str] | None = None
        self.stage_started = self.started
        self.width = 0                  # caracteres de la última línea dibujada
        self.drawn_at = float("-inf")
        self.full, self.empty = self._bar_chars()

    def _bar_chars(self) -> tuple[str, str]:
        try:
            "█░".encode(getattr(self.stream, "encoding", None) or "ascii")
            return "█", "░"
        except (LookupError, UnicodeEncodeError):
            return "#", "-"

    # -- avance -----------------------------------------------------------------
    def stage(self, key: str, label: str | None = None) -> None:
        """Comienza una etapa; la anterior se da por terminada."""
        self._close()
        self.key, self.label = key, label or key
        self.fraction, self.count = 0.0, None
        self.stage_started = self.clock()
        self.draw(force=True)

    def step(self, done: int, total: int, unit: str = "") -> None:
        """Avance de la etapa en curso: `done` de `total` unidades (tareas, filas, hojas...)."""
        self.fraction = min(1.0, done / total) if total > 0 else 1.0
        self.count = (done, total, unit)
        self.draw()

    def _close(self) -> None:
        if self.key is not None and self.key not in self.closed:
            self.closed.add(self.key)
            self.done += self.weights.get(self.key, 0.0)
        self.key = None

    def total(self) -> float:
        """% del total (0 a 1)."""
        current = self.weights.get(self.key, 0.0) * self.fraction if self.key is not None else 0.0
        return min(1.0, self.done + current)

    def eta(self) -> float | None:
        """Segundos estimados hasta terminar (None mientras no hay datos suficientes)."""
        p = self.total()
        elapsed = self.clock() - self.started
        if p < MIN_ETA_FRACTION or elapsed < MIN_ETA_SECONDS:
            return None
        return elapsed * (1 - p) / p

    # -- dibujo -----------------------------------------------------------------
    def line(self) -> str:
        p = self.total()
        filled = int(round(p * BAR_WIDTH))
        bar = self.full * filled + self.empty * (BAR_WIDTH - filled)
        text = f"[{bar}] {int(p * 100):>3} %  {self.label}"
        if self.count is not None:
            done, total, unit = self.count
            text += (f": {fmt_number(done, 0)} de {fmt_number(total, 0)}{' ' + unit if unit else ''} "
                     f"({int(self.fraction * 100)} %)")
        else:
            text += f" ({_clock_text(self.clock() - self.stage_started)})"
        eta = self.eta()
        text += f"  ·  ETA {_clock_text(eta)}" if eta is not None else "  ·  ETA calculando..."
        return text

    def draw(self, force: bool = False) -> None:
        if not self.enabled or self.key is None:
            return
        now = self.clock()
        if not force and now - self.drawn_at < REDRAW:
            return
        columns = shutil.get_terminal_size((100, 20)).columns - 1
        text = self.line()[:max(10, columns)]
        self.stream.write("\r" + text + " " * max(0, self.width - len(text)))
        self.stream.flush()
        self.width, self.drawn_at = len(text), now

    def clear(self) -> None:
        """Borra la barra (antes de escribir un mensaje o al terminar)."""
        if self.enabled and self.width:
            self.stream.write("\r" + " " * self.width + "\r")
            self.stream.flush()
            self.width = 0

    def finish(self) -> None:
        self._close()
        self.clear()


# ----------------------------------------------------------------------------
# Progreso del comando en curso (las partes del programa no necesitan recibirlo)
# ----------------------------------------------------------------------------
_active: Progress | None = None


@contextmanager
def tracking(stages: dict[str, float], **kwargs):
    """Activa la barra durante el bloque y la borra al terminar (también si hay un error)."""
    global _active
    previous, _active = _active, Progress(stages, **kwargs)
    try:
        yield _active
    finally:
        _active.finish()
        _active = previous


def stage(key: str, label: str | None = None) -> None:
    if _active is not None:
        _active.stage(key, label)


def step(done: int, total: int, unit: str = "") -> None:
    if _active is not None:
        _active.step(done, total, unit)


def say(text: str) -> None:
    """Escribe un mensaje en la salida estándar sin mezclarlo con la barra."""
    if _active is not None:
        _active.clear()
    print(text, flush=True)
    if _active is not None:
        _active.draw(force=True)
