"""Versión en pantalla y barra de progreso de la consola."""
import io

import pytest

from evm import __version__, progress
from evm.cli import main
from evm.progress import Progress


class Clock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now


class Tty(io.StringIO):
    """Consola interactiva simulada."""

    encoding = "utf-8"

    def isatty(self):
        return True


def test_every_command_shows_the_version(tmp_path, capsys):
    assert main(["analizar", "--libro", str(tmp_path / "no_existe.xlsx")]) == 2
    out = capsys.readouterr().out
    assert out.splitlines()[0] == f"Exportador EVM {__version__}"


def test_no_bar_when_the_output_is_not_a_console():
    stream = io.StringIO()
    p = Progress({"a": 1.0}, stream=stream)
    p.stage("a", "Lectura")
    p.step(5, 10, "tareas")
    p.finish()
    assert stream.getvalue() == ""


def test_bar_shows_stage_progress_total_and_eta():
    clock, stream = Clock(), Tty()
    p = Progress({"leer": 3.0, "guardar": 1.0}, stream=stream, clock=clock)
    p.stage("leer", "Lectura de MS Project")
    clock.now = 3.0
    p.step(500, 1000, "tareas")
    line = p.line()
    assert "Lectura de MS Project: 500 de 1.000 tareas (50 %)" in line
    assert " 37 %" in line                                   # 3/4 × 50 %
    assert "ETA 0:05" in line                                # 3 s para el 37,5 %: faltan 5 s
    assert p.total() == pytest.approx(0.375)
    p.stage("guardar", "Guardado del libro")
    assert p.total() == pytest.approx(0.75)
    clock.now = 4.0
    assert "Guardado del libro (0:01)" in p.line()           # etapa sin conteo: tiempo de la etapa
    p.finish()
    assert stream.getvalue().endswith("\r")                  # la barra se borra al terminar
    assert p.width == 0


def test_eta_waits_for_enough_progress_and_bar_is_throttled():
    clock, stream = Clock(), Tty()
    p = Progress({"a": 1.0}, stream=stream, clock=clock)
    p.stage("a", "A")
    clock.now = 0.5
    p.step(1, 100)
    assert "ETA calculando..." in p.line()
    writes = stream.getvalue().count("\r")
    p.step(2, 100)                                           # menos de REDRAW segundos: no se vuelve a dibujar
    assert stream.getvalue().count("\r") == writes


def test_ascii_bar_when_the_console_cannot_show_blocks():
    class Ascii(Tty):
        encoding = "ascii"

    p = Progress({"a": 1.0}, stream=Ascii())
    p.stage("a", "A")
    p.step(1, 2)
    assert p.line().startswith("[##########----------]")


def test_module_functions_without_active_bar_do_nothing(capsys):
    progress.stage("x")
    progress.step(1, 2)
    progress.say("mensaje")
    assert capsys.readouterr().out == "mensaje\n"
    with progress.tracking({"x": 1.0}, stream=Tty()) as p:
        progress.stage("x", "X")
        assert p.key == "x"
    assert progress._active is None
