"""Versión en pantalla y barra de progreso de la consola."""
import io
import logging
import threading

import pytest

from evm import __version__, progress
from evm.cli import _book_stages, main
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


def test_bar_shows_stage_progress_total_elapsed_and_eta():
    clock, stream = Clock(), Tty()
    p = Progress({"leer": 3.0, "guardar": 1.0}, stream=stream, clock=clock)
    p.stage("leer", "Lectura de MS Project")
    clock.now = 3.0
    p.step(500, 1000, "tareas")
    line = p.line()
    assert "Lectura de MS Project: 500 de 1.000 tareas (50 %)" in line
    assert " 37 %" in line                                   # 3/4 × 50 %
    # La mitad de la lectura en 3 s: faltan 3 s de lectura y 2 s de guardado (1/3 de su peso, al mismo ritmo)
    assert "Transcurrido 0:03  ·  ETA 0:05  ·  Lectura" in line   # los tiempos, antes del texto de la etapa
    assert p.total() == pytest.approx(0.375)
    p.stage("guardar", "Guardado del libro")
    assert p.total() == pytest.approx(0.75)
    clock.now = 4.0
    assert "Guardado del libro (0:01)" in p.line()           # etapa sin conteo: tiempo de la etapa
    p.finish()
    assert stream.getvalue().endswith("\r")                  # la barra se borra al terminar
    assert p.width == 0


def track(p, clock, timeline):
    """ETA de cada segundo de una ejecución: [(etapa, segundos, unidades o 0 si no informa su avance)]."""
    etas = []
    for key, seconds, units in timeline:
        p.stage(key)
        start = clock.now
        for s in range(1, seconds + 1):
            clock.now = start + s
            if units:
                p.step(units * s // seconds, units)
            etas.append(p.eta())
    return etas


def test_eta_does_not_rise_while_a_stage_is_slower_than_its_weight():
    # MS Project ya abierto (2 s, menos que su peso) y una lectura larga: el ETA del total no sube y luego baja,
    # porque lo que falta de la lectura se estima a su propio ritmo
    clock = Clock()
    p = Progress({"abrir": 8.0, "leer": 10.0, "escribir": 6.0, "guardar": 4.0}, enabled=False, clock=clock)
    etas = track(p, clock, [("abrir", 2, 0), ("leer", 60, 6000), ("escribir", 9, 12), ("guardar", 6, 0)])
    reading = etas[2 + 6:2 + 60]                             # desde el 10 % de la lectura
    assert all(b <= a for a, b in zip(reading, reading[1:]))
    real = 77 - (2 + 30)                                     # a mitad de la lectura faltan 45 s
    assert etas[2 + 30 - 1] == pytest.approx(real, rel=0.25)


def test_stage_without_count_counts_down():
    # Lectura del libro (sin conteo) después de una etapa medida: el ETA baja segundo a segundo hasta el guardado
    clock = Clock()
    p = Progress({"leer": 4.0, "libro": 4.0, "guardar": 2.0}, enabled=False, clock=clock)
    etas = track(p, clock, [("leer", 4, 400), ("libro", 6, 0)])
    assert etas[3] == pytest.approx(6.0)                     # 1 s por unidad de peso: 4 s de libro y 2 s de guardado
    assert etas[4:8] == pytest.approx([5.0, 4.0, 3.0, 2.0])
    assert etas[8:] == pytest.approx([2.0, 2.0])             # el libro tarda más de lo esperado: queda el guardado


def test_eta_of_a_stage_combines_its_weight_and_its_own_pace_at_the_start():
    clock = Clock()
    p = Progress({"a": 1.0, "b": 1.0}, enabled=False, clock=clock)
    p.stage("a")
    clock.now = 10.0
    p.stage("b")                                             # 10 s por unidad de peso
    clock.now = 10.5
    p.step(1, 100)                     # 1 % en 0,5 s: a su ritmo faltarían 49,5 s; según su peso, unos 10 s
    assert 10.0 < p.eta() < 20.0                             # al comienzo de la etapa pesa más su peso
    clock.now = 13.0
    p.step(50, 100)                                          # desde el 10 %, solo su propio ritmo
    assert p.eta() == pytest.approx(3.0)


def test_book_stages_follow_the_months_of_the_horizon():
    assert _book_stages(120, libro=10.0) == {"libro": pytest.approx(10.0)}
    assert _book_stages(36, libro=10.0, guardar=4.0) == {"libro": pytest.approx(5.1), "guardar": pytest.approx(2.04)}


def test_refresher_redraws_during_a_stage_without_count(monkeypatch):
    monkeypatch.setattr(progress, "REFRESH", 0.01)
    stream = Tty()
    p = Progress({"a": 1.0}, stream=stream)
    p.stage("a", "Apertura en MS Project")
    drawn = stream.getvalue().count("\r")
    p.start()
    redrawn = threading.Event()
    for _ in range(200):                                     # hasta 2 s
        if stream.getvalue().count("\r") > drawn:
            redrawn.set()
            break
        threading.Event().wait(0.01)
    p.finish()
    assert redrawn.is_set()
    assert not p._refresher.is_alive()
    assert stream.getvalue().endswith("\r")


def test_stage_durations_are_logged(caplog):
    clock = Clock()
    p = Progress({"leer": 1.0, "guardar": 1.0}, enabled=False, clock=clock)
    p.stage("leer")
    clock.now = 2.5
    p.stage("guardar")
    clock.now = 3.0
    with caplog.at_level(logging.DEBUG, logger="evm.progress"):
        p.finish()
    assert "Duración de las etapas: leer 2,5 s, guardar 0,5 s (total 3,0 s)" in caplog.text


def test_total_time_after_a_command_shown_with_bar(monkeypatch, capsys, example_files):
    monkeypatch.setattr(progress, "_interactive", lambda stream: True)
    assert main(["validar", "--origen", str(example_files["base"]), "--config", str(example_files["config"])]) == 0
    captured = capsys.readouterr()
    assert captured.out.splitlines()[-1].startswith("Tiempo total: 0:0")
    assert "Transcurrido" in captured.err and captured.err.endswith("\r")
    monkeypatch.setattr(progress, "_interactive", lambda stream: False)
    assert main(["validar", "--origen", str(example_files["base"]), "--config", str(example_files["config"])]) == 0
    assert "Tiempo total" not in capsys.readouterr().out    # salida redirigida: sin barra ni tiempo total


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
