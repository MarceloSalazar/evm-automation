"""Pruebas opcionales con MS Project real: lectura del .mpp (COM) frente al XML del mismo proyecto.

Solo se ejecutan en un PC con MS Project y si se indica el proyecto con variables de entorno:

    $env:EVM_MPP = "C:\\Proyectos\\Programa.mpp"
    $env:EVM_MPP_XML = "C:\\Proyectos\\Programa.xml"     # el mismo proyecto: Archivo > Guardar como > XML
    python -m pytest -m msproject -v

Opcionales:
- EVM_CONFIG: archivo de configuración (por defecto config.toml si existe; si no, config.ejemplo.toml).
- EVM_LINEA_BASE: línea base a comparar (0..10), como --linea-base.
- EVM_TOLERANCIA_PV: diferencia admitida en el % planificado de cada mes (por defecto 0.005, es decir,
  0,5 %). El motor XML reparte la duración de las tareas con el calendario del proyecto; el motor COM, con
  lunes a viernes y los feriados de la configuración.

Antes de guardar el XML se debe definir la fecha de estado (Proyecto > Fecha de estado) y guardar el
.mpp: ambos archivos deben ser el mismo proyecto con la misma fecha de estado. Los mensajes de las
pruebas solo muestran códigos EDT y porcentajes, no nombres de tareas. Los dos libros reciben el mismo BAC
en cada fila, como se ingresa a mano.
"""
import contextlib
import faulthandler
import io
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

from evm.cli import main
from evm.update import read_book

from builders import set_bac

ROOT = Path(__file__).resolve().parents[1]
MPP = os.environ.get("EVM_MPP", "")
XML = os.environ.get("EVM_MPP_XML", "")
LIMIT = 15

MISSING = [name for name, value in (("EVM_MPP", MPP), ("EVM_MPP_XML", XML)) if not value]

pytestmark = [
    pytest.mark.msproject,
    pytest.mark.skipif(sys.platform != "win32", reason="MS Project solo funciona en Windows"),
    pytest.mark.skipif(bool(MISSING), reason=(
        f"Falta definir {' y '.join(MISSING)} en esta terminal. En PowerShell: "
        '$env:EVM_MPP = "C:\\ruta\\Programa.mpp"; $env:EVM_MPP_XML = "C:\\ruta\\Programa.xml" '
        "(en cmd: set EVM_MPP=C:\\ruta\\Programa.mpp). El XML lo guarda MS Project: Archivo > Guardar como > "
        "Formato XML.")),
]


def ms_project_running() -> bool:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINPROJ.EXE", "/NH"], capture_output=True, text=True)
    return "WINPROJ.EXE" in out.stdout.upper()


def config_path() -> str:
    """Configuración de la prueba: EVM_CONFIG, config.toml si existe o config.ejemplo.toml."""
    if os.environ.get("EVM_CONFIG"):
        return os.environ["EVM_CONFIG"]
    local = ROOT / "config.toml"
    return str(local if local.exists() else ROOT / "config.ejemplo.toml")


def run(*args) -> str:
    """Ejecuta un comando del programa y devuelve lo que muestra en pantalla."""
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = main([str(a) for a in args])
    text = out.getvalue() + err.getvalue()
    assert rc == 0, f"'{args[0]}' terminó con código {rc}:\n{text}"
    return text


def wait_until_closed(seconds: float = 60) -> bool:
    """Espera a que termine el proceso de MS Project (True si terminó)."""
    deadline = time.monotonic() + seconds
    while ms_project_running():
        if time.monotonic() >= deadline:
            return False
        time.sleep(1)
    return True


@pytest.fixture(scope="module", autouse=True)
def quiet_faulthandler():
    """En Windows, faulthandler informa como "Windows fatal exception: code 0x800706ba" los errores de RPC
    que COM maneja internamente mientras MS Project se cierra; no son fatales. Se desactiva solo aquí."""
    enabled = faulthandler.is_enabled()
    faulthandler.disable()
    yield
    if enabled:
        faulthandler.enable()


def report(diffs: list[str], what: str) -> str:
    more = f"\n... y {len(diffs) - LIMIT} más" if len(diffs) > LIMIT else ""
    return f"{len(diffs)} diferencias en {what} (COM frente a XML):\n" + "\n".join(diffs[:LIMIT]) + more


@pytest.fixture(scope="module")
def books(tmp_path_factory):
    """Genera y actualiza un libro con cada motor (COM desde el .mpp, XML desde el XML)."""
    pytest.importorskip("win32com.client", reason="Se necesita pywin32")
    for path in (MPP, XML):
        assert Path(path).is_file(), f"No existe el archivo {path}"
    d = tmp_path_factory.mktemp("msproject")
    extra = ["--linea-base", os.environ["EVM_LINEA_BASE"]] if os.environ.get("EVM_LINEA_BASE") else []
    result = {"was_running": ms_project_running()}
    closed = []          # el MS Project iniciado por el programa terminó después de cada paso
    for engine, origin in (("com", MPP), ("xml", XML)):
        book = d / f"EVM_{engine}.xlsx"
        for command in (["generar", "--salida", book], ["actualizar", "--libro", book, "--sin-respaldo"]):
            run(command[0], "--origen", origin, "--motor", engine, "--config", config_path(), *command[1:], *extra)
            if command[0] == "generar":
                set_bac(book)                   # el mismo BAC en cada fila de los dos libros
            if engine == "com" and not result["was_running"]:
                # MS Project tarda unos segundos en cerrarse después de Quit: se espera antes del paso siguiente
                closed.append(wait_until_closed())
        result[engine] = read_book(book)
    result["running_after"] = ms_project_running() if result["was_running"] else not all(closed)
    return result


def by_row(book):
    return {s.row: s for s in book.structure}


def test_same_structure_methods_and_horizon(books):
    com, xml = books["com"], books["xml"]
    assert (com.horizon.start, com.horizon.months) == (xml.horizon.start, xml.horizon.months)
    assert com.multi == xml.multi
    a = [(s.row, s.key, s.code, s.level, s.kind) for s in com.structure]
    b = [(s.row, s.key, s.code, s.level, s.kind) for s in xml.structure]
    diffs = [f"fila {x[0]}: COM {x[1:]} / XML {y[1:]}" for x, y in zip(a, b) if x != y]
    if len(a) != len(b):
        diffs.insert(0, f"cantidad de filas: COM {len(a)} / XML {len(b)}")
    assert not diffs, report(diffs, "la estructura EDT (fila, UID, código, nivel, tipo)")
    names = [f"fila {s.row} ({s.code})" for s, t in zip(com.structure, xml.structure)
             if s.name.strip() != t.name.strip()]
    assert not names, report(names, "los nombres de las tareas")
    methods = [f"{by_row(com)[r].code}: COM '{m}' / XML '{xml.method.get(r)}'" for r, m in com.method.items()
               if m != xml.method.get(r)]
    assert not methods, report(methods, "los métodos de medición")


def test_same_planned_percent(books):
    com, xml = books["com"], books["xml"]
    tol = float(os.environ.get("EVM_TOLERANCIA_PV", "0.005"))
    rows = by_row(com)
    diffs = []
    for role, company in com.roles():
        who = f" ({company})" if company else ""
        for r, months in com.pct_pv[role].items():
            other = xml.pct_pv[role][r]
            for i, (a, b) in enumerate(zip(months, other)):
                if abs(a - b) > tol:
                    diffs.append(f"{rows[r].code}{who} mes {i + 1} ({com.horizon.month_start(i):%m/%Y}): "
                                 f"COM {a:.2%} / XML {b:.2%}")
            if abs(sum(months) - sum(other)) > 1e-6:
                diffs.append(f"{rows[r].code}{who} total: COM {sum(months):.2%} / XML {sum(other):.2%}")
    assert not diffs, report(diffs, "el % planificado de cada mes")


def test_same_earned_percent_at_cutoff(books):
    com, xml = books["com"], books["xml"]
    assert com.status_date == xml.status_date, (
        f"Fecha de corte: COM {com.status_date} / XML {xml.status_date}. El XML se debe guardar después de "
        "definir la fecha de estado en MS Project.")
    idx = com.horizon.index(com.status_date)
    rows = by_row(com)
    diffs = []
    for role, company in com.roles():
        who = f" ({company})" if company else ""
        for r, months in com.pct_ev[role].items():
            x = sum(v or 0.0 for v in months[: idx + 1])
            y = sum(v or 0.0 for v in xml.pct_ev[role][r][: idx + 1])
            if abs(x - y) > 1e-6:
                diffs.append(f"{rows[r].code}{who}: COM {x:.2%} / XML {y:.2%}")
    assert not diffs, report(diffs, "el % ganado acumulado a la fecha de corte")


def test_ms_project_is_left_as_found(books):
    """Si MS Project no estaba abierto, el programa lo cierra; si estaba abierto, no lo cierra."""
    if books["was_running"]:
        assert books["running_after"], "El programa cerró la instancia de MS Project que ya estaba abierta."
    else:
        assert not books["running_after"], ("MS Project quedó abierto en segundo plano (WINPROJ.EXE) después de "
                                             "las pruebas: se debe cerrar desde el Administrador de tareas.")
