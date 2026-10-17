"""Pruebas opcionales con MS Project real: lectura del .mpp (COM) frente al XML del mismo proyecto.

Solo se ejecutan en un PC con MS Project y si se indica el proyecto con variables de entorno:

    $env:EVM_MPP = "C:\\Proyectos\\Programa.mpp"
    $env:EVM_MPP_XML = "C:\\Proyectos\\Programa.xml"     # el mismo proyecto: Archivo > Guardar como > XML
    python -m pytest -m msproject -v

Opcionales:
- EVM_CONFIG: archivo de configuración (por defecto config.toml si existe; si no, config.ejemplo.toml).
- EVM_LINEA_BASE: línea base a comparar (0..10), como --linea-base.
- EVM_TOLERANCIA_PV: diferencia admitida en el PV de cada mes, como fracción del BAC de la fila
  (por defecto 0.005). El motor XML reparte por el calendario del proyecto los tramos que cruzan de un
  mes a otro; MS Project usa además los calendarios de recursos y tareas.

Antes de guardar el XML se debe definir la fecha de estado (Proyecto > Fecha de estado) y guardar el
.mpp: ambos archivos deben ser el mismo proyecto con la misma fecha de estado. Los mensajes de las
pruebas solo muestran códigos EDT e importes, no nombres de tareas.
"""
import contextlib
import io
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

from evm.cli import main
from evm.update import read_book

ROOT = Path(__file__).resolve().parents[1]
MPP = os.environ.get("EVM_MPP", "")
XML = os.environ.get("EVM_MPP_XML", "")
LIMIT = 15

pytestmark = [
    pytest.mark.msproject,
    pytest.mark.skipif(sys.platform != "win32", reason="MS Project solo funciona en Windows"),
    pytest.mark.skipif(not (MPP and XML), reason="Se deben indicar EVM_MPP y EVM_MPP_XML (ver el docstring)"),
]


def ms_project_running() -> bool:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINPROJ.EXE", "/NH"], capture_output=True, text=True)
    return "WINPROJ.EXE" in out.stdout.upper()


def config_path() -> str:
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
    for engine, origin in (("com", MPP), ("xml", XML)):
        book = d / f"EVM_{engine}.xlsx"
        run("generar", "--origen", origin, "--motor", engine, "--config", config_path(), "--salida", book, *extra)
        run("actualizar", "--origen", origin, "--motor", engine, "--config", config_path(), "--libro", book,
            "--sin-respaldo", *extra)
        result[engine] = read_book(book)
    # MS Project tarda unos segundos en cerrarse después de Quit
    deadline = time.monotonic() + 30
    running = ms_project_running()
    while running and not result["was_running"] and time.monotonic() < deadline:
        time.sleep(1)
        running = ms_project_running()
    result["running_after"] = running
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


def test_same_bac(books):
    com, xml = books["com"], books["xml"]
    rows = by_row(com)
    diffs = [f"{rows[r].code}: COM {v:,.2f} / XML {xml.bac[r]:,.2f}" for r, v in com.bac.items()
             if abs(v - xml.bac[r]) > 0.01]
    if com.multi:
        diffs += [f"{rows[r].code} (empresa A): COM {v:,.2f} / XML {xml.bac_a[r]:,.2f}" for r, v in com.bac_a.items()
                  if abs(v - xml.bac_a[r]) > 0.01]
    assert not diffs, report(diffs, "el BAC de las filas")


def test_same_monthly_pv(books):
    com, xml = books["com"], books["xml"]
    tol = float(os.environ.get("EVM_TOLERANCIA_PV", "0.005"))
    rows = by_row(com)
    diffs = []
    for r, months in com.pv.items():
        limit = max(1.0, tol * abs(com.bac[r]))
        for i, (a, b) in enumerate(zip(months, xml.pv[r])):
            if abs(a - b) > limit:
                diffs.append(f"{rows[r].code} mes {i + 1} ({com.horizon.month_start(i):%m/%Y}): COM {a:,.2f} / "
                             f"XML {b:,.2f}")
        if abs(sum(months) - sum(xml.pv[r])) > 0.05:
            diffs.append(f"{rows[r].code} total: COM {sum(months):,.2f} / XML {sum(xml.pv[r]):,.2f}")
    assert not diffs, report(diffs, "el PV mensual")


def test_same_ev_and_ac_at_cutoff(books):
    com, xml = books["com"], books["xml"]
    assert com.status_date == xml.status_date, (
        f"Fecha de corte: COM {com.status_date} / XML {xml.status_date}. El XML se debe guardar después de "
        "definir la fecha de estado en MS Project.")
    idx = com.horizon.index(com.status_date)
    rows = by_row(com)
    series = [("EV", com.ev, xml.ev), ("AC", com.ac, xml.ac)]
    if com.multi:
        series += [("EV empresa A", com.ev_a, xml.ev_a), ("AC empresa A", com.ac_a, xml.ac_a)]
    diffs = []
    for label, a, b in series:
        for r in a:
            x = sum(v or 0.0 for v in a[r][: idx + 1])
            y = sum(v or 0.0 for v in b[r][: idx + 1])
            if abs(x - y) > max(0.05, 1e-6 * abs(com.bac.get(r, 0.0))):
                diffs.append(f"{rows[r].code} {label}: COM {x:,.2f} / XML {y:,.2f}")
    assert not diffs, report(diffs, "el EV y el AC acumulados a la fecha de corte")


def test_ms_project_is_left_as_found(books):
    """Si MS Project no estaba abierto, el programa lo cierra; si estaba abierto, no lo cierra."""
    if books["was_running"]:
        assert books["running_after"], "El programa cerró la instancia de MS Project que ya estaba abierta."
    else:
        assert not books["running_after"], ("MS Project quedó abierto en segundo plano (WINPROJ.EXE) después de "
                                             "las pruebas: se debe cerrar desde el Administrador de tareas.")
