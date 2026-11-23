"""Pruebas opcionales con MS Project real de corregir-linea-base: mover la línea base de una tarea a una línea base
libre en una copia del .mpp cerrada (con respaldo y guardado) y volver a moverla; borrar en una copia abierta (sin
guardar, con una sola acción de deshacer); establecer una línea base libre con los valores actuales; el texto de NA
que acepta MS Project («NA» en inglés, «ND» en español); y la escritura de la línea base de una tarea resumen
(trabajo y costo previstos).

Se ejecutan en un PC con MS Project:

    $env:EVM_MPP = "C:\\Proyectos\\Programa.mpp"
    $env:EVM_CONFIG = "C:\\Proyectos\\config.toml"
    $env:EVM_LB_TAREA = "2.7.1.3"       # opcional: por defecto, la primera tarea de detalle con la línea base del PV
    $env:EVM_LB_LIBRE = "10"            # opcional: por defecto, la última línea base sin ninguna tarea
    python -m pytest tests/test_msproject_linea_base.py -m msproject -v -s

El archivo de EVM_MPP no se modifica: las pruebas trabajan con copias en la carpeta temporal de pytest. Los mensajes
solo muestran UID, códigos EDT y números de línea base, nunca nombres de tareas.
"""
import contextlib
import hashlib
import io
import os
import shutil
import sys
from pathlib import Path

import pytest

from evm import baseline_fix as bf
from evm.cli import main
from evm.config import load_config
from evm.reader_com import ComSession
from evm.wbs import build_tree, select_scope

from test_msproject_real import (MPP, Hidden, config_path, ms_project_running, quiet_faulthandler,  # noqa: F401
                                 wait_until_closed)

TASK = os.environ.get("EVM_LB_TAREA", "")
FREE = os.environ.get("EVM_LB_LIBRE", "")

pytestmark = [
    pytest.mark.msproject,
    pytest.mark.skipif(sys.platform != "win32", reason="MS Project solo funciona en Windows"),
    pytest.mark.skipif(not MPP, reason='Falta definir EVM_MPP (en PowerShell: $env:EVM_MPP = "C:\\ruta\\Programa.mpp").'),
]


def _run(*args) -> tuple[int, str]:
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        rc = main([str(a) for a in args])
    return rc, out.getvalue()


def _hash(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _codes_of(text: str) -> str:
    """Solo los códigos de los mensajes (sin nombres de tareas) para los mensajes de error de la prueba."""
    import re
    return ", ".join(sorted(set(re.findall(r"\b[A-Z]{3}-\d{3}\b", text)))) or "ninguno"


def _values(mpp: Path, uid: int, slots: list[int]) -> dict[int, bf.Values]:
    with ComSession(str(mpp)) as s:
        com = bf.ComStore(s)
        out = {k: bf.read_values(com.tasks[uid], k) for k in slots}
        out["notes"] = bf._notes_of(com.tasks[uid])
    return out


@pytest.fixture(scope="module")
def setup(tmp_path_factory):
    pytest.importorskip("win32com.client", reason="Se necesita pywin32")
    assert Path(MPP).is_file(), f"No existe el archivo {MPP}"
    cfg = load_config(config_path())
    d = tmp_path_factory.mktemp("linea_base")
    print(f"\nArchivos de la prueba: {d}")
    data = Hidden(was_running=ms_project_running(), folder=d, source_hash=_hash(MPP))
    data["copy"] = d / "lb.mpp"
    shutil.copyfile(MPP, data["copy"])
    with ComSession(str(data["copy"])) as s:
        project = s.read_project(cfg, links=False, review=True)
        store = bf.ComStore(s)
        pv = project.baseline_number
        counts = bf.baseline_counts(project)
        free = int(FREE) if FREE else next((k for k in range(10, -1, -1) if not counts[k]), None)
        if free is None:
            pytest.skip("Todas las líneas base tienen tareas: se debe indicar una libre con EVM_LB_LIBRE.")
        scope = select_scope(build_tree(project.tasks), cfg.plan)
        if TASK:
            task = bf.find_task(scope, TASK)
        else:
            task = next((t for t in scope.tasks() if not t.children and (t.wbs or "").strip() and t.has_baseline
                         and pv in t.baseline_dates and not t.milestone), None)
        if task is None:
            pytest.skip("No hay ninguna tarea de detalle con código EDT y con la línea base del PV.")
        summary = next((t for t in scope.tasks() if t.children and (t.wbs or "").strip()
                        and pv in t.baseline_dates), None)
        data.update(cfg=cfg, pv=pv, free=free, uid=task.uid, code=task.wbs.strip(),
                    original=bf.read_values(store.tasks[task.uid], pv),
                    notes=bf._notes_of(store.tasks[task.uid]),
                    summary=summary.uid if summary is not None else None,
                    summary_code=(summary.wbs or "").strip() if summary is not None else "")
    if not data["was_running"]:
        wait_until_closed()
    print(f"Tarea {data['code']} (UID {data['uid']}): línea base {data['pv']} -> {data['free']} y vuelta")
    return data


def test_closed_file_moves_saves_and_backs_up(setup):
    d, mpp = setup["folder"], setup["copy"]
    rc, text = _run("corregir-linea-base", "--origen", mpp, "--config", config_path(), "--tarea", setup["code"],
                    "--de", setup["pv"], "--a", setup["free"], "--motivo", "Prueba automática", "--continuar",
                    "--informe", d / "mover.md")
    assert rc == 0, f"corregir-linea-base terminó con código {rc} (mensajes: {_codes_of(text)})"
    assert len(list(d.glob("lb.respaldo-*.mpp"))) == 1
    got = _values(mpp, setup["uid"], [setup["pv"], setup["free"]])
    if not setup["was_running"]:
        wait_until_closed()
    assert got[setup["pv"]].empty, f"La línea base {setup['pv']} no quedó vacía: {got[setup['pv']].dates()}"
    assert bf.same(got[setup["free"]], setup["original"], numbers=False), \
        f"Fechas copiadas: {got[setup['free']].dates()}; se esperaba {setup['original'].dates()}"
    assert abs(got[setup["free"]].duration - setup["original"].duration) < 0.01, "La duración prevista cambió"
    assert got["notes"].replace("\r", "\n").rstrip().endswith(
        f"Línea base {setup['pv']} movida a línea base {setup['free']}: Prueba automática")
    assert _hash(MPP) == setup["source_hash"], "Se modificó el archivo de EVM_MPP"


def test_moving_back_restores_the_original_values(setup):
    mpp = setup["copy"]
    rc, text = _run("corregir-linea-base", "--origen", mpp, "--config", config_path(), "--tarea", setup["code"],
                    "--de", setup["free"], "--a", setup["pv"], "--motivo", "Vuelta", "--continuar", "--sin-respaldo",
                    "--informe", setup["folder"] / "vuelta.md")
    assert rc == 0, f"corregir-linea-base terminó con código {rc} (mensajes: {_codes_of(text)})"
    got = _values(mpp, setup["uid"], [setup["pv"], setup["free"]])
    if not setup["was_running"]:
        wait_until_closed()
    assert got[setup["free"]].empty
    assert bf.same(got[setup["pv"]], setup["original"]), \
        f"Valores: {got[setup['pv']].full(480)}; se esperaba {setup['original'].full(480)}"


def test_open_file_delete_is_not_saved_and_one_undo_restores_it(setup):
    mpp = setup["folder"] / "abierto.mpp"
    shutil.copyfile(MPP, mpp)
    disk = _hash(mpp)
    with ComSession(str(mpp), read_only=False) as outer:
        opened = _hash(mpp)
        rc, text = _run("corregir-linea-base", "--origen", mpp, "--config", config_path(), "--tarea", setup["code"],
                        "--borrar", setup["pv"], "--motivo", "Borrado de prueba", "--continuar",
                        "--informe", setup["folder"] / "abierto.md")
        after = _hash(mpp)
        # Diagnóstico: MS Project puede reescribir el .mpp al abrirlo para escribir o con el autoguardado; lo que se
        # comprueba de verdad es que en el disco sigue la línea base original (al final, ya cerrado sin guardar).
        print(f"\nArchivo en el disco: {'cambió' if opened != disk else 'igual'} al abrirlo en MS Project; "
              f"{'cambió' if after != opened else 'igual'} durante corregir-linea-base")
        assert rc == 0, f"corregir-linea-base terminó con código {rc} (mensajes: {_codes_of(text)})"
        assert "sin guardar" in text, "corregir-linea-base no dejó el archivo abierto sin guardar"
        assert not list(setup["folder"].glob("abierto.respaldo-*.mpp")), "Quedó un respaldo de un archivo sin guardar"
        com = bf.ComStore(outer).tasks[setup["uid"]]
        assert bf.read_values(com, setup["pv"]).empty
        assert "Ctrl+Z" in text, "Los cambios no se agruparon en una sola acción de deshacer (LBA-019)"
        outer.app.EditUndo()
        com = bf.ComStore(outer).tasks[setup["uid"]]
        assert bf.same(bf.read_values(com, setup["pv"]), setup["original"]), "Un Ctrl+Z no restauró la línea base"
        assert bf._notes_of(com) == setup["notes"], "Un Ctrl+Z no restauró las Notas"
    got = _values(mpp, setup["uid"], [setup["pv"]])
    if not setup["was_running"]:
        wait_until_closed()
    assert bf.same(got[setup["pv"]], setup["original"]), \
        f"El archivo se guardó con la línea base {setup['pv']} cambiada: {got[setup['pv']].dates()}"


def test_establish_copies_the_current_values_to_a_free_baseline(setup):
    """--establecer en una copia cerrada: la línea base libre recibe los valores actuales (comienzo, fin, duración,
    trabajo y coste) y la línea base del PV no cambia."""
    mpp = setup["folder"] / "establecer.mpp"
    shutil.copyfile(MPP, mpp)
    with ComSession(str(mpp)) as s:
        current = bf.read_values(bf.ComStore(s).tasks[setup["uid"]], None)
    if not setup["was_running"]:
        wait_until_closed()
    rc, text = _run("corregir-linea-base", "--origen", mpp, "--config", config_path(), "--tarea", setup["code"],
                    "--establecer", setup["free"], "--motivo", "Prueba de establecer", "--continuar",
                    "--sin-respaldo", "--informe", setup["folder"] / "establecer.md")
    assert rc == 0, f"corregir-linea-base terminó con código {rc} (mensajes: {_codes_of(text)})"
    got = _values(mpp, setup["uid"], [setup["pv"], setup["free"]])
    if not setup["was_running"]:
        wait_until_closed()
    assert bf.same(got[setup["free"]], current, numbers=False), \
        f"Fechas establecidas: {got[setup['free']].dates()}; se esperaban las actuales {current.dates()}"
    print(f"\nEstablecer: duración, trabajo y coste "
          f"{'iguales a los actuales' if bf.same(got[setup['free']], current) else 'recalculados (LBA-020)'}")
    assert bf.same(got[setup["pv"]], setup["original"]), "Cambió la línea base del PV"
    assert got["notes"].replace("\r", "\n").rstrip().endswith(
        f"Línea base {setup['free']} establecida con los valores actuales: Prueba de establecer")


def test_na_text_and_summary_baseline_can_be_written(setup):
    """Texto de NA que acepta este MS Project y escritura de los cinco campos en una tarea resumen (en una copia
    abierta que se cierra sin guardar)."""
    mpp = setup["folder"] / "resumen.mpp"
    shutil.copyfile(MPP, mpp)
    with ComSession(str(mpp), read_only=False) as s:
        project = s.read_project(setup["cfg"], links=False, review=True)
        store = bf.ComStore(s)
        writer = bf.ComWriter(s, store)
        task = next(t for t in project.tasks if t.uid == setup["uid"])
        original = bf.read_values(store.tasks[task.uid], setup["pv"])
        writer.write_values([(task, setup["free"], original)])
        writer.write_values([(task, setup["free"], bf.EMPTY)])
        print(f"\nTexto de NA aceptado por MS Project: {writer.na}")
        assert writer.na in ("NA", "ND")
        if setup["summary"] is not None:
            summary = next(t for t in project.tasks if t.uid == setup["summary"])
            values = bf.read_values(store.tasks[summary.uid], setup["pv"])
            writer.write_values([(summary, setup["free"], values)])
            got = bf.read_values(store.tasks[summary.uid], setup["free"])
            print(f"Tarea resumen {setup['summary_code']} (UID {summary.uid}): escrita; trabajo y coste "
                  f"{'iguales' if bf.same(got, values) else 'recalculados por MS Project (LBA-020)'}")
            assert bf.same(got, values, numbers=False)
    if not setup["was_running"]:
        wait_until_closed()
