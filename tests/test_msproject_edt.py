"""Pruebas opcionales con MS Project real de actualizar-edt: escritura de los códigos en una copia del .mpp cerrada
(con respaldo y guardado) y en una copia abierta en MS Project (sin guardar, con una sola acción de deshacer).

Se ejecutan en un PC con MS Project, con el campo del código EDT en un campo personalizado ([campos] edt, por
ejemplo "Item EDT"):

    $env:EVM_MPP = "C:\\Proyectos\\Programa.mpp"
    $env:EVM_CONFIG = "C:\\Proyectos\\config.toml"
    $env:EVM_EDT_DESDE = "2.7.1"        # opcional: por defecto, la primera tarea de nivel_maximo con cambios
    python -m pytest tests/test_msproject_edt.py -m msproject -v

El archivo de EVM_MPP no se modifica: las pruebas trabajan con copias en la carpeta temporal de pytest. Los mensajes
solo muestran UID y códigos EDT, nunca nombres de tareas.
"""
import contextlib
import copy
import hashlib
import io
import os
import shutil
import sys
from pathlib import Path

import pytest

from evm.cli import main
from evm.config import ConfigError, load_config
from evm.edt_renumber import check_field, plan_renumber
from evm.reader_com import ComSession
from evm.wbs import build_tree, code_level, select_scope

from test_msproject_real import (MPP, Hidden, config_path, ms_project_running, quiet_faulthandler,  # noqa: F401
                                 wait_until_closed)

START = os.environ.get("EVM_EDT_DESDE", "")
LIMIT = 15

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


def _codes(project) -> dict[int, str]:
    return {t.uid: (t.wbs or "").strip() for t in project.tasks}


def _start(project, cfg) -> str:
    """EVM_EDT_DESDE o la primera tarea de nivel_maximo con subtareas cuyos códigos cambian sin errores."""
    if START:
        return START
    scope = select_scope(build_tree(project.tasks), cfg.plan)
    fallback = ""
    for t in scope.tasks():
        raw = scope.raw_code(t)
        if not t.children or not raw or code_level(raw, cfg.ignore_interior_zeros) != cfg.max_level:
            continue
        try:
            plan = plan_renumber(copy.deepcopy(project), cfg, [raw])
        except ConfigError:
            continue
        if not plan.errors:
            if plan.changes:
                return raw
            fallback = fallback or raw
    return fallback


def _diffs(got: dict[int, str], expected: dict[int, str]) -> list[str]:
    return [f"UID {uid}: '{got.get(uid)}' (se esperaba '{code}')" for uid, code in expected.items()
            if got.get(uid) != code]


def _report(diffs: list[str]) -> str:
    more = f"\n... y {len(diffs) - LIMIT} más" if len(diffs) > LIMIT else ""
    return f"{len(diffs)} códigos distintos:\n" + "\n".join(diffs[:LIMIT]) + more


@pytest.fixture(scope="module")
def setup(tmp_path_factory):
    pytest.importorskip("win32com.client", reason="Se necesita pywin32")
    assert Path(MPP).is_file(), f"No existe el archivo {MPP}"
    cfg = load_config(config_path())
    d = tmp_path_factory.mktemp("edt")
    print(f"\nArchivos de la prueba: {d}")
    data = Hidden(was_running=ms_project_running(), folder=d, source_hash=_hash(MPP))
    data["copy"] = d / "edt.mpp"
    shutil.copyfile(MPP, data["copy"])
    with ComSession(str(data["copy"])) as s:
        project = s.read_project(cfg, links=False)
    if not data["was_running"]:
        wait_until_closed()
    try:
        check_field(project, cfg)
    except ConfigError as exc:
        pytest.skip(f"El código EDT no está en un campo personalizado ({exc.id}).")
    start = _start(project, cfg)
    if not start:
        pytest.skip("No hay ninguna tarea de nivel_maximo cuyas subtareas se puedan numerar sin errores.")
    plan = plan_renumber(copy.deepcopy(project), cfg, [start])
    before = _codes(project)
    expected = dict(before)
    expected.update({c.task.uid: c.new for c in plan.changes})
    data.update(cfg=cfg, start=start, before=before, expected=expected, changes=len(plan.changes))
    print(f"--desde {start}: {len(plan.changes)} códigos cambian")
    return data


def test_closed_file_is_written_saved_and_backed_up(setup):
    d, mpp = setup["folder"], setup["copy"]
    rc, text = _run("actualizar-edt", "--origen", mpp, "--config", config_path(), "--desde", setup["start"],
                    "--informe", d / "primera.md")
    assert rc == 0, f"actualizar-edt terminó con código {rc}"
    backups = list(d.glob("edt.respaldo-*.mpp"))
    assert len(backups) == (1 if setup["changes"] else 0)
    with ComSession(str(mpp)) as s:
        got = _codes(s.read_project(setup["cfg"], links=False))
    if not setup["was_running"]:
        wait_until_closed()
    diffs = _diffs(got, setup["expected"])
    assert not diffs, _report(diffs)
    assert _hash(MPP) == setup["source_hash"], "Se modificó el archivo de EVM_MPP"


def test_second_run_has_no_changes(setup):
    rc, text = _run("actualizar-edt", "--origen", setup["copy"], "--config", config_path(), "--desde",
                    setup["start"], "--simular", "--informe", setup["folder"] / "segunda.md")
    assert rc == 0 and "no hay cambios" in text


def test_open_file_is_not_saved_and_one_undo_restores_it(setup):
    if not setup["changes"]:
        pytest.skip("Los códigos de la rama ya están actualizados.")
    mpp = setup["folder"] / "abierto.mpp"
    shutil.copyfile(MPP, mpp)
    disk = _hash(mpp)
    with ComSession(str(mpp), read_only=False) as outer:
        rc, text = _run("actualizar-edt", "--origen", mpp, "--config", config_path(), "--desde", setup["start"],
                        "--informe", setup["folder"] / "abierto.md")
        assert rc == 0, f"actualizar-edt terminó con código {rc}"
        assert "sin guardar" in text and not list(setup["folder"].glob("abierto.respaldo-*.mpp"))
        got = _codes(outer.read_project(setup["cfg"], links=False))
        diffs = _diffs(got, setup["expected"])
        assert not diffs, _report(diffs)
        assert _hash(mpp) == disk, "El archivo abierto se guardó"
        assert "Ctrl+Z" in text, "Los cambios no se agruparon en una sola acción de deshacer (NUM-017)"
        outer.app.EditUndo()
        got = _codes(outer.read_project(setup["cfg"], links=False))
        diffs = _diffs(got, setup["before"])
        assert not diffs, "Un Ctrl+Z no deshizo todos los cambios. " + _report(diffs)
    if not setup["was_running"]:
        wait_until_closed()
