"""Pruebas opcionales con MS Project real de generar-cliente: el .mpp del cliente generado a partir de la planificación
interna.

Se ejecutan en un PC con MS Project, con las mismas variables de entorno que test_msproject_real.py:

    $env:EVM_MPP = "C:\\Proyectos\\Programa.mpp"
    $env:EVM_MPP_XML = "C:\\Proyectos\\Programa.xml"     # opcional aquí: compara la lectura del .mpp y del XML
    $env:EVM_CONFIG = "C:\\Proyectos\\config.toml"       # con [cliente] campo (y planificacion vacía si hay una)
    python -m pytest tests/test_msproject_cliente.py -m msproject -v

El archivo del cliente se genera con el comando y la configuración de EVM_CONFIG ([cliente] actividad_interna: por
defecto, las tareas que no se muestran quedan como «Otras tareas»), y una vez más, con la función, con el otro valor de
actividad_interna. Cada prueba deja en la carpeta temporal de pytest el archivo del cliente, el informe de control y el
registro de diagnóstico (la ruta se muestra con -s). Los mensajes solo muestran UID, códigos EDT, fechas y cantidades:
nunca nombres de tareas.
"""
import contextlib
import dataclasses
import hashlib
import io
import os
import sys
from datetime import datetime
from pathlib import Path

import pytest

from evm.cli import main, open_project
from evm.client_check import compare
from evm.client_plan import INTERNAL_NAME, build_client_plan, write_control
from evm.client_writer import generate_client
from evm.config import load_config
from evm.reader_com import ComSession

from test_msproject_real import (MPP, XML, Hidden, config_path, ms_project_running, quiet_faulthandler,  # noqa: F401
                                 wait_until_closed)

LIMIT = 15


def _client_config():
    cfg = load_config(config_path())
    return cfg if cfg.client_field else None


pytestmark = [
    pytest.mark.msproject,
    pytest.mark.skipif(sys.platform != "win32", reason="MS Project solo funciona en Windows"),
    pytest.mark.skipif(not MPP, reason='Falta definir EVM_MPP (en PowerShell: $env:EVM_MPP = "C:\\ruta\\Programa.mpp").'),
    pytest.mark.skipif(bool(MPP) and _client_config() is None,
                       reason="La configuración de EVM_CONFIG no tiene [cliente] campo (por ejemplo campo = "
                              "\"Etiqueta Cliente\")."),
]


def _hash(path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _run(*args) -> tuple[int, str]:
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        rc = main([str(a) for a in args])
    return rc, out.getvalue() + err.getvalue()


def _names(project) -> set[str]:
    """Nombres de las tareas que no deben aparecer en el registro ni en el informe (los de 6 o más letras, para no
    confundirlos con palabras de los mensajes)."""
    return {t.name.strip() for t in project.tasks if len(t.name.strip()) >= 6}


@pytest.fixture(scope="module")
def generated(tmp_path_factory):
    """Simula con el .mpp, genera el archivo del cliente y lo vuelve a leer."""
    pytest.importorskip("win32com.client", reason="Se necesita pywin32")
    assert Path(MPP).is_file(), f"No existe el archivo {MPP}"
    cfg = _client_config()
    d = tmp_path_factory.mktemp("cliente")
    print(f"\nArchivos de la prueba: {d}")
    result = Hidden(was_running=ms_project_running(), folder=d)
    before = (_hash(MPP), os.path.getmtime(MPP))
    with ComSession(MPP) as session:
        project = session.read_project(cfg, client=True)
        result["plan"] = build_client_plan(project, cfg)
        result["names"] = _names(project)
        result["field"] = session._resolve_field(cfg.client_field)      # ('attr', 'Text15') o ('fid', id)
    if not result["was_running"]:
        wait_until_closed()
    destino, log = d / "cliente.mpp", d / "registro.log"
    rc, text = _run("--registro", log, "generar-cliente", "--origen", MPP, "--config", config_path(),
                    "--destino", destino)
    result["rc"], result["output"], result["log"], result["destino"] = rc, text, log, destino
    result["report"] = d / "cliente_control.md"
    result["source_after"] = (_hash(MPP), os.path.getmtime(MPP))
    result["source_before"] = before
    if not result["was_running"]:
        result["closed"] = wait_until_closed()
    if rc == 0:
        _read_client(result, cfg)
    return result


def _read_client(result, cfg) -> None:
    """Lee el archivo del cliente generado (otra sesión de MS Project) y lo compara con la planificación interna."""
    with open_project(str(result["destino"]), cfg, "com") as client:
        result["client"] = client
        result["tasks"] = {t.uid: t for t in client.tasks}
    if not result["was_running"]:
        wait_until_closed()
    result["comparison"] = compare(result["plan"], result["client"], cfg, generated=True)


@pytest.fixture(scope="module")
def other_mode(generated, tmp_path_factory):
    """El archivo del cliente con el otro valor de actividad_interna, generado con la función (no con el comando)."""
    cfg = dataclasses.replace(_client_config(), client_internal_activity=not generated["plan"].internal_activity)
    d = tmp_path_factory.mktemp("cliente_otro_modo")
    result = Hidden(was_running=ms_project_running(), folder=d, destino=d / "cliente.mpp")
    try:
        plan, outcome = generate_client(MPP, result["destino"], cfg)
    except Exception as exc:                        # el mensaje no lleva nombres de tareas (CLI-018)
        result["error"] = f"{type(exc).__name__}: {exc}"
        return result
    finally:
        if not result["was_running"]:
            wait_until_closed()
    result["plan"], result["outcome"] = plan, outcome
    write_control(plan, MPP, d / "cliente_control.md", "Prueba", outcome=outcome)
    _read_client(result, cfg)
    return result


def _summary(comp) -> str:
    b = comp.book
    return (f"% del proyecto: interno {comp.percent_internal}, cliente {comp.percent_client}; libro EVM: {b.same} "
            f"filas iguales, {len(b.moved)} con otro código {[(x.code, y.code) for x, y in b.moved][:LIMIT]}, "
            f"{len(b.different)} con otras tareas o valores {[x.code for x, _ in b.different][:LIMIT]}, solo interno "
            f"{[x.code for x in b.only_internal][:LIMIT]}, solo cliente {[x.code for x in b.only_client][:LIMIT]}, "
            f"-D {b.d_rows[:LIMIT]}{' ' + b.error if b.error else ''}")


def test_generation_ends_without_errors(generated):
    assert generated["rc"] == 0, (f"generar-cliente terminó con código {generated['rc']}. El registro de diagnóstico "
                                  f"está en {generated['log']}.")
    assert generated["destino"].is_file() and not os.access(generated["destino"], os.W_OK)


def test_source_is_not_modified(generated):
    assert generated["source_after"] == generated["source_before"], "El archivo de origen cambió."


def test_client_file_has_the_expected_tasks(generated):
    """Las que se muestran y, con actividad_interna, también las «Otras tareas» (sin las inactivas que no se
    muestran)."""
    plan = generated["plan"]
    expected = plan.remaining
    found = set(generated.get("tasks", {}))
    extra, missing = sorted(found - expected), sorted(expected - found)
    assert not extra and not missing, (f"Tareas que sobran (UID): {extra[:LIMIT]}; tareas que faltan (UID): "
                                       f"{missing[:LIMIT]}")


def test_other_tasks_have_no_name_nor_code(generated):
    plan, tasks = generated["plan"], generated.get("tasks", {})
    if not plan.internal_activity:
        pytest.skip("actividad_interna = false: las tareas que no se muestran se borran")
    named = [t.uid for t in plan.anonymous if t.uid in tasks and tasks[t.uid].name != INTERNAL_NAME]
    coded = [t.uid for t in plan.anonymous if t.uid in tasks and tasks[t.uid].wbs] if plan.codes_removed else []
    renamed = [t.uid for t in plan.tasks if t.uid in plan.shown and t.uid in tasks and tasks[t.uid].name != t.name]
    assert not named, f"{len(named)} tareas que no se muestran conservan su nombre (UID): {named[:LIMIT]}"
    assert not coded, f"{len(coded)} «{INTERNAL_NAME}» conservan su código EDT (UID): {coded[:LIMIT]}"
    assert not renamed, f"{len(renamed)} tareas que se muestran cambiaron de nombre (UID): {renamed[:LIMIT]}"


def test_client_field_is_emptied(generated):
    cfg = load_config(config_path())
    with open_project(str(generated["destino"]), cfg, "com", client=True) as client:
        found = "cliente" in client.fields_found
    if not generated["was_running"]:
        wait_until_closed()
    kind, key = generated["field"]
    # Los valores se leen por el campo (Text15...), no por su nombre, que ya no debería existir
    by_field = dataclasses.replace(cfg, client_field=str(key)) if kind == "attr" else cfg
    with open_project(str(generated["destino"]), by_field, "com", client=True) as client:
        with_value = [t.uid for t in client.tasks if t.client_mark]
    if not generated["was_running"]:
        wait_until_closed()
    assert not with_value, f"Tareas con valor en el campo del cliente (UID): {with_value[:LIMIT]}"
    assert not found, ("El campo del cliente sigue teniendo su nombre en el archivo generado (CLI-017): ver el "
                       "registro de diagnóstico.")


def test_every_task_keeps_its_data(generated):
    """actividad_interna: todas las tareas del archivo del cliente (también los resúmenes y las «Otras tareas») tienen
    las fechas, el % completado, la duración y la línea base de la planificación interna."""
    if not generated["plan"].internal_activity:
        pytest.skip("actividad_interna = false: ver test_dates_match_the_internal_plan")
    comp = generated["comparison"]
    lines = [f"UID {d.task.uid}: " + "; ".join(f"{label} {a} -> {b}" for label, a, b in d.fields)
             for d in comp.tasks]
    assert comp.checked and not lines, f"{len(lines)} tareas con datos distintos:\n" + "\n".join(lines[:LIMIT])


def test_project_percent_and_evm_book(generated):
    """actividad_interna: el % completado del proyecto y cada fila del libro EVM coinciden con la planificación
    interna (las filas que cambian de código, como las -D, tienen las mismas tareas y valores)."""
    comp = generated["comparison"]
    print("\n" + _summary(comp))
    if not generated["plan"].internal_activity:
        pytest.skip("actividad_interna = false: los % pueden cambiar (el informe de control indica las diferencias)")
    assert not comp.percent_differs, _summary(comp)
    assert not comp.book.error and comp.book.deviations == 0, _summary(comp)


def test_the_other_mode(other_mode):
    """El otro valor de actividad_interna también genera el archivo, con las tareas esperadas y la comparación."""
    assert "error" not in other_mode, other_mode.get("error")
    plan, tasks, comp = other_mode["plan"], other_mode["tasks"], other_mode["comparison"]
    print("\n" + ("«Otras tareas»" if plan.internal_activity else "borrado") + ": " + _summary(comp))
    extra, missing = sorted(set(tasks) - plan.remaining), sorted(plan.remaining - set(tasks))
    assert not extra and not missing, f"Tareas que sobran: {extra[:LIMIT]}; tareas que faltan: {missing[:LIMIT]}"
    if plan.internal_activity:
        assert not comp.tasks and not comp.percent_differs and comp.book.deviations == 0, _summary(comp)


def test_dates_match_the_internal_plan(generated):
    """actividad_interna = false: las tareas que se muestran (sin los resúmenes) tienen las fechas de la planificación
    interna, salvo las que el informe de control explica (comenzadas, con otra delimitación...)."""
    plan, tasks = generated["plan"], generated.get("tasks", {})
    if plan.internal_activity:
        pytest.skip("actividad_interna = true: ver test_every_task_keeps_its_data")
    collapsed = {t.uid for t in plan.collapsed}
    report = generated["report"].read_text(encoding="utf-8")
    diffs = []
    for t in plan.tasks:
        if t.uid not in plan.shown or (t.children and t.uid not in collapsed) or t.uid not in tasks:
            continue
        c = tasks[t.uid]
        if (c.start, c.finish) != (t.start, t.finish) and f"(UID {t.uid})" not in report:
            diffs.append(f"UID {t.uid}: interna {t.start:%d/%m/%Y %H:%M} - {t.finish:%d/%m/%Y %H:%M}, cliente "
                         f"{c.start:%d/%m/%Y %H:%M} - {c.finish:%d/%m/%Y %H:%M}")
    assert not diffs, f"{len(diffs)} tareas con fechas distintas que el informe no explica:\n" + "\n".join(diffs[:LIMIT])


def test_progress_is_kept_and_summaries_as_predicted(generated):
    """actividad_interna = false: las tareas de detalle y los resúmenes convertidos conservan su % completado; el de los
    demás resúmenes lo calcula MS Project con las tareas que quedan y debe coincidir con el previsto (CLI-020)."""
    plan, tasks = generated["plan"], generated.get("tasks", {})
    if plan.internal_activity:
        pytest.skip("actividad_interna = true: ver test_every_task_keeps_its_data")
    collapsed = {t.uid for t in plan.collapsed}
    expected = {c.task.uid: c.client for c in plan.progress_changes}
    lost, unexpected = [], []
    for t in plan.tasks:
        if t.uid not in plan.shown or t.uid not in tasks:
            continue
        got = tasks[t.uid].percent_complete
        if t.children and t.uid not in collapsed:
            want = expected.get(t.uid, t.percent_complete)
            if abs(got - want) >= 1:
                unexpected.append(f"UID {t.uid}: interno {t.percent_complete:g} %, previsto {want:.1f} %, MS Project "
                                  f"{got:g} %")
        elif abs(got - t.percent_complete) >= 0.5:
            lost.append(f"UID {t.uid}: interno {t.percent_complete:g} %, cliente {got:g} %")
    assert not lost, f"{len(lost)} tareas no conservan su % completado:\n" + "\n".join(lost[:LIMIT])
    assert not unexpected, (f"{len(unexpected)} resúmenes con un % distinto del previsto:\n"
                            + "\n".join(unexpected[:LIMIT]))


def test_log_and_report_have_no_task_names(generated):
    texts = {"registro": generated["log"].read_text(encoding="utf-8"),
             "informe": generated["report"].read_text(encoding="utf-8")}
    found = [f"{where}: un nombre de tarea" for where, text in texts.items()
             for name in generated["names"] if name in text]
    assert not found, f"{len(found)} nombres de tareas en el registro o el informe (no se muestran)."


def test_ms_project_is_left_as_found(generated):
    if generated["was_running"]:
        assert ms_project_running()
    else:
        assert generated.get("closed", True), "MS Project sigue abierto después de generar el archivo del cliente."


def test_overwrite_replaces_the_read_only_file(generated, tmp_path):
    destino = generated["destino"]
    stamp = datetime.fromtimestamp(os.path.getmtime(destino))
    rc, text = _run("--registro", tmp_path / "registro2.log", "generar-cliente", "--origen", MPP, "--config",
                    config_path(), "--destino", destino, "--sobrescribir")
    if not generated["was_running"]:
        wait_until_closed()
    assert rc == 0, f"generar-cliente --sobrescribir terminó con código {rc}."
    assert datetime.fromtimestamp(os.path.getmtime(destino)) >= stamp and not os.access(destino, os.W_OK)


@pytest.mark.skipif(not XML, reason="Falta EVM_MPP_XML: no se compara la lectura del .mpp con la del XML.")
def test_same_plan_with_mpp_and_xml():
    cfg = _client_config()
    plans = {}
    for engine, origin in (("com", MPP), ("xml", XML)):
        with open_project(origin, cfg, engine, client=True) as project:
            plans[engine] = build_client_plan(project, cfg)
    com, xml = plans["com"], plans["xml"]
    assert com.shown == xml.shown, (f"Tareas que se muestran: COM {len(com.shown)} / XML {len(xml.shown)}; distintas "
                                    f"(UID): {sorted(com.shown ^ xml.shown)[:LIMIT]}")
    assert com.remaining == xml.remaining, f"Tareas del archivo del cliente: COM {len(com.remaining)} / XML " \
                                           f"{len(xml.remaining)}"
    a = {(x.pred.uid, x.succ.uid, x.kind, round(x.lag)) for x in com.new_links}
    b = {(x.pred.uid, x.succ.uid, x.kind, round(x.lag)) for x in xml.new_links}
    assert a == b, (f"Vínculos nuevos: COM {len(a)} / XML {len(b)}; solo COM {sorted(a - b)[:LIMIT]}; solo XML "
                    f"{sorted(b - a)[:LIMIT]}")
