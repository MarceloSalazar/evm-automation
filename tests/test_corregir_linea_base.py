"""corregir-linea-base: mover o borrar la línea base de tareas (con su rama), avisos que piden confirmación,
escritura con un MS Project simulado (en inglés, NA, y en español, ND), --interactivo caso a caso con deshacer, y el
comando con un XML (--simular). Las pruebas con MS Project real están en test_msproject_linea_base.py."""
from __future__ import annotations

import contextlib
import io
from datetime import datetime

import pytest

from evm import baseline_fix as bf
from evm.cli import main
from evm.config import Config, ConfigError
from evm.fields import TASK_TEXT_IDS
from evm.messages import CATALOG, ProjectAppError
from evm.wbs import build_tree, select_scope

from builders import P, T
from mspdi import project_xml, task_xml, write_xml

JAN = (datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17))
FEB = (datetime(2026, 2, 2, 8), datetime(2026, 2, 27, 17))
MAR = (datetime(2026, 3, 2, 8), datetime(2026, 3, 27, 17))
Q1 = (JAN[0], MAR[1])


def _task(uid, level, code, name, slots, pv=1):
    """Tarea con las fechas de cada línea base (slots: {número: (comienzo, fin)}); las de la línea base del PV, en
    baseline_start/baseline_finish (como los lectores)."""
    own = slots.get(pv)
    t = T(uid, level, code, name, bs=own[0] if own else None, bf=own[1] if own else None)
    t.baseline_dates = dict(slots)
    t.baselines = tuple(sorted(slots))
    t.baseline_duration = 9600.0 if own else None
    return t


def _project(pv=1):
    """PV con la línea base 1; la 0 es la aprobada. 1.1.2 (con sus subtareas) y 1.2 son erratas: solo están en la 1.
    1.3 se aprobó como una tarea y se desglosó después en 1.3.1, guardada por error en la 1."""
    tasks = [
        _task(1, 1, "1", "Fase", {0: Q1, 1: Q1}, pv),
        _task(2, 2, "1.1", "Paquete", {0: Q1, 1: Q1}, pv),
        _task(3, 3, "1.1.1", "Existente", {0: JAN, 1: JAN}, pv),
        _task(4, 3, "1.1.2", "Nueva", {1: (FEB[0], MAR[1])}, pv),
        _task(5, 4, "1.1.2.1", "Nueva A", {1: FEB}, pv),
        _task(6, 4, "1.1.2.2", "Nueva B", {1: MAR}, pv),
        _task(7, 2, "1.2", "Paquete nuevo", {1: FEB}, pv),
        _task(8, 3, "1.2.1", "Solo", {1: FEB}, pv),
        _task(9, 2, "1.3", "Desglosada", {0: MAR, 1: MAR}, pv),
        _task(10, 3, "1.3.1", "Detalle", {1: MAR}, pv),
    ]
    return P(tasks, baseline_number=pv, ev_baseline=pv)


class DictStore:
    """Valores de las líneas base a partir de las fechas de las tareas (duración de 20 días, trabajo y costo)."""

    def __init__(self, project, notes=None):
        self.project = project
        self._notes = notes or {}

    def values(self, uid, k):
        t = next(t for t in self.project.tasks if t.uid == uid)
        got = t.baseline_dates.get(k)
        return bf.Values(got[0], got[1], 9600.0, 9600.0, 1000.0 + uid) if got else bf.EMPTY

    def notes(self, uid):
        return self._notes.get(uid, "")


def _setup(pv=1, approved=0):
    project = _project(pv)
    cfg = Config(max_level=2)
    cfg.validate()
    scope = select_scope(build_tree(project.tasks), cfg.plan)
    state = bf.State(project, DictStore(project))
    return project, scope, state, bf.Context(project, cfg, pv, approved)


def _ids(case):
    return [i.id for i in case.issues]


# ----------------------------------------------------------------------------
# Acción y tareas
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("de, a, borrar", [(None, None, None), (1, None, None), (1, 1, None), (1, 2, 3)])
def test_action_must_be_move_or_delete(de, a, borrar):
    with pytest.raises(ConfigError) as exc:
        bf.parse_action(de, a, borrar)
    assert exc.value.id == "LBA-005"


def test_action_texts_and_note():
    move, delete = bf.parse_action(2, 10, None), bf.parse_action(None, None, 2)
    assert move.text() == "mover de la línea base 2 a la línea base 10" and delete.text() == "borrar la línea base 2"
    when = datetime(2026, 11, 21)
    assert move.note(" No comunicada ", when) == "21/11/2026 · Línea base 2 movida a línea base 10: No comunicada"
    assert delete.note("Error", when) == "21/11/2026 · Línea base 2 borrada: Error"


def test_targets_need_branch_for_a_summary_and_must_not_overlap():
    _p, scope, _s, _c = _setup()
    assert [t.name for t in bf.find_targets(scope, ["1.1.2.1", "1.2.1"], False)] == ["Nueva A", "Solo"]
    for codes, branch, code in [(["1.1.2"], False, "LBA-004"), (["9.9"], True, "LBA-001"),
                                (["1.1.2", "1.1.2.1"], True, "LBA-003"), (["1.2", "1.2.0"], True, "LBA-003")]:
        with pytest.raises(ConfigError) as exc:
            bf.find_targets(scope, codes, branch)
        assert exc.value.id == code


# ----------------------------------------------------------------------------
# Cambios y avisos
# ----------------------------------------------------------------------------
def test_move_copies_to_the_destination_and_leaves_the_origin_empty():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.1.2"), True, bf.Action(1, 10))
    assert case.members == 3 and len(case.changes) == 6 and not case.skipped
    by = {(c.task.wbs, c.slot): c for c in case.changes}
    assert by[("1.1.2.1", 1)].after.empty and by[("1.1.2.1", 10)].after.start == FEB[0]
    assert by[("1.1.2.1", 10)].before.empty
    bf.check_case(case, state, ctx)
    # La fila 1.1 conserva 1.1.1 en el PV: su PV cambia, no se queda sin PV; 1 conserva otras tareas: sin bloque.
    # 1.1 sigue de enero a marzo en la línea base 1, con solo 1.1.1 (enero): se informa, no se modifica
    assert _ids(case) == ["LBA-018", "LBA-017"] and "1.1" in case.issues[0].message and not case.confirm
    assert case.issues[1].code == "1.1"


def test_tasks_without_the_origin_baseline_do_not_change():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.1"), True, bf.Action(0, 2))
    assert [t.wbs for t in case.skipped] == ["1.1.2", "1.1.2.1", "1.1.2.2"]
    assert {c.task.wbs for c in case.changes} == {"1.1", "1.1.1"}
    bf.check_case(case, state, ctx)
    assert "LBA-016" in _ids(case)


def test_a_row_left_without_pv_and_a_destination_with_values_ask_for_confirmation():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.2"), True, bf.Action(1, 0))
    bf.check_case(case, state, ctx)
    assert sorted(_ids(case)) == ["LBA-014", "LBA-018"] and "1.2" in case.issues[0].message
    assert case.issues[0].message.startswith("La fila 1.2 se queda sin ninguna tarea")
    # Mover 1.1.1 a la 0, que ya tiene: el destino tiene valores
    case = bf.plan_branch(state, bf.find_task(scope, "1.1.1"), False, bf.Action(1, 0))
    bf.check_case(case, state, ctx)
    assert "LBA-012" in _ids(case) and "LBA-012" in [i.id for i in case.confirm]


def test_moving_into_the_pv_baseline_asks_for_confirmation():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.1.1"), False, bf.Action(0, 1))
    bf.check_case(case, state, ctx)
    assert {"LBA-012", "LBA-013"} <= set(_ids(case))


def test_a_summary_above_left_alone_with_the_baseline_is_a_block():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.3.1"), False, bf.Action(1, 10))
    bf.check_case(case, state, ctx)
    block = next(i for i in case.issues if i.id == "LBA-015")
    assert block.code == "1.3" and "bloque" in block.message
    assert "--tarea 1.3 --rama" in bf.options_text(block)
    # Sin el aviso del bloque (--interactivo lo pregunta aparte)
    bf.check_case(case, state, ctx, blocks=False)
    assert "LBA-015" not in _ids(case)
    # Con la resumen y su rama no queda ningún bloque, y la fila 1.3 se queda sin PV
    case = bf.plan_branch(state, bf.find_task(scope, "1.3"), True, bf.Action(1, 10))
    bf.check_case(case, state, ctx)
    assert "LBA-015" not in _ids(case) and "LBA-014" in _ids(case)


def test_summaries_above_that_no_longer_match_are_reported_not_changed():
    _p, scope, state, ctx = _setup()
    case = bf.plan_branch(state, bf.find_task(scope, "1.1.2.2"), False, bf.Action(1, 10))
    bf.check_case(case, state, ctx)
    notes = [i for i in case.issues if i.id == "LBA-017"]
    # 1.1.2 (febrero a marzo) queda con febrero; 1.1 (enero a marzo), con enero y febrero; 1 sigue de enero a marzo
    assert [i.code for i in notes] == ["1.1.2", "1.1"] and "27/02/2026" in notes[0].message
    case = bf.plan_branch(state, bf.find_task(scope, "1.2.1"), False, bf.Action(1, 10))
    bf.check_case(case, state, ctx)
    assert "LBA-017" not in _ids(case) and "LBA-015" in _ids(case)


def test_approved_baseline_that_is_not_the_pv_one():
    _p, _scope, _state, ctx = _setup(pv=1, approved=0)
    issue = bf.approved_issue(ctx)
    assert issue.id == "LBA-011" and "linea_base = 0" in issue.message
    ctx.approved = 1
    assert bf.approved_issue(ctx) is None


def test_confirmation_warnings_cannot_be_silenced():
    assert all(not CATALOG[c].silenceable for c in bf.CONFIRM)


# ----------------------------------------------------------------------------
# Escritura con un MS Project simulado
# ----------------------------------------------------------------------------
def _naive(value):
    """Como COM: una fecha con zona horaria se lee sin ella."""
    if isinstance(value, datetime):
        return datetime(value.year, value.month, value.day, value.hour, value.minute)
    return value


class FakeCom:
    """Tarea de MS Project: las fechas vacías valen «NA» (inglés) o «ND» (español); el otro texto da error."""

    def __init__(self, uid, slots, notes="", na="NA"):
        d = self.__dict__
        d.update(UniqueID=uid, ExternalTask=False, Notes=notes, _na=na, writes=[])
        for k in range(11):
            sfx = "" if k == 0 else str(k)
            got = slots.get(k)
            d[f"Baseline{sfx}Start"] = got[0] if got else na
            d[f"Baseline{sfx}Finish"] = got[1] if got else na
            d[f"Baseline{sfx}Duration"] = 9600.0 if got else 0.0
            d[f"Baseline{sfx}Work"] = 9600.0 if got else 0.0
            d[f"Baseline{sfx}Cost"] = 1000.0 + uid if got else 0.0

    def __setattr__(self, name, value):
        if name.endswith(("Start", "Finish")) and isinstance(value, str) and value != self._na:
            raise ValueError("El valor no es válido")       # como MS Project con el texto de otro idioma
        self.writes.append(name)
        self.__dict__[name] = _naive(value)


class FakeApp:
    def __init__(self, undo=True):
        self.calls = []
        self.undo = undo
        self.Visible = False

    def OpenUndoTransaction(self, name):
        if not self.undo:
            raise RuntimeError("no disponible")
        self.calls.append(("open", name))

    def CloseUndoTransaction(self):
        self.calls.append(("close",))


class FakeProjectCom:
    def __init__(self, tasks):
        self.Tasks = tasks


class FakeSession:
    def __init__(self, com_tasks, **app):
        self.app = FakeApp(**app)
        self.project = FakeProjectCom(com_tasks)

    def _activate(self):
        pass


def _com_setup(na="NA", notes=None, pv=1, approved=0, **app):
    project, scope, _state, ctx = _setup(pv, approved)
    com_tasks = [FakeCom(t.uid, t.baseline_dates, (notes or {}).get(t.uid, ""), na) for t in project.tasks]
    session = FakeSession(com_tasks, **app)
    store = bf.ComStore(session)
    state = bf.State(project, store)
    return project, scope, state, ctx, session, store, {c.UniqueID: c for c in com_tasks}


@pytest.mark.parametrize("na", ["NA", "ND"])
def test_writer_moves_the_branch_in_one_undo_action_and_appends_the_reason(na):
    project, scope, state, ctx, session, store, com = _com_setup(na=na, notes={5: "Nota previa"})
    case = bf.plan_branch(state, bf.find_task(scope, "1.1.2"), True, bf.Action(1, 10))
    writer = bf.ComWriter(session, store)
    when = datetime(2026, 11, 21)
    assert bf.write_cases(writer, [case], state, "No comunicada al cliente", when) is True
    assert session.app.calls == [("open", bf.UNDO_NAME), ("close",)] and not writer.warnings
    assert com[5].Baseline1Start == na and com[5].Baseline1Finish == na and com[5].Baseline1Duration == 0
    assert com[5].Baseline10Start == FEB[0] and com[5].Baseline10Finish == FEB[1]
    assert com[5].Baseline10Cost == 1005.0 and writer.na == na
    assert com[5].Notes == "Nota previa\r21/11/2026 · Línea base 1 movida a línea base 10: No comunicada al cliente"
    assert com[6].Notes.startswith("21/11/2026 ·") and com[3].Notes == ""      # fuera de la rama no cambia
    assert case.notes_before == {4: "", 5: "Nota previa", 6: ""}
    # El estado refleja lo escrito: 1.1.2 ya no tiene la línea base 1
    assert not state.has(4, 1) and state.has(4, 10)


def test_writer_without_undo_transactions_warns():
    _project, scope, state, _ctx, session, store, _com = _com_setup(undo=False)
    case = bf.plan_branch(state, bf.find_task(scope, "1.2.1"), False, bf.Action(1, None))
    writer = bf.ComWriter(session, store)
    assert bf.write_cases(writer, [case], state, "Error", datetime(2026, 11, 21)) is False
    assert [i.id for i in writer.warnings] == ["LBA-019"]


def test_a_value_that_ms_project_does_not_keep_restores_the_previous_ones(monkeypatch):
    _project, scope, state, _ctx, session, store, com = _com_setup(notes={8: "Antes"})
    case = bf.plan_branch(state, bf.find_task(scope, "1.2.1"), False, bf.Action(1, 10))
    real = FakeCom.__setattr__

    def stubborn(self, name, value):          # MS Project mueve el comienzo al escribir el fin de la línea base 10
        real(self, name, value)
        if name == "Baseline10Finish" and not isinstance(value, str):
            self.__dict__["Baseline10Start"] = datetime(2027, 1, 1)

    monkeypatch.setattr(FakeCom, "__setattr__", stubborn)
    with pytest.raises(ProjectAppError) as exc:
        bf.write_cases(bf.ComWriter(session, store), [case], state, "Motivo", datetime(2026, 11, 21))
    assert exc.value.id == "LBA-010" and "no coinciden" in str(exc.value)
    monkeypatch.setattr(FakeCom, "__setattr__", real)
    assert com[8].Baseline1Start == FEB[0] and com[8].Notes == "Antes" and state.has(8, 1)


# ----------------------------------------------------------------------------
# --interactivo
# ----------------------------------------------------------------------------
def _interactive(answers, **kw):
    project, scope, state, ctx, session, store, com = _com_setup(**kw)
    out: list[str] = []
    feed = iter(answers)
    console = bf.Console(ask=lambda q: (out.append(q), next(feed))[1], out=out.append)
    run = bf.Session(console, ctx, state, bf.ComWriter(session, store), "Errata", datetime(2026, 11, 21))
    return project, scope, run, com, out


def test_interactive_writes_each_change_and_asks_for_the_summary_above():
    project, scope, run, com, out = _interactive(["s", "s", "a", "s", "s"])
    bf.run_interactive(run, [bf.find_task(scope, "1.3.1")], False, bf.Action(1, 10))
    # 1.3.1 pasa a la línea base 10; 1.3 queda con la 1 sin subtareas: se mueve a la aprobada (0), que ya tenía
    assert com[10].Baseline1Start == "NA" and com[10].Baseline10Start == MAR[0]
    assert [c.kind for c in run.cases] == ["rama", "resumen"]
    summary = run.cases[1]
    assert summary.action == bf.Action(1, 0) and "LBA-012" in _ids(summary)
    assert com[9].Baseline1Start == "NA" and com[9].BaselineStart == MAR[0]
    assert com[9].Notes.endswith("Línea base 1 movida a línea base 0: Errata")
    text = "\n".join(out)
    assert "a: mover a la línea base aprobada (0)" in text and "m: mover a otra línea base" in text
    assert [c.result for c in run.cases] == ["escrito en MS Project, sin guardar"] * 2


def test_interactive_undo_restores_and_asks_again():
    project, scope, run, com, out = _interactive(["s", "d", "o"], notes={8: "Antes"})
    bf.run_interactive(run, [bf.find_task(scope, "1.2.1")], False, bf.Action(1, 10))
    assert com[8].Baseline1Start == FEB[0] and com[8].Baseline10Start == "NA" and com[8].Notes == "Antes"
    assert [(c.decision, c.result) for c in run.cases] == [("aplicar", "escrito y deshecho"),
                                                           ("omitir", "no se aplicó")]
    assert run.state.has(8, 1) and not run.state.has(8, 10)


def test_interactive_other_baseline_keep_and_exit():
    # m: mover la resumen a otra línea base (se pide el número; el de origen no vale)
    project, scope, run, com, out = _interactive(["s", "s", "m", "1", "9", "s", "s"])
    bf.run_interactive(run, [bf.find_task(scope, "1.3.1")], False, bf.Action(1, 10))
    assert com[9].Baseline9Start == MAR[0] and com[9].Baseline1Start == "NA"
    assert any("Respuesta no válida" in line for line in out)
    # c: conservar la resumen
    project, scope, run, com, out = _interactive(["s", "s", "c"])
    bf.run_interactive(run, [bf.find_task(scope, "1.3.1")], False, bf.Action(1, 10))
    assert com[9].Baseline1Start == MAR[0] and run.cases[-1].decision == "omitir: se conserva"
    # x: salir conserva lo escrito
    project, scope, run, com, out = _interactive(["s", "x"])
    bf.run_interactive(run, [bf.find_task(scope, "1.3.1"), bf.find_task(scope, "1.2.1")], False, bf.Action(1, 10))
    assert com[10].Baseline10Start == MAR[0] and com[8].Baseline10Start == "NA" and len(run.cases) == 1
    assert "Salida a petición del usuario" in out[-1]


# ----------------------------------------------------------------------------
# Comando
# ----------------------------------------------------------------------------
TEXT5 = TASK_TEXT_IDS[5]


def _xml(tmp_path):
    def b(k, dates):
        return (k, dates[0], dates[1], 100.0, "PT160H0M0S")
    tasks = [task_xml(0, 0, "Proyecto"),
             task_xml(1, 1, "Fase", ext={TEXT5: "1"}, start=Q1[0], finish=Q1[1], baselines=[b(0, Q1), b(1, Q1)]),
             task_xml(2, 2, "Paquete", ext={TEXT5: "1.1"}, start=JAN[0], finish=JAN[1],
                      baselines=[b(0, JAN), b(1, JAN)]),
             task_xml(3, 2, "Nuevo", ext={TEXT5: "1.2"}, start=FEB[0], finish=FEB[1], baselines=[b(1, FEB)],
                      extra="<Notes>Nota</Notes>")]
    path = write_xml(tmp_path / "p.xml", project_xml(
        tasks, definitions=[(TEXT5, "Text5", "Item EDT")],
        header="<BaselineForEarnedValue>1</BaselineForEarnedValue>"))
    cfg = tmp_path / "c.toml"
    cfg.write_text('[campos]\nedt = "Item EDT"\n\n[exportacion]\nnivel_maximo = 2\n', encoding="utf-8")
    return path, cfg


def _run(*args):
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        rc = main([str(a) for a in args])
    return rc, out.getvalue()


def test_xml_store_reads_every_baseline_and_the_notes(tmp_path):
    xml, _cfg = _xml(tmp_path)
    store = bf.XmlStore(xml)
    v = store.values(3, 1)
    assert (v.start, v.finish, v.duration, v.cost) == (FEB[0], FEB[1], 9600.0, 100.0)
    assert store.values(3, 0).empty and store.notes(3) == "Nota"


def test_command_simulates_with_xml(tmp_path):
    xml, cfg = _xml(tmp_path)
    report = tmp_path / "r.md"
    rc, text = _run("corregir-linea-base", "--origen", xml, "--config", cfg, "--tarea", "1.2", "--de", "1", "--a",
                    "10", "--aprobada", "0", "--simular", "--informe", report)
    assert rc == 0, text
    flat = " ".join(text.split())
    assert "Simulación: no se escribió nada" in flat and "LBA-014" in flat and "LBA-011" in flat
    assert "la usa el PV; configurada en MS Project para el valor acumulado" in flat
    assert "Opciones de los avisos que piden confirmación" in flat
    body = report.read_text(encoding="utf-8")
    assert "## Casos" in body and "| 3 | 1.2 | Nuevo | 1 | 02/02/2026 – 27/02/2026" in body
    assert "| 1 | 3 |  | la usa el PV, MS Project: valor acumulado |" in body


@pytest.mark.parametrize("extra, code", [
    ([], "LBA-007"),                                    # un XML solo se simula
    (["--simular", "--interactivo"], "LBA-009"),
    (["--simular", "--de", "2"], "LBA-005"),            # --borrar y --de a la vez
])
def test_command_errors(tmp_path, extra, code):
    xml, cfg = _xml(tmp_path)
    rc, text = _run("corregir-linea-base", "--origen", xml, "--config", cfg, "--tarea", "1.2", "--borrar", "1",
                    *extra)
    assert rc == 2 and code in text, text


def test_command_summary_without_branch_is_an_error(tmp_path):
    xml, cfg = _xml(tmp_path)
    rc, text = _run("corregir-linea-base", "--origen", xml, "--config", cfg, "--tarea", "1", "--borrar", "1",
                    "--simular")
    assert rc == 2 and "LBA-004" in text


class FakeComSession:
    """ComSession simulada para el comando: el proyecto de _project() con sus tareas de MS Project."""

    last = None

    def __init__(self, origen, visible=False, read_only=True, opened=True):
        project, _scope, _state, _ctx, session, _store, com = _com_setup()
        self.data, self.app, self.project, self.com = project, session.app, session.project, com
        self.read_only, self.opened_by_us, self.started_by_us = read_only, opened, False
        self.saved = self.closed = False
        FakeComSession.last = self

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.closed = self.opened_by_us

    def read_project(self, cfg, links=True, review=False):
        assert review and not links
        return self.data

    def save(self):
        assert not self.read_only
        self.saved = True

    def _activate(self):
        pass

    def hand_over(self):
        self.opened_by_us = self.started_by_us = False
        self.app.Visible = True


@pytest.fixture
def mpp(tmp_path, monkeypatch):
    import evm.reader_com
    monkeypatch.setattr(evm.reader_com, "ComSession", FakeComSession)
    path = tmp_path / "p.mpp"
    path.write_bytes(b"mpp")
    cfg = tmp_path / "c.toml"
    cfg.write_text("[exportacion]\nnivel_maximo = 2\n", encoding="utf-8")
    return path, cfg


def test_command_does_not_write_without_confirming_and_writes_with_continuar(mpp, tmp_path):
    path, cfg = mpp
    args = ["corregir-linea-base", "--origen", path, "--config", cfg, "--tarea", "1.2", "--rama", "--de", "1", "--a",
            "10"]
    rc, text = _run(*args)
    assert rc == 2 and "LBA-008" in text                        # sin --motivo
    rc, text = _run(*args, "--motivo", "No comunicada", "--informe", tmp_path / "r1.md")
    s = FakeComSession.last
    assert rc == 1 and "No se escribió nada" in text and "LBA-014" in text, text
    assert s.com[7].Baseline1Start == FEB[0] and not s.saved and not list(tmp_path.glob("p.respaldo-*.mpp"))
    rc, text = _run(*args, "--motivo", "No comunicada", "--continuar", "--informe", tmp_path / "r2.md")
    s = FakeComSession.last
    assert rc == 0 and "archivo guardado (2 tareas)" in text, text
    assert s.com[7].Baseline1Start == "NA" and s.com[8].Baseline10Start == FEB[0] and s.saved
    assert len(list(tmp_path.glob("p.respaldo-*.mpp"))) == 1
    body = (tmp_path / "r2.md").read_text(encoding="utf-8")
    assert "Resultado: escrito y guardado" in body and "Motivo: No comunicada" in body


def test_command_interactive_keeps_the_file_open(mpp, tmp_path, monkeypatch):
    path, cfg = mpp
    answers = iter(["", "Errata", "s", "s"])     # Intro: la aprobada es la del PV (1); motivo; aplicar; correcto
    monkeypatch.setattr(bf, "console_available", lambda: True)
    monkeypatch.setattr("builtins.input", lambda q="": next(answers))
    rc, text = _run("corregir-linea-base", "--origen", path, "--config", cfg, "--tarea", "1.1.2.1", "--de", "1",
                    "--a", "10", "--interactivo", "--informe", tmp_path / "r.md")
    s = FakeComSession.last
    assert rc == 0, text
    assert s.com[5].Baseline1Start == "NA" and not s.saved and not s.closed and s.app.Visible
    assert "Interactivo: 1 tarea cambiada en MS Project, sin guardar" in text
    assert "Decisión: aplicar" in (tmp_path / "r.md").read_text(encoding="utf-8")


@pytest.mark.parametrize("started", [True, False])
def test_hand_over_restores_alerts_and_leaves_ms_project_open(started):
    from types import SimpleNamespace

    from evm.reader_com import MSO_AUTOMATION_SECURITY_BY_UI, ComSession
    s = ComSession("p.mpp")
    s.app = SimpleNamespace(DisplayAlerts=False, AutomationSecurity=3, Visible=False, UserControl=False)
    s.started_by_us, s.opened_by_us = started, True
    s._saved_app_props = {} if started else {"DisplayAlerts": True, "AutomationSecurity": 1}
    s.hand_over()
    assert s.app.DisplayAlerts is True and s.app.Visible and s.app.UserControl
    assert s.app.AutomationSecurity == (MSO_AUTOMATION_SECURITY_BY_UI if started else 1)
    assert not s.opened_by_us and not s.started_by_us     # al terminar no se cierra el archivo ni MS Project
