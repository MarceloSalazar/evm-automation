"""actualizar-edt: numeración de las subtareas de una tarea según la jerarquía, sin cambiar las protegidas (nivel 1 a
nivel_maximo, Visible = Sí y lista EDT), con las protegidas primero y sin saltos; escritura con un MS Project
simulado. Las pruebas con MS Project real están en test_msproject_edt.py."""
from __future__ import annotations

import contextlib
import io

import pytest

from evm.cli import main
from evm.config import Config, ConfigError
from evm.edt_renumber import check_field, plan_renumber, write_codes
from evm.fields import TASK_TEXT_IDS
from evm.messages import SEV_ERROR, SEV_WARN, EvmError

from builders import P, T
from mspdi import project_xml, task_xml, write_xml

TEXT5 = TASK_TEXT_IDS[5]


def tasks_of(text: str, **marks) -> list:
    """Tareas de un esquema con sangría de 2 espacios: '<código> <nombre>' ('-' = sin código). marks: {nombre:
    {campo: valor}} (p. ej. visible='Sí')."""
    out = []
    for line in text.strip("\n").splitlines():
        if not line.strip():
            continue
        code, name = line.split(maxsplit=1)
        level = (len(line) - len(line.lstrip(" "))) // 2 + 1
        out.append(T(len(out) + 1, level, "" if code == "-" else code, name, **marks.get(name, {})))
    return out


def renumber(text: str, *starts: str, cfg: Config | None = None, **marks):
    cfg = cfg or Config()
    cfg.validate()
    tasks = tasks_of(text, **marks)
    plan = plan_renumber(P(tasks, fields_found={"edt": "Text5"}), cfg, list(starts or ["2.7.1"]))
    return plan, {c.task.name: c.new for c in plan.changes}


def ids(plan) -> list[str]:
    return [i.id for i in plan.issues]


SI = {"visible": "Sí"}

BRANCH = """
2 Fase
  2.7 Grupo
    2.7.1 Paquete
      2.7.1.3 C
      - Nueva
        - Nueva hija
      2.7.1.1 A
        2.7.1.1.4 A1
    2.7.2 Otro paquete
      2.7.2.1 Fuera
- Hito contractual
"""


def test_children_are_numbered_by_outline_order_and_the_start_keeps_its_code():
    plan, new = renumber(BRANCH)
    assert new == {"C": "2.7.1.1", "Nueva": "2.7.1.2", "Nueva hija": "2.7.1.2.1", "A": "2.7.1.3", "A1": "2.7.1.3.1"}
    assert plan.tasks == 5 and plan.added == 2 and not plan.issues
    # Fuera de la rama no cambia nada: ni el otro paquete ni el hito contractual sin código
    assert "Fuera" not in new and "Hito contractual" not in new and "Paquete" not in new


def test_codes_already_in_order_give_no_changes():
    plan, new = renumber("""
2.7.1 Paquete
  2.7.1.1 A
    2.7.1.1.1 A1
  2.7.1.2 B
""")
    assert new == {} and not plan.issues


def test_visible_tasks_go_first_and_the_rest_follow_them():
    plan, new = renumber("""
2.7.1 Paquete
  2.7.1.1 V1
    2.7.1.1.7 V1 hija
  2.7.1.2 V2
  2.7.1.9 X
  - Y
""", V1=SI, V2=SI)
    assert new == {"V1 hija": "2.7.1.1.1", "X": "2.7.1.3", "Y": "2.7.1.4"}
    assert {k.task.name: k.reason for k in plan.kept} == {"V1": "Visible = Sí", "V2": "Visible = Sí"}
    assert not plan.issues


def test_a_visible_task_after_one_that_is_renumbered_is_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  2.7.1.2 X
  2.7.1.1 V
""", V=SI)
    assert ids(plan) == ["NUM-008"] and plan.errors and plan.issues[0].severity == SEV_ERROR
    assert "2.7.1.2 «X»" in plan.issues[0].message


def test_a_gap_in_the_protected_codes_is_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  2.7.1.1 V1
  2.7.1.3 V2
  - X
""", V1=SI, V2=SI)
    assert ids(plan) == ["NUM-009"] and "2.7.1.1, 2.7.1.3" in plan.issues[0].message


def test_protected_codes_out_of_order_are_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  2.7.1.2 V2
  2.7.1.1 V1
""", V1=SI, V2=SI)
    assert ids(plan) == ["NUM-009"]


def test_a_task_containing_a_visible_task_keeps_its_code():
    plan, new = renumber("""
2.7.1 Paquete
  2.7.1.1 Contenedor
    2.7.1.1.1 V
    2.7.1.1.5 Z
  2.7.1.4 X
""", V=SI)
    assert new == {"Z": "2.7.1.1.2", "X": "2.7.1.2"}
    assert {k.task.name: k.reason for k in plan.kept}["Contenedor"] == "contiene tareas protegidas"


def test_codes_of_the_wbs_list_are_protected():
    cfg = Config(wbs_list=["2.7.1", "2.7.1.1"])
    plan, new = renumber("""
2.7.1 Paquete
  2.7.1.1 L
    2.7.1.1.3 L1
  2.7.1.5 X
""", cfg=cfg)
    assert new == {"L1": "2.7.1.1.1", "X": "2.7.1.2"}
    assert [k.reason for k in plan.kept] == ["lista EDT"]


def test_from_a_level_above_the_maximum_the_book_levels_are_kept():
    plan, new = renumber("""
2.7 Grupo
  2.7.1 Paquete
    2.7.1.4 A
  2.7.2 Otro paquete
  - Actividad
    - Paso
""", "2.7")
    # Los paquetes (nivel 3) no cambian; la actividad sin paquete lleva los ceros hasta nivel_maximo
    assert new == {"A": "2.7.1.1", "Actividad": "2.7.0.1", "Paso": "2.7.0.1.1"}
    assert [k.reason for k in plan.kept] == ["nivel 3", "nivel 3"] and not plan.issues


def test_with_ignored_interior_zeros_a_new_book_level_code_is_not_assigned():
    plan, new = renumber("""
2.7 Grupo
  2.7.1 Paquete
  - Actividad
    - Paso
""", "2.7", cfg=Config(ignore_interior_zeros=True))
    assert new == {} and ids(plan) == ["NUM-013"] and plan.issues[0].severity == SEV_WARN
    assert [t.name for t in plan.skipped] == ["Actividad"]


def test_a_level_code_that_does_not_hang_from_its_parent_is_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  3.1 Copiada
  - X
""")
    assert ids(plan) == ["NUM-010"]


def test_a_task_without_code_containing_protected_tasks_is_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  - Grupo
    2.7.1.1 V
""", V=SI)
    assert "NUM-011" in ids(plan)


def test_visible_without_code_is_an_error():
    plan, _new = renumber("""
2.7.1 Paquete
  2.7.1.1 A
  - Nueva visible
""", **{"Nueva visible": SI})
    assert ids(plan) == ["NUM-012"]


def test_a_new_code_that_another_task_has_is_reported_and_still_changed():
    plan, new = renumber("""
2.7.1 Paquete
  - Nueva
2.7.2 Otro
  2.7.1.1 Mal ubicado
""")
    assert new == {"Nueva": "2.7.1.1"}
    [issue] = plan.issues
    assert issue.id == "NUM-014" and issue.severity == SEV_WARN and "Mal ubicado" in issue.message
    assert not plan.errors


def test_inactive_tasks_are_numbered_in_their_place():
    tasks_text = """
2.7.1 Paquete
  2.7.1.5 A
  2.7.1.6 Inactiva
  2.7.1.7 B
"""
    _plan, new = renumber(tasks_text, Inactiva={"active": False})
    assert new == {"A": "2.7.1.1", "Inactiva": "2.7.1.2", "B": "2.7.1.3"}


def test_the_plan_prefix_is_kept():
    cfg = Config(plan="INTERNO")
    _plan, new = renumber("""
INTERNO Interno
  INTERNO.2.7.1 Paquete
    - Nueva
""", "2.7.1", cfg=cfg)
    assert new == {"Nueva": "INTERNO.2.7.1.1"}


@pytest.mark.parametrize("start, code", [("9.9.9", "NUM-001"), ("INTERNO", "NUM-003")])
def test_start_not_found_or_not_numeric(start, code):
    with pytest.raises(ConfigError) as exc:
        renumber(BRANCH, start)
    assert exc.value.id == code


def test_duplicated_start_and_overlapping_starts():
    with pytest.raises(ConfigError) as exc:
        renumber("""
2.7.1 A
2.7.1 B
""")
    assert exc.value.id == "NUM-002"
    with pytest.raises(ConfigError) as exc:
        renumber(BRANCH, "2.7", "2.7.1")
    assert exc.value.id == "NUM-004"


@pytest.mark.parametrize("found", [{}, {"edt": "WBS"}, {"edt": "OutlineNumber"}])
def test_only_a_custom_text_field_can_be_written(found):
    with pytest.raises(ConfigError) as exc:
        check_field(P([], fields_found=found), Config())
    assert exc.value.id == "NUM-005"


# ----------------------------------------------------------------------------
# Escritura con un MS Project simulado
# ----------------------------------------------------------------------------
class FakeTask:
    def __init__(self, uid, code, broken=False):
        self.UniqueID, self.ExternalTask, self.broken = uid, False, broken
        self.__dict__["Text5"] = code

    def __setattr__(self, name, value):
        if name == "Text5" and self.__dict__.get("broken"):
            return                          # MS Project que no guarda el valor
        object.__setattr__(self, name, value)


class FakeApp:
    def __init__(self, undo=True, formula=""):
        self.undo, self.formula, self.ScreenUpdating, self.log = undo, formula, True, []

    def OpenUndoTransaction(self, name):
        if not self.undo:
            raise RuntimeError("no disponible")
        self.log.append(("abrir", name))

    def CloseUndoTransaction(self):
        self.log.append(("cerrar",))

    def CustomFieldGetFormula(self, fid):
        return self.formula


class FakeProject:
    def __init__(self, tasks):
        self.Tasks = tasks


class FakeSession:
    def __init__(self, com_tasks, **app):
        self.app, self.project = FakeApp(**app), FakeProject(com_tasks)

    def _resolve_field(self, name):
        return ("attr", "Text5")

    def _activate(self):
        pass


def _plan():
    plan, _new = renumber("""
2.7.1 Paquete
  2.7.1.3 C
  - Nueva
""")
    return plan


def test_codes_are_written_in_one_undo_action_and_checked():
    com = [FakeTask(1, "2.7.1"), FakeTask(2, "2.7.1.3"), FakeTask(3, "")]
    session = FakeSession(com)
    out = write_codes(session, _plan(), Config())
    assert [t.Text5 for t in com] == ["2.7.1", "2.7.1.1", "2.7.1.2"]
    assert out.written == 2 and out.undo_grouped and not out.issues
    assert session.app.log == [("abrir", "Actualizar códigos EDT"), ("cerrar",)]
    assert session.app.ScreenUpdating is True


def test_without_undo_transactions_there_is_a_warning():
    com = [FakeTask(1, "2.7.1"), FakeTask(2, "2.7.1.3"), FakeTask(3, "")]
    out = write_codes(FakeSession(com, undo=False), _plan(), Config())
    assert [i.id for i in out.issues] == ["NUM-017"] and out.written == 2


def test_a_code_that_is_not_kept_by_ms_project_stops_with_an_error():
    com = [FakeTask(1, "2.7.1"), FakeTask(2, "2.7.1.3", broken=True), FakeTask(3, "")]
    with pytest.raises(EvmError) as exc:
        write_codes(FakeSession(com), _plan(), Config())
    assert exc.value.id == "NUM-016"


def test_a_field_with_formula_cannot_be_written():
    com = [FakeTask(1, "2.7.1"), FakeTask(2, "2.7.1.3"), FakeTask(3, "")]
    with pytest.raises(ConfigError) as exc:
        write_codes(FakeSession(com, formula="[Text1]"), _plan(), Config())
    assert exc.value.id == "NUM-005"


# ----------------------------------------------------------------------------
# Comando
# ----------------------------------------------------------------------------
def _xml(tmp_path):
    tasks = [task_xml(0, 0, "Proyecto"),
             task_xml(1, 1, "Paquete", ext={TEXT5: "2.7.1"}),
             task_xml(2, 2, "C", ext={TEXT5: "2.7.1.3"}),
             task_xml(3, 2, "Nueva")]
    path = write_xml(tmp_path / "p.xml", project_xml(tasks, definitions=[(TEXT5, "Text5", "Item EDT")]))
    cfg = tmp_path / "c.toml"
    cfg.write_text('[campos]\nedt = "Item EDT"\n', encoding="utf-8")
    return path, cfg


def _run(*args):
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        rc = main([str(a) for a in args])
    return rc, out.getvalue()


def test_command_simulates_with_xml_and_writes_the_report(tmp_path):
    xml, cfg = _xml(tmp_path)
    report = tmp_path / "cambios.md"
    rc, text = _run("actualizar-edt", "--origen", xml, "--config", cfg, "--desde", "2.7.1", "--simular",
                    "--informe", report)
    assert rc == 0, text
    assert "2.7.1.3 -> 2.7.1.1" in " ".join(text.split()) and "Simulación" in text
    body = report.read_text(encoding="utf-8")
    assert "| 3 | (sin código) | 2.7.1.2 | Nueva |  |" in body and "Simulación" in body


def test_command_needs_ms_project_to_write(tmp_path):
    xml, cfg = _xml(tmp_path)
    rc, text = _run("actualizar-edt", "--origen", xml, "--config", cfg, "--desde", "2.7.1")
    assert rc == 2 and "NUM-007" in text


def test_command_with_errors_ends_with_code_1(tmp_path):
    xml, cfg = _xml(tmp_path)
    xml.write_text(xml.read_text(encoding="utf-8").replace("<Value>2.7.1.3</Value>", "<Value>3.1</Value>"),
                   encoding="utf-8")
    rc, out = _run("actualizar-edt", "--origen", xml, "--config", cfg, "--desde", "2.7.1", "--simular",
                   "--informe", tmp_path / "r.md")
    assert rc == 1 and "NUM-010" in out and "No se escribió ningún código" in out
