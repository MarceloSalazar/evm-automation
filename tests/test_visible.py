"""Campo Visible: fila propia (Sí), sin fila y sumada en la fila superior (No) o según el nivel máximo (vacío)."""
from datetime import date

import pytest
from openpyxl import load_workbook

from evm.config import Config, load_config
from evm.fields import parse_visible
from evm.migrate import plan_migration, prior_from_book
from evm.reader_xml import read_xml
from evm.update import read_book, run_update
from evm.wbs import KIND_LEAF, KIND_RESIDUAL, KIND_SUMMARY, SEV_INFO, SEV_WARN, LinearPVSource, build_report
from evm.workbook import read_structure, write_workbook

from builders import P, by_name, issues_with, rows_by_code, tree
from mspdi import project_xml, task_xml, write_xml

COST = pytest.mark.modo("costo")

EDT = '''
    1 Gestión 0
      1.1 Dirección 0
        1.1.1 PMO 100
        1.1.2 Control 200
      1.2 Contratos 0
        1.2.1 Licitación 0
          1.2.1.1 Pliegos 120
          1.2.1.2 Adjudicación 180
        1.2.2 Obras 400
    2 Ingeniería 0
      2.1 Básica 500
'''


def tasks(**visible) -> list:
    """EDT de prueba; visible: {nombre de la tarea: valor del campo Visible}."""
    out = tree(EDT)
    names = by_name(out)
    for name, value in visible.items():
        names[name].visible = value
    return out


def report(task_list, **kw):
    cfg = Config(**kw)
    cfg.validate()
    return build_report(P(task_list), cfg, LinearPVSource())


def codes(rep) -> list[str]:
    return [r.code for r in rep.rows]


# ----------------------------------------------------------------------------
# Valores y configuración
# ----------------------------------------------------------------------------
def test_values_are_recognized_in_any_case_and_with_or_without_accent():
    for text in ("Sí", "si", "SI", "sí", "S", "x", "1", "Yes", True, 1.0):
        assert parse_visible(text) == (True, True), text
    for text in ("No", "NO", "no", "n", "0", "false", False, 0):
        assert parse_visible(text) == (False, True), text
    for text in ("", "  ", None):
        assert parse_visible(text) == (None, True)
    assert parse_visible("Quizás") == (None, False)


def test_config_key(tmp_path):
    assert Config().visible_field == "Visible" and not Config().visible_field_explicit
    path = tmp_path / "c.toml"
    path.write_text('[campos]\nvisible = "Mostrar"\n', encoding="utf-8")
    cfg = load_config(path)
    assert cfg.visible_field == "Mostrar" and cfg.visible_field_explicit
    path.write_text('[campos]\nvisible = ""\n', encoding="utf-8")
    assert load_config(path).visible_field == ""


def test_xml_reader_reads_the_field_by_its_alias(tmp_path):
    text5 = 188743743
    xml = project_xml([task_xml(1, 1, "Gestión", wbs="1", ext={text5: "No"}), task_xml(2, 1, "Obras", wbs="2")],
                      definitions=[(text5, "Text5", "Visible")])
    path = write_xml(tmp_path / "p.xml", xml)
    project = read_xml(path, Config())
    assert [t.visible for t in project.tasks] == ["No", ""] and not project.warnings
    plain = write_xml(tmp_path / "sin_campo.xml", project_xml([task_xml(1, 1, "Gestión", wbs="1")]))
    assert not read_xml(plain, Config()).warnings                    # campo por defecto: no se avisa si no existe
    explicit = Config(visible_field="Mostrar", visible_field_explicit=True)
    assert any("campo Visible 'Mostrar'" in w for w in read_xml(plain, explicit).warnings)
    assert [t.visible for t in read_xml(path, Config(visible_field="")).tasks] == ["", ""]


# ----------------------------------------------------------------------------
# Filas del libro
# ----------------------------------------------------------------------------
@COST
def test_without_values_rows_follow_the_maximum_level():
    rep = report(tasks())
    assert codes(rep) == ["1.0.0", "1.1.0", "1.1.1", "1.1.2", "1.2.0", "1.2.1", "1.2.2", "2.0.0", "2.1.0"]
    assert rows_by_code(rep)["1.2.1"].bac == 300


@COST
def test_hidden_tasks_are_summed_in_their_parent_row():
    rep = report(tasks(PMO="No", Control="no"))
    rows = rows_by_code(rep)
    assert "1.1.1" not in rows and "1.1.2" not in rows
    assert rows["1.1.0"].kind == KIND_LEAF and rows["1.1.0"].bac == 300
    assert rep.total_bac == 1500
    assert issues_with(rep, "2 tareas tienen Visible = No", SEV_INFO)


@COST
def test_partially_hidden_children_form_the_other_costs_row():
    rep = report(tasks(Obras="NO"))
    rows = rows_by_code(rep)
    assert "1.2.2" not in rows and rows["1.2.0"].kind == KIND_SUMMARY
    assert rows["1.2.0-D"].kind == KIND_RESIDUAL and rows["1.2.0-D"].bac == 400
    assert rep.total_bac == 1500 and issues_with(rep, "u ocultas (Visible = No)")


def test_percent_mode_hidden_task_goes_to_the_other_tasks_row():
    rep = report(tasks(Obras="No"))
    rows = rows_by_code(rep)
    assert rows["1.2.0-D"].name.startswith("Otras tareas") and [t.name for t in rows["1.2.0-D"].leaves] == ["Obras"]


@COST
def test_yes_shows_tasks_below_the_maximum_level():
    rep = report(tasks(Pliegos="Sí"))
    rows = rows_by_code(rep)
    assert rows["1.2.1"].kind == KIND_SUMMARY
    assert rows["1.2.1.1"].kind == KIND_LEAF and rows["1.2.1.1"].bac == 120 and rows["1.2.1.1"].level == 4
    assert rows["1.2.1-D"].bac == 180                     # la tarea hermana sin Sí: fila de otros costos
    assert rep.total_bac == 1500
    assert issues_with(rep, "1 tarea tiene Visible = Sí y fila propia por debajo del nivel máximo (3)", SEV_INFO)
    both = rows_by_code(report(tasks(Pliegos="Sí", Adjudicación="si")))
    assert both["1.2.1.2"].bac == 180 and "1.2.1-D" not in both


@COST
def test_no_prevails_over_yes_in_its_branch():
    rep = report(tasks(Contratos="No", Pliegos="Sí"))
    rows = rows_by_code(rep)
    assert not [c for c in rows if c.startswith("1.2")]
    assert rows["1.0.0-D"].bac == 700 and rep.total_bac == 1500
    assert issues_with(rep, "depende de 1.2.0 «Contratos», con Visible = No", SEV_WARN)


@COST
def test_no_without_a_parent_row_is_ignored_and_other_problems_are_warned():
    task_list = tasks(Gestión="No", Básica="Quizás")
    task_list.append(tree("- Suelta 0")[0])
    task_list[-1].uid = task_list[-1].id = 99
    task_list[-1].outline_level = 3
    task_list[-1].visible = "Sí"
    rep = report(task_list)
    assert "1.0.0" in rows_by_code(rep) and "2.1.0" in rows_by_code(rep)
    assert issues_with(rep, "sin fila superior donde sumarlo: se ignora", SEV_WARN)
    assert issues_with(rep, "Visible = 'Quizás' no se reconoce", SEV_WARN)
    assert issues_with(rep, "Visible = Sí en una tarea sin código EDT", SEV_WARN)


@COST
def test_field_disabled_in_the_configuration_ignores_the_values():
    rep = report(tasks(PMO="No"), visible_field="")
    assert "1.1.1" in rows_by_code(rep)


# ----------------------------------------------------------------------------
# actualizar y --migrar-desde
# ----------------------------------------------------------------------------
def test_update_counts_hidden_tasks_and_warns_when_the_field_changes(tmp_path):
    path = tmp_path / "libro.xlsx"
    cfg = Config()
    task_list = tasks(Obras="No")
    write_workbook(build_report(P(task_list), cfg, LinearPVSource()), cfg, path)
    wb = load_workbook(path)
    rows = {s.code: s.row for s in read_structure(wb)}
    wb["PV"].cell(row=rows["1.2.0-D"], column=4, value=1000)
    wb.save(path)
    meta = {k: v for k, v in wb["_meta"].iter_rows(min_row=2, values_only=True)}
    assert meta["campo_visible"] == "Visible"
    progress = tasks(Obras="No")
    by_name(progress)["Obras"].percent_complete = 100
    res = run_update(P(progress), Config(), path, date(2026, 3, 31), backup=False)
    book = read_book(path)
    assert sum(v or 0 for v in book.ev[rows["1.2.0-D"]]) == pytest.approx(1000)
    assert not [w for w in res.warnings if "campo Visible" in w]

    changed = tasks(Obras="Sí")
    res = run_update(P(changed), Config(), path, date(2026, 3, 31), backup=False, dry_run=True)
    assert [w for w in res.warnings if "campo Visible de 1 tarea (1.2.2)" in w and "--migrar-desde" in w]


@COST
def test_migration_sums_the_history_of_rows_that_are_now_hidden(tmp_path):
    old, new = tmp_path / "old.xlsx", tmp_path / "new.xlsx"
    cfg = Config(months=12)
    write_workbook(build_report(P(tasks()), cfg, LinearPVSource()), cfg, old)
    progress = tasks()
    for name in ("PMO", "Control"):
        by_name(progress)[name].percent_complete = 50
    run_update(P(progress), cfg, old, date(2026, 2, 28), backup=False)
    rep = build_report(P(tasks(PMO="No", Control="No")), cfg, LinearPVSource())
    prior = prior_from_book(old, "Programa")
    rep.issues.extend(plan_migration(rep, prior, cfg.max_level))
    write_workbook(rep, cfg, new, prior)
    a, b = read_book(old), read_book(new)
    ra = {s.code: s.row for s in a.structure}
    rb = {s.code: s.row for s in b.structure}
    assert sum(v or 0 for v in b.ev[rb["1.1.0"]]) == pytest.approx(
        sum(v or 0 for r in ("1.1.1", "1.1.2") for v in a.ev[ra[r]]))
    assert issues_with(rep, "Se sumó el historial")
