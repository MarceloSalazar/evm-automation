"""Campo Visible: fila propia (Sí), sin fila y sumada en la fila superior (No) o según el nivel máximo (vacío)."""
from datetime import date

import pytest
from openpyxl import load_workbook

from evm.config import Config, load_config
from evm.fields import parse_visible
from evm.migrate import plan_migration, prior_from_book
from evm.reader_xml import read_xml
from evm.update import read_book, run_update
from evm.wbs import KIND_LEAF, KIND_SUMMARY, SEV_INFO, SEV_WARN, build_report
from evm.workbook import read_structure, write_workbook

from builders import P, by_name, issues_with, rows_by_code, tree
from mspdi import project_xml, task_xml, write_xml


EDT = '''
    1 Gestión
      1.1 Dirección
        1.1.1 PMO
        1.1.2 Control
      1.2 Contratos
        1.2.1 Licitación
          1.2.1.1 Pliegos
          1.2.1.2 Adjudicación
        1.2.2 Obras
    2 Ingeniería
      2.1 Básica
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
    return build_report(P(task_list), cfg)


def codes(rep) -> list[str]:
    return [r.code for r in rep.rows]


def leaves(row) -> list[str]:
    """Tareas cuyo avance forma la fila."""
    return [t.name for t in row.leaves]


def all_leaves(rep) -> set[str]:
    return {t.name for r in rep.rows if r.is_input for t in r.leaves}


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
    assert any("campo Visible 'Mostrar'" in w.message for w in read_xml(plain, explicit).warnings)
    assert [t.visible for t in read_xml(path, Config(visible_field="")).tasks] == ["", ""]


# ----------------------------------------------------------------------------
# Filas del libro
# ----------------------------------------------------------------------------
def test_without_values_rows_follow_the_maximum_level():
    rep = report(tasks())
    assert codes(rep) == ["1.0.0", "1.1.0", "1.1.1", "1.1.2", "1.2.0", "1.2.1", "1.2.2", "2.0.0", "2.1.0"]
    assert leaves(rows_by_code(rep)["1.2.1"]) == ["Pliegos", "Adjudicación"]


def test_when_all_subtasks_are_hidden_the_parent_is_the_work_package():
    rep = report(tasks(PMO="No", Control="no"))
    rows = rows_by_code(rep)
    assert "1.1.1" not in rows and "1.1.2" not in rows
    assert rows["1.1.0"].kind == KIND_LEAF and not [r for r in rows if r.endswith("-D")]
    assert leaves(rows["1.1.0"]) == ["PMO", "Control"]          # la fila es el paquete, con toda su rama
    assert issues_with(rep, "2 tareas tienen Visible = No (1.1.1, 1.1.2): no tienen fila propia y cuentan en su fila "
                            "superior", SEV_INFO)


def test_no_counts_in_the_row_above():
    """Visible = No (desde la 2.7.0): sin fila propia, pero cuenta en su fila superior (no queda fuera del cálculo)."""
    rep = report(tasks(Obras="NO"))
    rows = rows_by_code(rep)
    assert "1.2.2" not in rows and rows["1.2.0"].kind == KIND_SUMMARY
    assert leaves(rows["1.2.0-D"]) == ["Obras"]                  # tarea suelta bajo la fila resumen: fila -D
    rep = report(tasks(Obras="No", Pliegos="No"))
    rows = rows_by_code(rep)
    assert leaves(rows["1.2.1"]) == ["Pliegos", "Adjudicación"]  # dentro del paquete, también cuenta
    assert "Obras" in all_leaves(rep)


def test_yes_decides_the_rows_like_a_wbs_list():
    """Con alguna tarea con Sí, solo las tareas con Sí tienen fila; las demás se suman en su fila superior con Sí."""
    rep = report(tasks(Gestión="Sí", Dirección="si", Licitación="Sí", Pliegos="Sí", Ingeniería="Sí"))
    rows = rows_by_code(rep)
    assert codes(rep) == ["1.0.0", "1.1.0", "1.2.1", "1.2.1.1", "1.2.1-D", "1.0.0-D", "2.0.0"]
    assert rows["1.1.0"].kind == KIND_LEAF and leaves(rows["1.1.0"]) == ["PMO", "Control"]
    assert leaves(rows["1.2.1.1"]) == ["Pliegos"] and leaves(rows["1.2.1-D"]) == ["Adjudicación"]
    assert leaves(rows["1.0.0-D"]) == ["Obras"]                  # Contratos (sin Sí): su tarea suelta, en Gestión
    assert rows["2.0.0"].kind == KIND_LEAF and leaves(rows["2.0.0"]) == ["Básica"]
    assert issues_with(rep, "Filas del libro según el campo Visible: 5 tareas con Sí tienen fila propia", SEV_INFO)


def test_yes_below_a_no_has_its_row():
    rep = report(tasks(Gestión="Sí", Contratos="No", Pliegos="Sí", Ingeniería="Sí"))
    rows = rows_by_code(rep)
    assert leaves(rows["1.2.1.1"]) == ["Pliegos"]
    assert {"PMO", "Control", "Adjudicación", "Obras"} <= set(leaves(rows["1.0.0-D"]))
    assert not any(i.id in ("VIS-002", "VIS-003") for i in rep.issues)


def test_other_problems():
    task_list = tasks(Básica="Quizás")
    task_list.append(tree("- Suelta 0")[0])
    task_list[-1].uid = task_list[-1].id = 99
    task_list[-1].outline_level = 3
    task_list[-1].visible = "Sí"
    rep = report(task_list)
    assert issues_with(rep, "Visible = 'Quizás' no se reconoce", SEV_WARN)
    assert issues_with(rep, "Visible = Sí en una tarea sin código EDT", SEV_WARN)


def test_field_disabled_in_the_configuration_ignores_the_values():
    rep = report(tasks(PMO="No"), visible_field="")
    assert "1.1.1" in rows_by_code(rep)


# ----------------------------------------------------------------------------
# actualizar y --migrar-desde
# ----------------------------------------------------------------------------
def bac(path, values: dict[str, float]) -> dict[str, int]:
    """Ingresa el BAC de las filas indicadas en la hoja PV y devuelve las filas del libro por código."""
    wb = load_workbook(path)
    rows = {s.code: s.row for s in read_structure(wb)}
    for code, value in values.items():
        wb["PV"].cell(row=rows[code], column=4, value=value)
    wb.save(path)
    return rows


def test_update_counts_hidden_tasks_and_warns_when_the_field_changes(tmp_path):
    path = tmp_path / "libro.xlsx"
    cfg = Config()
    write_workbook(build_report(P(tasks(Pliegos="No")), cfg), cfg, path)
    rows = bac(path, {"1.2.1": 1000})
    meta = {k: v for k, v in load_workbook(path)["_meta"].iter_rows(min_row=2, values_only=True)}
    assert meta["campo_visible"] == "Visible"
    progress = tasks(Pliegos="No")
    by_name(progress)["Pliegos"].percent_complete = 100                    # Visible = No: cuenta en su fila
    by_name(progress)["Adjudicación"].percent_complete = 50
    res = run_update(P(progress), Config(), path, date(2026, 3, 31), backup=False)
    book = read_book(path)
    assert sum(v or 0 for v in book.ev[rows["1.2.1"]]) == pytest.approx(750)
    assert not [w for w in res.warnings if "campo Visible" in w.message]

    changed = tasks(Pliegos="No", PMO="No")                                 # PMO pierde su fila
    res = run_update(P(changed), Config(), path, date(2026, 3, 31), backup=False, dry_run=True)
    assert [w for w in res.warnings if "campo Visible de 1 tarea (1.1.1)" in w.message and "--migrar-desde" in w.message]


def test_migration_sums_the_history_of_rows_that_are_now_hidden(tmp_path):
    old, new = tmp_path / "old.xlsx", tmp_path / "new.xlsx"
    cfg = Config(months=12)
    write_workbook(build_report(P(tasks()), cfg), cfg, old)
    bac(old, {"1.1.1": 100, "1.1.2": 200})
    progress = tasks()
    for name in ("PMO", "Control"):
        by_name(progress)[name].percent_complete = 50
    run_update(P(progress), cfg, old, date(2026, 2, 28), backup=False)
    rep = build_report(P(tasks(PMO="No", Control="No")), cfg)
    prior = prior_from_book(old, "Programa")
    rep.issues.extend(plan_migration(rep, prior, cfg.max_level))
    write_workbook(rep, cfg, new, prior)
    a, b = read_book(old), read_book(new)
    ra = {s.code: s.row for s in a.structure}
    rb = {s.code: s.row for s in b.structure}
    assert sum(v or 0 for v in b.ev[rb["1.1.0"]]) == pytest.approx(
        sum(v or 0 for r in ("1.1.1", "1.1.2") for v in a.ev[ra[r]])) == pytest.approx(150)
    assert b.bac[rb["1.1.0"]] == 300
    assert issues_with(rep, "Se sumó el historial")
