"""Lista de códigos EDT en un archivo ([exportacion] lista_edt): decide las filas del libro en lugar del nivel
máximo y del campo Visible."""
from datetime import datetime

import pytest
from openpyxl import load_workbook

from evm.config import Config, ConfigError, load_config, read_wbs_list
from evm.update import read_book, run_update
from evm.validation import combine, validate_part
from evm.wbs import KIND_LEAF, KIND_RESIDUAL, KIND_SUMMARY, SEV_INFO, SEV_WARN, build_report
from evm.workbook import SHEET_WBS_LIST, read_meta, write_workbook

from builders import P, by_name, issues_with, report, rows_by_code, set_bac, tree

TASKS = '''
    1 Fase
      1.1 A
        1.1.1 A1
        1.1.2 A2
      1.2 B
    2 Fase 2
      2.1 C
        2.1.1 C1
'''


def listed(codes: list[str], **kw) -> Config:
    return Config(wbs_list=list(codes), wbs_list_file="edt.txt", **kw)


# ----------------------------------------------------------------------------
# Archivo
# ----------------------------------------------------------------------------
def test_read_list_ignores_comments_blank_lines_names_and_repeats(tmp_path):
    f = tmp_path / "edt.txt"
    f.write_text("# EDT del contrato\n\n1.0.0\n1.1.0\tIngeniería\n1.1.1 Estudios\n2.0;Obras\n1.1.0\n  2.1.0  \n",
                 encoding="utf-8")
    assert read_wbs_list(f) == ["1.0.0", "1.1.0", "1.1.1", "2.0", "2.1.0"]


@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-16", "cp1252"])
def test_read_list_encodings(tmp_path, encoding):
    f = tmp_path / "edt.txt"
    f.write_bytes("1.0.0 Diseño\n1.1.0 Construcción\n".encode(encoding))
    assert read_wbs_list(f) == ["1.0.0", "1.1.0"]


def test_read_list_errors(tmp_path):
    with pytest.raises(ConfigError, match="No se encontró el archivo de la lista EDT"):
        read_wbs_list(tmp_path / "falta.txt")
    empty = tmp_path / "vacio.txt"
    empty.write_text("# solo comentarios\n\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="no tiene códigos"):
        read_wbs_list(empty)


def test_config_key_resolves_relative_path(tmp_path):
    (tmp_path / "listas").mkdir()
    (tmp_path / "listas" / "edt.txt").write_text("1\n1.1\n", encoding="utf-8")
    cfg_file = tmp_path / "config.toml"
    cfg_file.write_text("[exportacion]\nlista_edt = 'listas/edt.txt'\n", encoding="utf-8")
    cfg = load_config(cfg_file)
    assert cfg.wbs_list == ["1", "1.1"] and cfg.wbs_list_file == str(tmp_path / "listas" / "edt.txt")
    assert load_config(None).wbs_list is None


# ----------------------------------------------------------------------------
# Filas del libro
# ----------------------------------------------------------------------------
def test_list_defines_rows_and_unlisted_tasks_go_to_nearest_listed_element():
    rep = report(tree(TASKS), listed(["1", "1.1", "1.1.1", "2"]))
    rows = rows_by_code(rep)
    assert list(rows) == ["1.0.0", "1.1.0", "1.1.1", "1.1.0-D", "1.0.0-D", "2.0.0"]
    assert rows["1.0.0"].kind == KIND_SUMMARY and rows["1.1.0"].kind == KIND_SUMMARY
    assert rows["1.1.1"].kind == KIND_LEAF
    assert [t.name for t in rows["1.1.0-D"].leaves] == ["A2"]               # 1.1.2 no está en la lista
    assert [t.name for t in rows["1.0.0-D"].leaves] == ["B"]                # 1.2 tampoco
    assert rows["2.0.0"].kind == KIND_LEAF and [t.name for t in rows["2.0.0"].leaves] == ["C1"]
    assert issues_with(rep, "que no están en la lista EDT: forman la fila 1.1.0-D")


def test_list_replaces_max_level_and_visible_field():
    tasks = tree(TASKS + "          2.1.1.1 Detalle\n")
    by_name(tasks)["B"].visible = "No"
    rep = report(tasks, listed(["1", "2", "2.1", "2.1.1", "2.1.1.1"], max_level=1))
    rows = rows_by_code(rep)
    assert list(rows) == ["1", "2", "2.1", "2.1.1", "2.1.1.1"]       # más detallado que nivel_maximo
    assert rows["2.1.1"].kind == KIND_SUMMARY and rows["2.1.1.1"].kind == KIND_LEAF
    assert "B" in [t.name for t in rows["1"].leaves]                 # Visible = No no excluye con la lista
    assert issues_with(rep, "el campo Visible no se usa", SEV_INFO)


def test_listed_codes_are_padded_like_the_book_and_planificacion_prefix_is_removed():
    tasks = tree('''
        INTERNO Plan
          INTERNO.1 Fase
            INTERNO.1.1 A
    ''')
    rep = report(tasks, listed(["INTERNO.1", "1.1.0"], plan="INTERNO"))
    assert list(rows_by_code(rep)) == ["1.0.0", "1.1.0"]


@pytest.mark.parametrize("codes, level", [
    (["1.0.0", "1.2.0", "1.2.1"], 3),                   # con los ceros del nivel 3
    (["1.0.0.0", "1.2.0.0", "1.2.1.0"], 4),             # con los ceros del nivel 4
    (["1", "1.2", "1.2.1"], 3),                         # sin ceros finales
    (["1.0.0.0.0", "1.2.0.0", "1.2.1.0.0"], 3),         # con más ceros
])
def test_listed_codes_match_tasks_without_trailing_zeros(codes, level):
    """1.2, 1.2.0 y 1.2.0.0 son el mismo código: la línea de la lista coincide con su tarea; el libro muestra el
    código de MS Project (completado con los ceros del nivel)."""
    tasks = tree('''
        1.0.0.0 Fase
          1.2.0.0 Elemento
            1.2.1.0 Paquete
              1.2.1.1 Actividad
    ''')
    rep = report(tasks, listed(codes, max_level=level))
    assert list(rows_by_code(rep)) == ["1.0.0.0", "1.2.0.0", "1.2.1.0"]
    assert not issues_with(rep, "no existe")
    short = report(tree('''
        1 Fase
          1.2 Elemento
    '''), listed(["1.0.0.0", "1.2.0.0"], max_level=level))
    assert list(rows_by_code(short)) == (["1.0.0", "1.2.0"] if level == 3 else ["1.0.0.0", "1.2.0.0"])


def test_missing_codes_warn_and_no_listed_code_is_an_error():
    rep = report(tree(TASKS), listed(["1", "1.1", "7.7"]))
    assert issues_with(rep, "1 código de la lista EDT no existe en MS Project", SEV_WARN)
    assert any(i.code == "7.7" for i in rep.issues if i.severity == SEV_WARN)
    with pytest.raises(ConfigError, match="Ningún código de la lista EDT"):
        report(tree(TASKS), listed(["7.7", "8"]))


def test_branches_without_listed_element_are_left_out_and_child_without_listed_parent_is_reported():
    rep = report(tree(TASKS), listed(["1.1"]))
    assert list(rows_by_code(rep)) == ["1.1.0"]
    assert issues_with(rep, "su tarea superior 1.0.0 «Fase» no: la fila queda en el primer nivel del libro")
    omitted = issues_with(rep, "no están en la lista EDT ni dependen de un elemento de la lista")
    assert omitted and "2.0.0" in omitted[0].message and omitted[0].severity == SEV_WARN


def test_validar_reports_missing_codes():
    part = validate_part(P(tree(TASKS)), listed(["1", "9"]))
    _report, issues = combine([part], listed(["1", "9"]))
    assert any("lista EDT no existe" in i.message for i in issues)


# ----------------------------------------------------------------------------
# Libro y 'actualizar'
# ----------------------------------------------------------------------------
def test_book_keeps_its_list_and_actualizar_uses_it(tmp_path):
    tasks = tree(TASKS, finish=datetime(2026, 3, 27, 17))
    cfg = listed(["1", "1.1", "1.1.1", "2"])
    book = tmp_path / "libro.xlsx"
    write_workbook(build_report(P(tasks), cfg), cfg, book)
    set_bac(book)
    wb = load_workbook(book)
    assert [c.value for c in wb[SHEET_WBS_LIST]["A"]][1:] == ["1", "1.1", "1.1.1", "2"]
    assert read_meta(wb)["lista_edt"] == "edt.txt"
    assert read_book(book).wbs_list == ["1", "1.1", "1.1.1", "2"]
    for t in tasks:
        t.percent_complete = 50.0
    # Sin lista en la configuración: se usa la del libro (1.2 sigue sumándose en 1.0.0-D)
    res = run_update(P(tasks), Config(), book, datetime(2026, 2, 28).date(), backup=False, dry_run=True)
    codes = {r.code for r in res.rows}
    assert codes == {"1.1.1", "1.1.0-D", "1.0.0-D", "2.0.0"}
    assert not any("lista" in w.message for w in res.warnings)
    # Lista distinta en la configuración: aviso, y se mantiene la del libro
    res = run_update(P(tasks), listed(["1", "2"]), book, datetime(2026, 2, 28).date(), backup=False,
                     dry_run=True)
    assert any("La lista EDT de la configuración cambió" in w.message for w in res.warnings)
    assert {r.code for r in res.rows} == codes
    # Misma lista escrita de otra forma (con otra cantidad de ceros finales): sin aviso
    for same in (["1.0.0", "1.1.0", "1.1.1", "2.0.0"], ["1.0.0.0", "1.1.0.0", "1.1.1.0", "2"]):
        res = run_update(P(tasks), listed(same), book, datetime(2026, 2, 28).date(), backup=False, dry_run=True)
        assert not any("lista" in w.message for w in res.warnings)


def test_actualizar_warns_when_book_has_no_list(tmp_path):
    tasks = tree(TASKS)
    book = tmp_path / "libro.xlsx"
    write_workbook(build_report(P(tasks), Config()), Config(), book)
    res = run_update(P(tasks), listed(["1"]), book, datetime(2026, 2, 28).date(), backup=False, dry_run=True)
    assert any("el libro se generó sin lista EDT" in w.message for w in res.warnings)


def test_kinds_of_residual_rows_with_list():
    rep = report(tree(TASKS), listed(["1", "1.1", "1.1.1", "2"]))
    assert [r.kind for r in rep.rows if r.code.endswith("-D")] == [KIND_RESIDUAL, KIND_RESIDUAL]
