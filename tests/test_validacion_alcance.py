"""Alcance de las observaciones (hojas Validación y Validación detalle, consola de generar y validar): si se refieren a
una fila del libro (las de las hojas PV, EV y AC), a una tarea de detalle que se suma en una fila, a una tarea fuera
del libro o al proyecto en general. Las de detalle y fuera del libro se agrupan por código y fila del libro."""
from datetime import datetime

from openpyxl import load_workbook

from evm import cli
from evm.cli import main
from evm.config import Config
from evm.messages import (SEV_ERROR, SEV_INFO, SEV_WARN, TARGET_DETAIL, TARGET_GENERAL, TARGET_OUT, TARGET_ROW, Issue,
                          by_priority, grouped)
from evm.wbs import build_part, code_key, located, locate_issues
from evm.workbook import SHEET_VALID, VALID_HEADER_ROW, V_ROW, write_workbook

from builders import P, by_name, report, tree, validation_rows
from mspdi import project_xml, task_xml, write_xml

# 1.1.1 es un paquete de trabajo con actividades de nivel 4 y una tarea sin código; 1.0.0 tiene una tarea sin código
# (fila 1.0.0-D); 2.1 tiene Visible = No (su hermana 2.2 no) y "Fuera" no depende de ningún elemento EDT
TASKS = '''
    1 Fase
      1.1 Entregable
        1.1.1 Paquete
          1.1.1.1 Actividad a
          1.1.1.2 Actividad b
          - Sin código
        1.1.2 Paquete dos
      - Suelta
    2 Fase dos
      2.1 Oculta
      2.2 Visible
    - Fuera
'''
ESTIMATED = {"Actividad a", "Actividad b", "Sin código", "Paquete dos", "Suelta", "Fuera"}


def tasks():
    out = tree(TASKS)
    for t in out:
        t.estimated = t.name in ESTIMATED
    by_name(out)["Oculta"].visible = "No"
    return out


def where(issues, id: str) -> dict[str, tuple[str, str]]:
    """{tarea: (alcance, fila del libro)} de las observaciones del código `id`."""
    return {i.name: (i.target, i.row) for i in issues if i.id == id}


def test_each_issue_has_its_target_and_book_row():
    rep = report(tasks())
    issues = located(rep.rows, rep.issues)
    assert where(issues, "FEC-007") == {
        "Actividad a": (TARGET_DETAIL, "1.1.1"), "Actividad b": (TARGET_DETAIL, "1.1.1"),
        "Sin código (Id 6)": (TARGET_DETAIL, "1.1.1"), "Paquete dos": (TARGET_ROW, "1.1.2"),
        "Suelta (Id 8)": (TARGET_DETAIL, "1.0.0-D"), "Fuera (Id 12)": (TARGET_OUT, "")}
    assert where(issues, "VIS-008") == {"Oculta": (TARGET_OUT, "")}
    assert where(issues, "EDT-018") == {"Fase": (TARGET_ROW, "1.0.0")}
    assert all(i.target == TARGET_GENERAL and not i.row for i in issues if not i.code and not i.name)


def test_issue_without_task_id_is_located_by_its_code():
    t = tasks()
    part = build_part(P(t), Config())
    found = locate_issues(part, [Issue("FEC-007", "1.1.1.2"), Issue("HIT-005", "1.1.2"), Issue("EDT-014", "9.9"),
                                 Issue("CFG-044"), Issue("FEC-007", uid=by_name(t)["Fase"].uid)])
    assert [(i.target, i.row) for i in found] == [(TARGET_DETAIL, "1.1.1"), (TARGET_ROW, "1.1.2"),
                                                  (TARGET_GENERAL, ""), (TARGET_GENERAL, ""), (TARGET_ROW, "1.0.0")]


def test_task_without_row_below_a_summary_is_detail_of_that_row():
    """Un hito sin peso que cuelga de una tarea resumen no forma la fila -D: es detalle de la fila resumen."""
    t = tree('''
        1 Fase
          1.1 Paquete
          - Hito
    ''')
    hito = by_name(t)["Hito"]
    hito.milestone = True
    part = build_part(P(t), Config())
    assert [r.code for r in part.rows] == ["1.0.0", "1.1.0"]
    (found,) = locate_issues(part, [Issue("HIT-001", uid=hito.uid)])
    assert (found.target, found.row) == (TARGET_DETAIL, "1.0.0")


def test_priority_and_groups():
    issues = [Issue("FEC-007", "1.1.1.1", "A", "x", SEV_WARN, target=TARGET_DETAIL, row="1.1.1"),
              Issue("LIB-001", message="y", severity=SEV_INFO, target=TARGET_GENERAL),
              Issue("FEC-007", "1.1.2", "B", "x", SEV_WARN, target=TARGET_ROW, row="1.1.2"),
              Issue("FEC-007", "", "C (Id 9)", "x", SEV_WARN, target=TARGET_DETAIL, row="1.1.1"),
              Issue("EDT-005", "1.14.2", "D", "x", SEV_ERROR, target=TARGET_ROW, row="1.14.2"),
              Issue("FEC-007", "", "E (Id 10)", "x", SEV_WARN, target=TARGET_OUT)]
    ordered = by_priority(issues)
    assert [(i.id, i.target) for i in ordered] == [
        ("EDT-005", TARGET_ROW), ("FEC-007", TARGET_ROW), ("FEC-007", TARGET_DETAIL), ("FEC-007", TARGET_DETAIL),
        ("FEC-007", TARGET_OUT), ("LIB-001", TARGET_GENERAL)]
    lines = grouped(ordered, "Una por tarea en otra parte.")
    assert len(lines) == 5
    group = lines[2]
    assert (group.tasks, group.name, group.code, group.row) == (2, "2 tareas", "", "1.1.1")
    assert group.message == ("Duración estimada: 2 tareas de detalle de la fila 1.1.1 (1.1.1.1, C (Id 9)). "
                             "Una por tarea en otra parte.")
    assert lines[3] is ordered[4]                   # una sola tarea fuera del libro: sin agrupar


def test_validation_sheets(tmp_path):
    path = tmp_path / "libro.xlsx"
    write_workbook(report(tasks()), Config(), path)
    detail = validation_rows(path)
    assert [(r["alcance"], r["fila"], r["edt"]) for r in detail if r["codigo"] == "FEC-007"] == [
        (TARGET_ROW, "1.1.2", "1.1.2"), (TARGET_DETAIL, "1.1.1", "1.1.1.1"), (TARGET_DETAIL, "1.1.1", "1.1.1.2"),
        (TARGET_DETAIL, "1.1.1", None), (TARGET_DETAIL, "1.0.0-D", None), (TARGET_OUT, None, None)]
    lines = validation_rows(path, grouped=True)
    fec = [r for r in lines if r["codigo"] == "FEC-007"]
    assert [(r["alcance"], r["fila"], r["tarea"]) for r in fec] == [
        (TARGET_ROW, "1.1.2", "Paquete dos"), (TARGET_DETAIL, "1.1.1", "3 tareas"),
        (TARGET_DETAIL, "1.0.0-D", "Suelta (Id 8)"), (TARGET_OUT, None, "Fuera (Id 12)")]
    assert "Una por tarea en la hoja Validación detalle." in fec[1]["detalle"]
    severities = [r["severidad"] for r in lines]
    assert severities == sorted(severities, key=[SEV_ERROR, SEV_WARN, SEV_INFO].index)

    ws = load_workbook(path)[SHEET_VALID]
    table = [[ws.cell(r, c).value for c in range(1, 7)] for r in range(5, 10)]
    assert table[0] == ["Observaciones", TARGET_GENERAL, TARGET_ROW, TARGET_DETAIL, TARGET_OUT, "Total"]
    total = table[4]
    assert total[0] == "Total" and total[5] == len(detail)
    assert total[1:5] == [sum(1 for r in detail if r["alcance"] == t)
                          for t in (TARGET_GENERAL, TARGET_ROW, TARGET_DETAIL, TARGET_OUT)]
    pv_rows = {code_key(load_workbook(path)["PV"].cell(r, 1).value or ""): r for r in range(1, 30)}
    link = ws.cell(VALID_HEADER_ROW + 1 + lines.index(fec[1]), V_ROW)
    assert link.value == "1.1.1" and link.hyperlink.location == f"'PV'!A{pv_rows['1.1.1']}"


S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)


def detail_xml(tmp_path):
    """Paquete 1.1 con dos actividades de nivel 3 con duración estimada (con nivel máximo 2, son de detalle)."""
    def task(uid, level, name, wbs, estimated=False):
        return task_xml(uid, level, name, wbs=wbs, start=S, finish=F, baselines=[(0, S, F, 0)],
                        extra="<Estimated>1</Estimated>" if estimated else "")
    return write_xml(tmp_path / "p.xml", project_xml([
        task(1, 1, "Fase", "1"), task(2, 2, "Paquete", "1.1"), task(3, 3, "Montaje", "1.1.1", True),
        task(4, 3, "Pruebas", "1.1.2", True), task(5, 2, "Cierre", "1.2", True)]))


def test_validar_groups_detail_and_shows_each_one_with_todas(tmp_path, capsys):
    path = detail_xml(tmp_path)
    assert main(["validar", "--origen", str(path), "--nivel-max", "2"]) == 0
    out = capsys.readouterr().out
    assert "[!] FEC-007 1.2 Cierre: Duración estimada" in out
    assert ("[!] FEC-007 Duración estimada: 2 tareas de detalle de la fila 1.1 (1.1.1, 1.1.2). Una por tarea con "
            "--todas.") in out
    assert out.index("FEC-007 1.2 Cierre") < out.index("2 tareas de detalle")      # primero las filas EDT
    assert ("Resultado: 5 advertencias, 3 informativas.\n  Por alcance: generales: 2 advertencias, 3 informativas; "
            "en filas EDT: 1 advertencia; en tareas de detalle: 2 advertencias.") in out
    main(["validar", "--origen", str(path), "--nivel-max", "2", "--todas"])
    out = capsys.readouterr().out
    assert "[!] FEC-007 1.1.1 Montaje (detalle de 1.1): Duración estimada" in out and "2 tareas de" not in out


def test_generar_console_groups_detail(tmp_path, capsys):
    path = detail_xml(tmp_path)
    assert main(["generar", "--origen", str(path), "--nivel-max", "2", "--salida", str(tmp_path / "l.xlsx")]) == 0
    out = capsys.readouterr().out
    assert "2 tareas de detalle de la fila 1.1 (1.1.1, 1.1.2). Una por tarea en la hoja 'Validación detalle'." in out
    assert "  Por alcance: " in out and "en tareas de detalle: 2 advertencias" in out
    assert cli._issue_line(Issue("FEC-007", "1.1.1", "Montaje", "m", target=TARGET_DETAIL, row="1.1")) == (
        "  [!] FEC-007 1.1.1 Montaje (detalle de 1.1): m")
