"""Robustez del libro: actualizar, migrar, escritura, Cuadro de Mando, CLI y configuración.

Las pruebas marcadas con xfail(strict=True) documentan hallazgos confirmados en la revisión: fallan
hoy por el motivo indicado en `reason` y deben pasar (y perder la marca) cuando se corrija el código.
"""
import re
import shutil
import sys
import zipfile
from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.writer.excel import ExcelWriter

from evm import dashboard
from evm.cli import main
from evm.config import Config, ConfigError, load_config
from evm.migrate import plan_migration, prior_from_book
from evm.model import ProjectData, Task
from evm.update import book_totals, read_book, run_update
from evm.wbs import LinearPVSource, build_report
from evm.workbook import FIRST_MONTH_COL, FIRST_ROW, WorkbookFormatError, write_workbook

S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)
ERROR = "error"     # resultado de una lectura rechazada con WorkbookFormatError


def hallazgo(reason: str, raises=AssertionError):
    """Marca de un hallazgo confirmado: la prueba falla hoy por ese motivo (y solo por ese)."""
    return pytest.mark.xfail(strict=True, reason=f"Hallazgo: {reason}", raises=raises)


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def T(uid, level, wbs, name, cost=0.0, **kw):
    kw.setdefault("cost", cost)
    return Task(uid=uid, id=kw.pop("id", uid), name=name, outline_level=level, wbs=wbs, baseline_cost=cost,
                baseline_start=kw.pop("bs", S), baseline_finish=kw.pop("bf", F), has_baseline=True,
                start=S, finish=F, **kw)


def P(tasks, name="Programa", status=None, **kw):
    return ProjectData(name=name, source="p.xml", backend="xml", tasks=tasks, status_date=status, **kw)


def make_book(path, project, cfg=None, prior=None):
    cfg = cfg or Config()
    rep = build_report(project, cfg, LinearPVSource())
    if prior is not None:
        rep.issues.extend(plan_migration(rep, prior, cfg.max_level))
    write_workbook(rep, cfg, path, prior)
    return rep


def base_tasks():
    return [T(1, 1, "1", "Fase", 300), T(2, 2, "1.1", "B", 100), T(3, 2, "1.2", "C", 200)]


def progress_tasks():
    """Avance: B 50 % (costo real 60), C 25 % (costo real 40)."""
    tasks = base_tasks()
    tasks[1].percent_complete, tasks[1].acwp = 50, 60.0
    tasks[2].percent_complete, tasks[2].acwp = 25, 40.0
    tasks[0].acwp = 100.0
    return tasks


def row_of(ws, code):
    """Fila de la hoja cuyo código EDT (columna A) es `code`."""
    return next(r for r in range(FIRST_ROW, ws.max_row + 1) if ws.cell(r, 1).value == code)


def book_row(book, code):
    return next(s.row for s in book.structure if s.code == code)


def tamper(path, fn):
    """Modifica el libro como lo haría el usuario en Excel: abre, aplica `fn(wb)` y guarda."""
    wb = load_workbook(path)
    fn(wb)
    wb.save(path)
    return path


def run_cli(capsys, *args):
    """Ejecuta la línea de comandos y devuelve (código de salida, salida estándar, salida de error)."""
    capsys.readouterr()
    rc = main([str(a) for a in args])
    captured = capsys.readouterr()
    return rc, captured.out, captured.err


def backups(folder):
    return sorted(p.name for p in folder.glob("*.respaldo-*"))


SMALL = Config(months=24)      # horizonte corto: libros más rápidos de leer y guardar


@pytest.fixture(scope="session")
def small_template(tmp_path_factory):
    """Libro de base_tasks() generado una sola vez (se copia en cada prueba)."""
    path = tmp_path_factory.mktemp("plantilla_chica") / "b.xlsx"
    make_book(path, P(base_tasks()), SMALL)
    return path


@pytest.fixture()
def small_book(tmp_path, small_template):
    path = tmp_path / "b.xlsx"
    shutil.copy(small_template, path)
    return path


@pytest.fixture(scope="session")
def example_template(tmp_path_factory, example_files):
    """Libro del proyecto de ejemplo generado una sola vez (se copia en cada prueba)."""
    out = tmp_path_factory.mktemp("plantilla") / "EVM.xlsx"
    assert main(["generar", "--origen", str(example_files["base"]), "--config", str(example_files["config"]),
                 "--salida", str(out)]) == 0
    return out


@pytest.fixture()
def example_book(tmp_path, example_template):
    path = tmp_path / "EVM.xlsx"
    shutil.copy(example_template, path)
    return path


# ----------------------------------------------------------------------------
# actualizar: cálculo y asociación de filas
# ----------------------------------------------------------------------------
def residual_tasks():
    """Fase con un gasto sin código EDT (fila 1.0.0-D) terminado y los paquetes con fila propia sin avance."""
    tasks = [T(1, 1, "1", "Fase", 400, company="EmpresaA"), T(2, 2, "1.1", "B", 100), T(3, 2, "1.2", "C", 200),
             T(4, 2, "", "Gasto sin código", 100, company="EmpresaA", percent_complete=100)]
    for t in tasks:
        t.acwp = 0.0
    tasks[0].acwp = tasks[3].acwp = 100.0
    return tasks


@pytest.mark.parametrize("multi", [
    pytest.param(False, marks=hallazgo("en una sola empresa el EV de la fila -D usa el avance de toda la tarea "
                                       "resumen, incluidos los paquetes con fila propia")),
    True,
], ids=["una_empresa", "multiempresa"])
def test_residual_row_ev_uses_its_own_tasks(tmp_path, multi):
    cfg = Config(multi_enabled=multi, months=24)
    book = tmp_path / "b.xlsx"
    make_book(book, P(residual_tasks()), cfg)
    res = run_update(P(residual_tasks()), cfg, book, date(2026, 1, 31), backup=False)
    residual = next(r for r in res.rows if r.code == "1.0.0-D")
    assert residual.bac == pytest.approx(100)
    assert residual.ev_cum == pytest.approx(100)          # su única tarea está terminada


@hallazgo("no se detectan filas insertadas en las hojas EV/AC: los valores se escriben en otras filas",
          raises=pytest.fail.Exception)
def test_update_rejects_shifted_rows(small_book):
    def insert_row(wb):
        ws = wb["EV"]
        ws.insert_rows(6)                                   # como "Insertar fila" en Excel
        ws.cell(6, 1, "Nota")
    tamper(small_book, insert_row)
    before = small_book.read_bytes()
    with pytest.raises(WorkbookFormatError, match="EV"):
        run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
    assert small_book.read_bytes() == before


@pytest.mark.parametrize("with_residual", [
    False,
    pytest.param(True, marks=hallazgo("el costo real de un paquete nuevo con código EDT se suma sin aviso en la "
                                      "fila -D de su tarea resumen")),
], ids=["sin_fila_D", "con_fila_D"])
def test_new_coded_package_without_row_warns(tmp_path, with_residual):
    tasks = base_tasks()
    if with_residual:
        tasks[0].baseline_cost = tasks[0].cost = 400
        tasks.append(T(4, 2, "", "Gasto sin código", 100))
    book = tmp_path / "b.xlsx"
    make_book(book, P(tasks), SMALL)
    now = [T(t.uid, t.outline_level, t.wbs, t.name, t.baseline_cost, acwp=0.0) for t in tasks]
    now.append(T(9, 2, "1.3", "Paquete nuevo", 0, acwp=500.0, percent_complete=50))
    now[0].acwp = 500.0
    res = run_update(P(now, summary_acwp=500.0), Config(), book, date(2026, 1, 31), dry_run=True)
    assert any("1.3 'Paquete nuevo'" in w and "--migrar-desde" in w for w in res.warnings)


@hallazgo("no se detecta una fórmula de la fila TOTAL reemplazada por un valor")
def test_overwritten_total_formula_is_reported(small_book):
    total = row_of(load_workbook(small_book)["AC"], "TOTAL")
    tamper(small_book, lambda wb: wb["AC"].cell(total, FIRST_MONTH_COL, 0))
    try:
        res = run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
        message = " ".join(res.warnings)
    except WorkbookFormatError as exc:
        message = str(exc)
    assert f"AC!G{total}" in message


@hallazgo("una tarea desactivada después de generar el libro sigue cargando EV y AC sin aviso")
def test_task_made_inactive_after_generation_is_reported(small_book):
    tasks = progress_tasks()
    tasks[2].active = False                                  # C se desactiva en MS Project
    res = run_update(P(tasks, summary_acwp=60.0), Config(), small_book, date(2026, 1, 31), dry_run=True)
    loaded = {r.code for r in res.rows}
    assert "1.2.0" not in loaded or any("inactiv" in w.lower() for w in res.warnings)


def renamed_tasks(shift=0):
    """Fase con dos paquetes y una tarea sin código (fila 1.0.0-D), con avance; UID desplazados en `shift`."""
    tasks = [T(1 + shift, 1, "1", "Fase", 400), T(2 + shift, 2, "1.1", "B", 100),
             T(3 + shift, 2, "1.2", "C", 200), T(4 + shift, 2, "", "Sin código", 100)]
    tasks[1].percent_complete, tasks[1].acwp = 50, 60.0
    tasks[2].percent_complete, tasks[2].acwp = 25, 40.0
    tasks[3].percent_complete, tasks[3].acwp = 100, 30.0
    tasks[0].acwp = 130.0
    return tasks


@pytest.mark.parametrize("cfg", [
    Config(),
    pytest.param(Config(max_level=4), marks=hallazgo("con otro nivel máximo no se asocian las filas de un proyecto "
                                                     "renombrado (se usa el nivel de la configuración y no el del "
                                                     "libro)")),
], ids=["mismo_nivel", "nivel_4"])
def test_renamed_project_update_matches_by_code(tmp_path, cfg):
    book = tmp_path / "b.xlsx"
    make_book(book, P(renamed_tasks()), SMALL)
    status = date(2026, 1, 31)
    same = run_update(P(renamed_tasks()), Config(), book, status, dry_run=True)
    renamed = run_update(P(renamed_tasks(100), name="Programa (copia)"), cfg, book, status, dry_run=True)

    def values(res):
        return {r.code: (r.ev_cum, r.ac_cum) for r in res.rows}

    assert values(renamed) == values(same)
    assert set(values(same)) == {"1.1.0", "1.2.0", "1.0.0-D"}
    assert any("se asocian por código EDT" in w for w in renamed.warnings)


# ----------------------------------------------------------------------------
# actualizar: fecha de corte, simulación y --forzar
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("status", [date(2025, 12, 31), date(2027, 1, 31)], ids=["antes", "despues"])
def test_cutoff_outside_horizon_is_rejected(tmp_path, status):
    book = tmp_path / "b.xlsx"
    make_book(book, P(base_tasks()), Config(months=12))
    before = book.read_bytes()
    with pytest.raises(ValueError, match="fuera del horizonte"):
        run_update(P(progress_tasks()), Config(), book, status, backup=False)
    assert book.read_bytes() == before


def test_cutoff_warnings_and_negative_corrections(tmp_path):
    book = tmp_path / "b.xlsx"
    make_book(book, P(base_tasks()), Config(months=12))
    res = run_update(P(progress_tasks()), Config(), book, date(2026, 1, 15), dry_run=True)
    assert any("último día del mes" in w for w in res.warnings)

    # Menos avance que en el mes anterior: valor del mes negativo, informado como corrección
    run_update(P(progress_tasks()), Config(), book, date(2026, 2, 28), backup=False)
    less = progress_tasks()
    less[1].percent_complete, less[1].acwp = 10, 20.0
    res = run_update(P(less), Config(), book, date(2026, 3, 31), backup=False)
    row = next(r for r in res.rows if r.code == "1.1.0")
    assert row.ev_month == pytest.approx(-40) and row.ac_month == pytest.approx(-40)
    assert "EV del mes negativo (corrección)" in row.note and "AC del mes negativo (corrección)" in row.note

    # Fecha de corte anterior a la última cargada (sin valores posteriores de AC): se advierte
    other = tmp_path / "b2.xlsx"
    make_book(other, P(base_tasks()), Config(months=12))
    run_update(P(progress_tasks()), Config(), other, date(2026, 2, 28), only="ev", backup=False)
    res = run_update(P(progress_tasks()), Config(), other, date(2026, 1, 31), only="ac", backup=False)
    assert any("anterior a la última cargada" in w for w in res.warnings)


def test_force_allows_loading_an_earlier_month(small_book):
    run_update(P(progress_tasks()), Config(), small_book, date(2026, 2, 28), backup=False)
    before = small_book.read_bytes()
    with pytest.raises(ValueError, match="posteriores"):
        run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
    assert small_book.read_bytes() == before                 # sin --forzar no se modifica
    res = run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False, force=True)
    assert any("anterior a la última cargada" in w for w in res.warnings)
    b = read_book(small_book)
    assert b.status_date == date(2026, 1, 31)
    assert b.ev[book_row(b, "1.1.0")][0] == pytest.approx(50)   # EV acumulado a enero (sin meses previos)


def test_dry_run_is_side_effect_free(tmp_path, small_template):
    folder = tmp_path / "simulacion"
    folder.mkdir()
    book = folder / "b.xlsx"
    shutil.copy(small_template, book)
    tamper(book, lambda wb: wb["AC"].cell(row_of(wb["AC"], "1.1.0"), FIRST_MONTH_COL, "60"))
    data, mtime = book.read_bytes(), book.stat().st_mtime_ns
    res = run_update(P(progress_tasks()), Config(), book, date(2026, 2, 28), dry_run=True)   # respaldo activado
    assert res.dry_run and res.backup is None
    assert book.read_bytes() == data and book.stat().st_mtime_ns == mtime
    assert [p.name for p in folder.iterdir()] == ["b.xlsx"]
    assert res.book.ac[book_row(res.book, "1.1.0")][0] == 60                # valor convertido solo en memoria
    assert any("AC!G" in w for w in res.warnings)


# ----------------------------------------------------------------------------
# actualizar: celdas de ingreso con texto
# ----------------------------------------------------------------------------
def load_text_in_ac(path, text):
    """Valor de AC!G (fila 1.1.0) después de 'actualizar' con `text` escrito como texto, o ERROR."""
    tamper(path, lambda wb: wb["AC"].cell(row_of(wb["AC"], "1.1.0"), FIRST_MONTH_COL, text))
    try:
        run_update(P(progress_tasks()), Config(), path, date(2026, 2, 28), only="ev", backup=False)
    except WorkbookFormatError:
        return ERROR
    ws = load_workbook(path)["AC"]
    return ws.cell(row_of(ws, "1.1.0"), FIRST_MONTH_COL).value


@pytest.mark.parametrize("text, accepted", [
    ("60", {60}),
    ("12,5", {12.5}),
    ("1 234", {1234}),
    ("1.234,5", {1234.5, ERROR}),
    ("1,234.50", {1234.5, ERROR}),
    pytest.param("1.234", {1234, ERROR}, marks=hallazgo("el texto '1.234' (separador de miles) se lee como 1,234")),
    pytest.param("nan", {ERROR}, marks=hallazgo("el texto 'nan' se acepta como número y la celda queda vacía")),
])
def test_text_numbers_in_input_cells(small_book, text, accepted):
    assert load_text_in_ac(small_book, text) in accepted


# ----------------------------------------------------------------------------
# Guardado del libro
# ----------------------------------------------------------------------------
@hallazgo("el guardado no es atómico: un error durante el guardado destruye el libro")
def test_failed_save_keeps_original_book(small_book, monkeypatch):
    before = small_book.read_bytes()

    def fail(self, ws):
        raise OSError("disco lleno (simulado)")

    monkeypatch.setattr(ExcelWriter, "write_worksheet", fail)
    with pytest.raises(OSError, match="simulado"):
        run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
    monkeypatch.undo()
    assert small_book.read_bytes() == before


def lock_on_save(monkeypatch):
    """Simula el libro abierto en Excel: guardar falla con PermissionError."""
    def locked(self, filename):
        raise PermissionError(13, "El proceso no tiene acceso al archivo", str(filename))
    monkeypatch.setattr(Workbook, "save", locked)


def test_locked_book_reports_clear_error(small_book, monkeypatch):
    before = small_book.read_bytes()
    lock_on_save(monkeypatch)
    with pytest.raises(PermissionError, match="se debe cerrar"):
        run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31))
    assert small_book.read_bytes() == before


@hallazgo("con el libro bloqueado queda un respaldo huérfano en cada intento")
def test_locked_book_leaves_no_orphan_backup(small_book, monkeypatch):
    lock_on_save(monkeypatch)
    with pytest.raises(PermissionError):
        run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31))
    assert backups(small_book.parent) == []


@pytest.mark.skipif(sys.platform != "win32", reason="bloqueo de archivos de Windows")
def test_real_windows_lock_reports_clear_error(small_book):
    win32file = pytest.importorskip("win32file")
    before = small_book.read_bytes()
    handle = win32file.CreateFile(str(small_book), win32file.GENERIC_READ, win32file.FILE_SHARE_READ, None,
                                  win32file.OPEN_EXISTING, 0, None)        # como Excel: solo lectura compartida
    try:
        with pytest.raises(PermissionError, match="se debe cerrar"):
            run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
    finally:
        handle.Close()
    assert small_book.read_bytes() == before


# ----------------------------------------------------------------------------
# Libro modificado o dañado
# ----------------------------------------------------------------------------
def _rename_ev(wb):
    wb["EV"].title = "EV mensual"


def _delete(name):
    def fn(wb):
        del wb[name]
    return fn


def _meta(key, value=None, delete=False):
    def fn(wb):
        ws = wb["_meta"]
        r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == key)
        if delete:
            ws.delete_rows(r)
        else:
            ws.cell(r, 2).value = value
    return fn


@pytest.mark.parametrize("mutate", [
    pytest.param(_rename_ev, id="hoja_EV_renombrada", marks=hallazgo("hoja EV renombrada: error inesperado (KeyError)")),
    pytest.param(_delete("Resumen"), id="sin_Resumen", marks=hallazgo("hoja Resumen eliminada: error inesperado")),
    pytest.param(_delete("_graficos"), id="sin_graficos", marks=hallazgo("hoja _graficos eliminada: error inesperado")),
    pytest.param(_meta("horizonte_inicio", "2026-01-01"), id="inicio_como_texto",
                 marks=hallazgo("inicio del horizonte guardado como texto en _meta: error inesperado")),
    pytest.param(_meta("meses", "abc"), id="meses_no_numerico"),
    pytest.param(_meta("fila_total", delete=True), id="falta_fila_total"),
    pytest.param(_delete("_meta"), id="sin_meta"),
])
def test_tampered_book_fails_with_format_error(example_book, example_files, capsys, mutate):
    tamper(example_book, mutate)
    before = example_book.read_bytes()
    rc, _out, err = run_cli(capsys, "actualizar", "--origen", example_files["jun"], "--config",
                            example_files["config"], "--libro", example_book, "--sin-respaldo")
    assert example_book.read_bytes() == before
    assert "inesperado" not in err and "ERROR:" in err
    assert rc == 2


def test_missing_company_sheet_in_multi_book(tmp_path):
    cfg = Config(multi_enabled=True, months=24)
    book = tmp_path / "multi.xlsx"
    make_book(book, P(base_tasks()), cfg)
    tamper(book, _delete("EV-EmpresaB"))
    with pytest.raises(WorkbookFormatError, match="Falta la hoja 'EV-EmpresaB'"):
        run_update(P(progress_tasks()), cfg, book, date(2026, 1, 31), backup=False)


# ----------------------------------------------------------------------------
# Escritura del libro y Cuadro de Mando
# ----------------------------------------------------------------------------
def feature_snapshot(path):
    """Elementos del libro que openpyxl debe conservar al reabrirlo y guardarlo."""
    wb = load_workbook(path)
    sheets = {ws.title: (len(ws.data_validations.dataValidation), len(list(ws.conditional_formatting)),
                         ws.sheet_state, str(ws.freeze_panes), len(ws.merged_cells.ranges)) for ws in wb.worksheets}
    return sheets, sorted(wb.defined_names.keys()), wb["Cuadro de Mando"].print_area


def chart_count(path):
    with zipfile.ZipFile(path) as z:
        return sum(1 for n in z.namelist() if re.fullmatch(r"xl/charts/chart\d+\.xml", n))


def test_update_roundtrip_preserves_workbook_features(small_book):
    before = feature_snapshot(small_book)
    assert chart_count(small_book) == 3
    run_update(P(progress_tasks()), Config(), small_book, date(2026, 1, 31), backup=False)
    run_update(P(progress_tasks()), Config(), small_book, date(2026, 2, 28), backup=False)
    assert feature_snapshot(small_book) == before
    assert chart_count(small_book) == 3


@pytest.mark.parametrize("months", [
    1, 14, 336,
    pytest.param(337, marks=hallazgo("eje de fechas sin paso para más de 336 meses", raises=StopIteration)),
    pytest.param(600, marks=hallazgo("eje de fechas sin paso para más de 336 meses", raises=StopIteration)),
])
def test_date_axis_supports_every_horizon(months):
    assert dashboard._date_axis(months).majorUnit >= 1


@hallazgo("un plan de más de 336 meses impide generar el libro (StopIteration)", raises=StopIteration)
def test_generate_and_update_book_beyond_336_months(tmp_path):
    tasks = base_tasks()
    tasks[1].baseline_finish = datetime(2055, 12, 28, 17)       # 30 años de plan
    book = tmp_path / "largo.xlsx"
    make_book(book, P(tasks), Config(months=400))
    res = run_update(P(progress_tasks()), Config(), book, date(2054, 12, 31), backup=False)
    assert res.month_index == 347
    assert chart_count(book) == 3


@pytest.mark.parametrize("where", ["tarea", "proyecto"])
@hallazgo("un carácter de control en un nombre impide generar el libro", raises=IllegalCharacterError)
def test_control_characters_in_names(tmp_path, where):
    tasks, name = base_tasks(), "Programa"
    if where == "tarea":
        tasks[1].name = "B\x07campana"
    else:
        name = "Prog\x0bVT"
    book = tmp_path / "b.xlsx"
    make_book(book, P(tasks, name=name), SMALL)
    ws = load_workbook(book)["PV"]
    assert ws.cell(row_of(ws, "1.1.0"), 2).value == ("Bcampana" if where == "tarea" else "B")
    res = run_update(P(progress_tasks(), name=name), Config(), book, date(2026, 1, 31), dry_run=True)
    assert not any("fue generado para el proyecto" in w for w in res.warnings)


# ----------------------------------------------------------------------------
# --migrar-desde
# ----------------------------------------------------------------------------
def old_book_with_ev(path):
    """Libro anterior con EV en enero (10), febrero (20) y mayo (5) de 2026 y fecha de corte en mayo."""
    make_book(path, P(base_tasks()), SMALL)

    def load(wb):
        ws = wb["EV"]
        r = row_of(ws, "1.1.0")
        for i, v in ((0, 10), (1, 20), (4, 5)):
            ws.cell(r, FIRST_MONTH_COL + i, v)
        wb["Resumen"]["B5"].value = date(2026, 5, 31)
    tamper(path, load)
    return path


def test_migration_renamed_project_by_code(tmp_path):
    old = tmp_path / "old.xlsx"
    make_book(old, P(base_tasks()), SMALL)

    def load(wb):
        for code, v in (("1.1.0", 10), ("1.2.0", 20)):
            wb["EV"].cell(row_of(wb["EV"], code), FIRST_MONTH_COL, v)
    tamper(old, load)
    new_tasks = [T(11, 1, "1", "Fase", 300), T(12, 2, "1.1", "B", 100), T(13, 2, "1.2", "C renombrada", 200)]
    new = tmp_path / "new.xlsx"
    rep = make_book(new, P(new_tasks, name="Otro nombre"), SMALL, prior=prior_from_book(old, "Otro nombre"))
    b = read_book(new)
    assert {s.code: sum(v or 0 for v in b.ev[s.row]) for s in b.input_rows} == {"1.1.0": 10, "1.2.0": 20}
    messages = [(i.severity, i.code, i.message) for i in rep.issues]
    assert any(sev == "Advertencia" and "corresponde al proyecto 'Programa'" in m for sev, _c, m in messages)
    assert any(code == "1.2.0" and "Recibió el historial" in m for _s, code, m in messages)


def test_migration_values_after_new_horizon_end_are_accumulated(tmp_path):
    old = old_book_with_ev(tmp_path / "old.xlsx")
    new = tmp_path / "new.xlsx"
    rep = make_book(new, P(base_tasks()), Config(months=2), prior=prior_from_book(old, "Programa"))
    b = read_book(new)
    assert b.ev[book_row(b, "1.1.0")] == [10, 25]
    assert any("después del último mes" in i.message for i in rep.issues)


@hallazgo("los EV/AC anteriores al nuevo mes 1 se acumulan en ese mes sin aviso")
def test_migration_values_before_new_horizon_start_warn(tmp_path):
    old = old_book_with_ev(tmp_path / "old.xlsx")
    new = tmp_path / "new.xlsx"
    rep = make_book(new, P(base_tasks()), Config(start_month=date(2026, 4, 1), months=12),
                    prior=prior_from_book(old, "Programa"))
    b = read_book(new)
    assert sum(book_totals(b)[1]) == pytest.approx(35)        # el total se conserva
    assert any("EV/AC" in i.message and "antes" in i.message for i in rep.issues)


@hallazgo("la migración descarta los meses cargados con 0")
def test_migration_keeps_explicit_zero_months(tmp_path):
    old = tmp_path / "old.xlsx"
    make_book(old, P(base_tasks()), SMALL)
    no_progress = base_tasks()
    for t in no_progress:
        t.acwp = 0.0
    run_update(P(no_progress), Config(), old, date(2026, 1, 31), backup=False)       # enero cargado con 0
    run_update(P(progress_tasks()), Config(), old, date(2026, 2, 28), backup=False)
    new = tmp_path / "new.xlsx"
    make_book(new, P(base_tasks()), SMALL, prior=prior_from_book(old, "Programa"))
    a, b = read_book(old), read_book(new)
    row = book_row(a, "1.1.0")
    assert a.ev[row][:2] == [0.0, 50.0]
    assert b.ev[book_row(b, "1.1.0")][:2] == [0.0, 50.0]


# ----------------------------------------------------------------------------
# Línea de comandos
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("args, message", [
    (["generar", "--origen", "{tmp}/no_existe.xml", "--salida", "{tmp}/x.xlsx"], "No existe el archivo de origen"),
    (["generar", "--origen", "{tmp}/p.txt", "--salida", "{tmp}/x.xlsx"], "No se reconoce la extensión"),
    (["generar", "--origen", "{tmp}/p.xml", "--salida", "{tmp}/x.xls"], "debe ser un archivo .xlsx"),
    (["generar", "--origen", "{tmp}/p.xml", "--salida", "{tmp}/x.xlsx", "--migrar-desde", "{tmp}/x.xlsx"],
     "debe ser un archivo distinto"),
    (["actualizar", "--origen", "{tmp}/p.xml", "--libro", "{tmp}/no_existe.xlsx"], "No existe el libro"),
    (["actualizar", "--origen", "{tmp}/p.xml", "--libro", "{tmp}/x.xlsx", "--fecha-corte", "2026-02-30"],
     "Fecha inválida"),
], ids=["origen_inexistente", "extension", "salida_xls", "migrar_igual_salida", "libro_inexistente",
        "fecha_invalida"])
def test_cli_errors_exit_with_code_2(tmp_path, capsys, args, message):
    rc, _out, err = run_cli(capsys, *[a.replace("{tmp}", str(tmp_path)) for a in args])
    assert rc == 2
    assert err.startswith("ERROR: ") and message in err


@pytest.mark.parametrize("error, code, message", [
    (KeyError("x"), 1, "ERROR inesperado: KeyError"),
    (KeyboardInterrupt(), 130, "Cancelado."),
    (RuntimeError("falla de COM"), 2, "ERROR: falla de COM"),
], ids=["inesperado", "cancelado", "runtime"])
def test_cli_unexpected_errors(tmp_path, capsys, monkeypatch, error, code, message):
    def boom(args):
        raise error
    monkeypatch.setattr("evm.cli.cmd_generar", boom)
    rc, _out, err = run_cli(capsys, "generar", "--origen", tmp_path / "p.xml")
    assert rc == code and message in err
    if code == 1:
        assert "-v" in err                                   # indica cómo ver el detalle técnico


@pytest.mark.parametrize("option", [
    pytest.param(["--meses", "0"], id="meses_0"),
    pytest.param(["--nivel-max", "0"], id="nivel_max_0"),
])
@hallazgo("--meses 0 y --nivel-max 0 se ignoran en silencio (se usan 120 meses y nivel 3)")
def test_cli_rejects_zero_overrides(tmp_path, capsys, example_files, option):
    rc, _out, err = run_cli(capsys, "generar", "--origen", example_files["base"], "--config",
                            example_files["config"], "--salida", tmp_path / "x.xlsx", *option)
    assert rc == 2 and "debe estar entre" in err


# ----------------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("content, message", [
    pytest.param(b'[exportacion]\nmoneda = 5\n', "debe ser texto", id="moneda_numerica"),
    pytest.param(b'[metodos]\nlista = "0/100"\n', "lista de textos", id="lista_texto"),
    pytest.param(b'[exportacion]\nferiados = "2026-01-01"\n', "lista de fechas", id="feriados_texto"),
    pytest.param(b'[horizonte]\nmeses = 0\n', "meses debe estar entre", id="meses_0"),
    pytest.param(b'[exportacion]\ndecimales = 5\n', "decimales", id="decimales_5"),
    pytest.param(b'[semaforo]\nrojo = 0.97\namarillo = 0.95\n', "umbrales", id="umbrales_invertidos"),
    pytest.param(b'[campos\n', "Error de sintaxis", id="sintaxis"),
    pytest.param('[campos]\nmetodo = "Método EVM"\n'.encode("cp1252"), "UTF-8", id="ansi_cp1252",
                 marks=hallazgo("un archivo en ANSI produce un error de decodificación en inglés",
                                raises=UnicodeDecodeError)),
    # Sección que no es tabla y enteros con decimales: ver test_robustez_calculo.py
])
def test_config_rejects_invalid_files(tmp_path, content, message):
    path = tmp_path / "config.toml"
    path.write_bytes(content)
    with pytest.raises(ConfigError, match=message):
        load_config(path)


def test_config_missing_file(tmp_path):
    with pytest.raises(ConfigError, match="No se encontró"):
        load_config(tmp_path / "no_existe.toml")


@hallazgo("un archivo UTF-8 con BOM (Bloc de notas, PowerShell 5.1) se rechaza como error de sintaxis",
          raises=ConfigError)
def test_config_accepts_utf8_with_bom(tmp_path):
    path = tmp_path / "config.toml"
    path.write_bytes('[campos]\nmetodo = "Método EVM"\n'.encode("utf-8-sig"))
    assert load_config(path).method_field == "Método EVM"
