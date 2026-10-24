"""Comando validar y controles del archivo de MS Project: códigos EDT, fechas y duración, pesos, método,
empresas, avance a la fecha de estado y lógica del cronograma."""
from datetime import date, datetime

import pytest
from openpyxl import load_workbook

from evm.cli import main
from evm.config import Config, ConfigError, load_config
from evm.model import CONSTRAINT_MSO, LINK_FS, LINK_SF, LINK_SS, Link
from evm.reader_xml import duration_minutes, read_xml
from evm.validation import logic_checks, progress_checks, validate_part
from evm.wbs import (SEV_ERROR, SEV_INFO, SEV_WARN, TOPIC_DATES, TOPIC_LOGIC, TOPIC_METHOD, TOPIC_PROGRESS,
                     TOPIC_WBS, TOPIC_WEIGHTS, build_part, build_report)

from builders import P, by_name, tree
from mspdi import NUMBER1, TEXT10, project_xml, task_xml, write_xml

S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)


def report(tasks, cfg=None, **project_kw):
    cfg = cfg or Config()
    cfg.validate()
    return build_report(P(tasks, **project_kw), cfg)


def found(rep_or_issues, topic=None, severity=None, text=""):
    issues = getattr(rep_or_issues, "issues", rep_or_issues)
    return [i for i in issues if (topic is None or i.topic == topic) and (severity is None or i.severity == severity)
            and text in i.message]


# ----------------------------------------------------------------------------
# Códigos EDT
# ----------------------------------------------------------------------------
def test_child_code_must_extend_parent_code():
    tasks = tree('''
        1 Fase
          1.2 Grupo
            1.2.1 Bien
            1.14.2 Mal
    ''')
    rep = report(tasks)
    errors = found(rep, TOPIC_WBS, SEV_ERROR)
    assert [(i.code, i.name) for i in errors] == [("1.14.2", "Mal")]
    assert "no depende del de su tarea superior 1.2.0 «Grupo»: se esperaba 1.2.x" in errors[0].message


def test_valid_codes_have_no_wbs_observations():
    tasks = tree('''
        1 Fase
          1.1 A
            1.1.1 A1
            1.1.2 A2
              - Actividad sin código
          1.2 B
        2 Fase 2
    ''')
    rep = report(tasks)
    assert not found(rep, TOPIC_WBS, SEV_ERROR) and not found(rep, TOPIC_WBS, SEV_WARN)


def test_code_format_errors():
    tasks = tree('''
        1 Fase
          1.0.2 Cero
          1..3 Vacio
          1.A Letra
    ''')
    rep = report(tasks)
    by = {i.name: i for i in found(rep, TOPIC_WBS) if i.name in ("Cero", "Vacio", "Letra")}
    assert by["Cero"].severity == SEV_ERROR and "el 0 solo se admite al final" in by["Cero"].message
    assert by["Vacio"].severity == SEV_ERROR and "formato no válido" in by["Vacio"].message
    assert by["Letra"].severity == SEV_WARN and "no numérico" in by["Letra"].message


def test_skipped_level_order_gaps_and_duplicates():
    tasks = tree('''
        1 Fase
          1.2 B
            1.2.3.1 Salta
          1.1 A
        2 Fase 2
          2.1 X
          2.5 Y
          2.5 Z
    ''')
    rep = report(tasks)
    wbs = found(rep, TOPIC_WBS)
    salta = next(i for i in wbs if i.name == "Salta")
    assert salta.severity == SEV_WARN and "salta un nivel" in salta.message
    order = next(i for i in wbs if i.name == "A")
    assert order.severity == SEV_WARN and "fuera de orden" in order.message and "1.2.0 «B»" in order.message
    gaps = [i for i in wbs if "saltos" in i.message]
    assert len(gaps) == 1 and gaps[0].severity == SEV_INFO and "2.1.0 → 2.5.0" in gaps[0].message
    dup = [i for i in wbs if "duplicado" in i.message]
    assert len(dup) == 1 and dup[0].severity == SEV_ERROR and dup[0].code == "2.5.0" and "2 tareas" in dup[0].message


def test_code_without_upper_code():
    tasks = tree('''
        - Contenedor
          1.2 Suelto
    ''')
    rep = report(tasks)
    issue = next(i for i in found(rep, TOPIC_WBS) if i.name == "Suelto")
    assert issue.severity == SEV_WARN and "debería depender de la tarea 1.0.0" in issue.message


def test_plan_code_is_not_compared():
    """Sin planificación elegida, INTERNO y CLIENTE son contenedores: sus subtareas no se comparan con ellos (pero
    los códigos repetidos entre planificaciones son duplicados)."""
    tasks = tree('''
        INTERNO Interna
          1 Fase
            1.1 A
        CLIENTE Cliente
          1 Fase
    ''')
    rep = report(tasks)
    wbs = found(rep, TOPIC_WBS)
    assert not [i for i in wbs if "no depende" in i.message or "no numérico" in i.message]
    assert [i.code for i in wbs if "duplicado" in i.message] == ["1.0.0"]
    assert "planificacion" in next(i for i in wbs if "duplicado" in i.message).message
    rep = report(tasks, Config(plan="INTERNO"))
    assert not found(rep, TOPIC_WBS, SEV_ERROR) and not found(rep, TOPIC_WBS, SEV_WARN)


def test_project_code_zero_is_a_container():
    tasks = tree('''
        0 Proyecto
          1 Fase
            1.1 A
          2 Fase 2
    ''')
    rep = report(tasks)
    assert not found(rep, TOPIC_WBS, SEV_ERROR) and not found(rep, TOPIC_WBS, SEV_WARN)


def test_visible_no_branch_codes_are_checked():
    tasks = tree('''
        1 Fase
          1.1 A
          2.7 Oculta
    ''')
    by_name(tasks)["Oculta"].visible = "No"
    rep = report(tasks)
    assert [i.name for i in found(rep, TOPIC_WBS, SEV_ERROR)] == ["Oculta"]


# ----------------------------------------------------------------------------
# Fechas y duración
# ----------------------------------------------------------------------------
def test_dates_and_duration():
    tasks = tree('''
        1 Fase
          1.1 A
            - Al reves
            - Hito largo
            - Cero
            - Estimada
            - Manual
    ''')
    t = by_name(tasks)
    t["Al reves"].start, t["Al reves"].finish = F, S
    t["Al reves"].baseline_start, t["Al reves"].baseline_finish = F, S
    t["Hito largo"].milestone, t["Hito largo"].duration = True, 480.0
    t["Cero"].duration = 0.0
    t["Estimada"].estimated, t["Estimada"].duration = True, 2400.0
    t["Manual"].manual = True
    rep = report(tasks)
    dates = found(rep, TOPIC_DATES)
    reverse = [i for i in dates if i.name.startswith("Al reves")]
    assert {i.severity for i in reverse} == {SEV_ERROR} and len(reverse) == 2
    assert reverse[0].name == "Al reves (Id 3)"                 # sin código EDT: con su Id
    assert any(i.name.startswith("Hito largo") and "marcada como hito y tiene duración" in i.message for i in dates)
    assert any(i.name.startswith("Cero") and "Duración 0" in i.message for i in dates)
    assert any("1 tarea tiene duración estimada" in i.message for i in dates)
    assert any(i.name.startswith("Estimada") and i.message == "Duración estimada («?»)." for i in dates)
    manual = [i for i in dates if "manualmente" in i.message]
    assert len(manual) == 1 and manual[0].severity == SEV_INFO and "Id 7" in manual[0].message


def test_duration_unknown_is_not_checked():
    """Sin el dato de la duración (XML sin <Duration>), no se revisan los hitos con duración ni la duración 0."""
    tasks = tree('''
        1 Fase
          1.1 A
            - Hito
    ''')
    by_name(tasks)["Hito"].milestone = True
    rep = report(tasks)
    assert not found(rep, TOPIC_DATES)


# ----------------------------------------------------------------------------
# Hitos y pesos, método de medición, empresas
# ----------------------------------------------------------------------------
def test_milestone_weight_checks():
    tasks = tree('''
        1 Fase
          1.1 A
            - Tarea con peso
          1.2 B
            - Tarea
    ''')
    t = by_name(tasks)
    t["Tarea con peso"].milestone_weight = 30
    t["B"].method = "Hitos ponderados"
    rep = report(tasks, Config(milestone_weight_field="Peso hito", method_field="Método EVM"))
    weights = found(rep, TOPIC_WEIGHTS)
    assert any(i.name.startswith("Tarea con peso") and "no es un hito" in i.message for i in weights)
    assert any(i.code == "1.2.0" and "no tiene hitos" in i.message for i in weights)


def test_weight_field_without_values():
    tasks = tree('''
        1 Fase
          1.1 A
            - Tarea
            - Hito
    ''')
    by_name(tasks)["Hito"].milestone = True
    rep = report(tasks, Config(milestone_weight_field="Peso hito"), fields_found={"peso_hito": "Peso hito"})
    assert len(found(rep, text="ningún hito tiene peso")) == 1


def test_method_checks():
    long_s, long_f = datetime(2026, 1, 5, 8), datetime(2026, 6, 30, 17)
    tasks = tree('''
        1 Fase
          1.1 Corto
          1.2 Largo
          1.3 Medio
          1.4 Raro
          1.5 Sin metodo
            - Sub
    ''')
    t = by_name(tasks)
    t["Corto"].method = "0/100"
    t["Largo"].method = "0/100"
    t["Medio"].method = "50/50"
    t["Raro"].method = "60/30"
    t["Sub"].method = "50/50"
    for name in ("Largo", "Medio"):
        t[name].baseline_start, t[name].baseline_finish = long_s, long_f
    t["Medio"].baseline_finish = datetime(2026, 3, 20, 17)          # 2,5 meses
    t["Corto"].baseline_start, t["Corto"].baseline_finish = datetime(2026, 1, 5, 8), datetime(2026, 1, 30, 17)
    rep = report(tasks, Config(method_field="Método EVM"))
    method = {i.code: i for i in found(rep, TOPIC_METHOD) if i.code}
    assert "1.1.0" not in method
    assert method["1.2.0"].severity == SEV_WARN and "0/100 en un paquete de 5,8 meses" in method["1.2.0"].message
    assert "el EV queda en el 50 %" in method["1.3.0"].message and "meses_formula_fija" in method["1.3.0"].message
    assert "no reconocido" in method["1.4.0"].message
    assert method["1.5.0"].severity == SEV_WARN and "La fila no tiene método" in method["1.5.0"].message

    # Umbrales configurables; 0 = no se controla
    rep = report(tasks, Config(method_field="Método EVM", max_months_0_100=6, max_months_fixed=0))
    codes = {i.code for i in found(rep, TOPIC_METHOD) if "paquete de" in i.message}
    assert codes == set()


def test_validation_thresholds_in_config(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text("[validacion]\nmeses_0_100 = 2\nmeses_formula_fija = 3.5\n", encoding="utf-8")
    cfg = load_config(path)
    assert (cfg.max_months_0_100, cfg.max_months_fixed) == (2.0, 3.5)
    path.write_text("[validacion]\nmeses_0_100 = -1\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="meses_0_100"):
        load_config(path)


def test_split_on_task_that_is_not_both():
    tasks = tree('''
        1 Fase
          1.1 A
          1.2 B
    ''')
    t = by_name(tasks)
    t["A"].company, t["A"].split_a = "EmpresaA", 70
    t["B"].company, t["B"].split_a = "Ambas", 70
    rep = report(tasks, Config(multi_enabled=True))
    ignored = found(rep, text="el reparto solo se usa")
    assert [i.code for i in ignored] == ["1.1.0"] and ignored[0].severity == SEV_INFO


# ----------------------------------------------------------------------------
# Avance a la fecha de estado (solo validar)
# ----------------------------------------------------------------------------
def progress(tasks, status, cfg=None, **project_kw):
    project = P(tasks, **project_kw)
    part = build_part(project, cfg or Config())
    return progress_checks(project, part.scope, part.code_of, status)


def test_actual_dates_consistency():
    tasks = tree('''
        1 Fase
          1.1 Fin sin comienzo
          1.2 Al reves
          1.3 Avance sin comienzo
          1.4 Completa sin fin
          1.5 Fin con avance parcial
    ''')
    t = by_name(tasks)
    t["Fin sin comienzo"].actual_finish = datetime(2026, 2, 1)
    t["Al reves"].actual_start, t["Al reves"].actual_finish = datetime(2026, 2, 10), datetime(2026, 2, 1)
    t["Al reves"].percent_complete = 100
    t["Avance sin comienzo"].percent_complete = 40
    t["Completa sin fin"].percent_complete, t["Completa sin fin"].actual_start = 100, S
    t["Fin con avance parcial"].actual_start, t["Fin con avance parcial"].actual_finish = S, F
    t["Fin con avance parcial"].percent_complete = 80
    issues = progress(tasks, None)
    by = {}
    for i in issues:
        by.setdefault(i.name, []).append(i)
    assert by["Fin sin comienzo"][0].severity == SEV_ERROR
    assert by["Al reves"][0].severity == SEV_ERROR and "anterior al Comienzo real" in by["Al reves"][0].message
    assert "40 % y no tiene Comienzo real" in by["Avance sin comienzo"][0].message
    assert "100 % completado y no tiene Fin real" in by["Completa sin fin"][0].message
    assert "80 %" in by["Fin con avance parcial"][0].message
    assert any("no tiene fecha de estado" in i.message and i.severity == SEV_WARN for i in issues)


def test_progress_against_status_date():
    tasks = tree('''
        1 Fase
          1.1 Futuro
          1.2 Atrasada
          1.3 Sin comenzar
          1.4 Terminada
          1.5 Pendiente
    ''')
    t = by_name(tasks)
    status = date(2026, 2, 28)
    t["Futuro"].actual_start, t["Futuro"].percent_complete = datetime(2026, 3, 2, 8), 10
    t["Atrasada"].actual_start, t["Atrasada"].percent_complete = S, 50
    t["Atrasada"].finish = datetime(2026, 2, 20, 17)
    t["Terminada"].actual_start, t["Terminada"].percent_complete = S, 100
    t["Terminada"].actual_finish = datetime(2026, 2, 13, 17)
    for name in ("Atrasada", "Terminada", "Sin comenzar"):
        t[name].baseline_finish = datetime(2026, 2, 13, 17)
    t["Pendiente"].start, t["Pendiente"].finish = datetime(2026, 3, 2, 8), datetime(2026, 3, 27, 17)
    t["Pendiente"].baseline_start = datetime(2026, 3, 2, 8)
    issues = progress(tasks, status)
    future = [i for i in issues if "posterior a la fecha de estado" in i.message]
    assert [(i.name, i.severity) for i in future] == [("Futuro", SEV_ERROR)]
    late = {i.name: i.message for i in issues if i.code and "anterior a la fecha de estado" in i.message}
    assert set(late) == {"Atrasada", "Sin comenzar"}
    assert "no está terminada (50 %)" in late["Atrasada"] and "no tiene avance" in late["Sin comenzar"]
    assert any("2 tareas tienen trabajo pendiente" in i.message for i in issues)
    bei = next(i for i in issues if "BEI" in i.message)
    assert "= 0,33" in bei.message and bei.severity == SEV_WARN and "1.2.0" in bei.message
    started = next(i for i in issues if "debían comenzar" in i.message or "debía comenzar" in i.message)
    assert "1.3.0" in started.message and "1.5.0" not in started.message


# ----------------------------------------------------------------------------
# Lógica del cronograma (solo validar)
# ----------------------------------------------------------------------------
def logic(tasks, cfg=None):
    project = P(tasks)
    part = build_part(project, cfg or Config())
    return logic_checks(project, part.scope, part.code_of)


def test_schedule_without_links():
    issues = logic(tree('''
        1 Fase
          1.1 A
    '''))
    assert len(issues) == 1 and "no tiene vínculos" in issues[0].message and issues[0].topic == TOPIC_LOGIC


def test_links_and_constraints():
    tasks = tree('''
        1 Fase
          1.1 Inicio
          1.2 Suelta
          1.3 Adelanto
          1.4 CF
          1.5 Fin
        2 Resumen
          2.1 Detalle
    ''')
    t = by_name(tasks)
    days = {"Inicio": 0, "Suelta": 10, "Adelanto": 20, "CF": 30, "Fin": 40, "Detalle": 25}
    for name, d in days.items():
        t[name].start = datetime(2026, 1, 5 + d % 20, 8) if d < 20 else datetime(2026, 2, d - 15, 8)
        t[name].finish = t[name].start.replace(hour=17)
    uid = {n: t[n].uid for n in t}
    t["Adelanto"].predecessors = [Link(uid["Inicio"], LINK_FS, -480)]
    t["CF"].predecessors = [Link(uid["Adelanto"], LINK_SF, 0)]
    t["Fin"].predecessors = [Link(uid["CF"], LINK_FS, 960), Link(uid["Detalle"], LINK_SS, 0)]
    t["Resumen"].predecessors = [Link(uid["Inicio"], LINK_FS, 0)]
    t["Detalle"].constraint_type, t["Detalle"].constraint_date = CONSTRAINT_MSO, datetime(2026, 2, 10, 8)
    issues = logic(tasks)
    by = {}
    for i in issues:
        by.setdefault(i.name, []).append(i.message)
    assert by["Suelta"] == ["Sin predecesoras ni sucesoras."]
    assert "Inicio" not in by and "Fin" not in by               # comienzo y fin del proyecto
    assert any("adelanto" in m for m in by["Adelanto"])
    assert any("comienzo a fin (CF)" in m for m in by["CF"])
    assert by["Resumen"] == ["Tarea resumen con vínculos: se recomienda vincular sus tareas de detalle."]
    assert by["Detalle"] == ["Delimitación «Debe comenzar el 10/02/2026»."]   # su predecesora es la de la resumen
    general = [i.message for i in issues if not i.code]
    assert any("1 tarea sin predecesoras y 1 tarea sin sucesoras" in m for m in general)
    assert any("1 vínculo tiene posición positiva" in m for m in general)
    assert any("son fin a comienzo (FC)" in m for m in general)
    assert any("1 tarea tiene delimitaciones" in m for m in general)


# ----------------------------------------------------------------------------
# Lectura de los campos nuevos
# ----------------------------------------------------------------------------
def test_duration_minutes():
    assert duration_minutes("PT40H0M0S") == 2400
    assert duration_minutes("PT7H30M0S") == 450
    assert duration_minutes("P1DT2H") == 1560
    assert duration_minutes("") is None and duration_minutes("x") is None


def test_xml_reads_schedule_fields(tmp_path):
    path = write_xml(tmp_path / "p.xml", project_xml([
        task_xml(1, 1, "A", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)],
                 extra="<Duration>PT480H0M0S</Duration><Estimated>1</Estimated><Manual>1</Manual>"
                       "<ConstraintType>2</ConstraintType><ConstraintDate>2026-01-05T08:00:00</ConstraintDate>"),
        task_xml(2, 1, "B", wbs="2", start=S, finish=F, baselines=[(0, S, F, 0)],
                 extra="<Duration>PT0H0M0S</Duration><PredecessorLink><PredecessorUID>1</PredecessorUID>"
                       "<Type>3</Type><LinkLag>-4800</LinkLag></PredecessorLink>"),
    ], definitions=[(TEXT10, "Text10", "Método EVM"), (NUMBER1, "Number1", "Peso hito")]))
    p = read_xml(path, Config(method_field="Método EVM", milestone_weight_field="Peso hito"))
    a, b = p.tasks
    assert (a.duration, a.estimated, a.manual, a.constraint_type) == (28800, True, True, CONSTRAINT_MSO)
    assert a.constraint_date == datetime(2026, 1, 5, 8) and a.predecessors == []
    assert b.duration == 0 and b.predecessors == [Link(1, LINK_SS, -4800)]
    assert p.fields_found == {"metodo": "Método EVM", "peso_hito": "Peso hito"}


def test_com_reads_schedule_fields():
    from test_reader_com_fake import FakeTask, make_session

    class Dep:
        def __init__(self, a, b, kind=LINK_FS, lag=0):
            self.From, self.To, self.Type, self.Lag = a, b, kind, lag

    a = FakeTask(1, 1, "A", "1")
    b = FakeTask(2, 1, "B", "2")
    a.Duration, a.Estimated, a.Manual, a.ConstraintType = 4800, True, False, 2
    a.ConstraintDate = datetime(2026, 1, 5, 8)
    b.Duration = "1 día?"                                       # no numérico: duración desconocida
    dep = Dep(a, b, LINK_SF, -60)
    a.TaskDependencies = [dep]
    b.TaskDependencies = [dep]
    p = make_session([a, b]).read_project(Config(method_field="Método EVM"))
    ta, tb = p.tasks
    assert (ta.duration, ta.estimated, ta.constraint_type, ta.predecessors) == (4800, True, 2, [])
    assert ta.constraint_date == datetime(2026, 1, 5, 8)
    assert tb.duration is None and tb.predecessors == [Link(1, LINK_SF, -60)]
    assert p.fields_found == {"metodo": "Método EVM"}


# ----------------------------------------------------------------------------
# Comando validar y hoja Validación
# ----------------------------------------------------------------------------
def problem_xml(tmp_path):
    return write_xml(tmp_path / "proyecto.xml", project_xml([
        task_xml(1, 1, "Fase", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)]),
        task_xml(2, 2, "Grupo", wbs="1.2", start=S, finish=F, baselines=[(0, S, F, 0)]),
        task_xml(3, 3, "Mal", wbs="1.14.2", start=S, finish=F, baselines=[(0, S, F, 0)]),
    ], header="<StatusDate>2026-02-27T17:00:00</StatusDate>"))


def test_cli_validar_reports_errors_and_creates_no_file(tmp_path, capsys):
    path = problem_xml(tmp_path)
    before = sorted(tmp_path.iterdir())
    assert main(["validar", "--origen", str(path)]) == 1
    out = capsys.readouterr().out
    assert "Estructura y códigos EDT: 1 error" in out
    assert "[X] 1.14.2 Mal: El código EDT no depende" in out
    assert "Fecha de estado: 27/02/2026" in out
    assert "Lógica del cronograma: 1 advertencia" in out
    assert "No se creó ni se modificó ningún archivo." in out
    assert "El PV y el EV de cada fila" not in out                # recordatorio del libro
    assert sorted(tmp_path.iterdir()) == before


def test_cli_validar_without_errors_and_limit(tmp_path, capsys):
    tasks = [task_xml(1, 1, "Fase", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)])]
    tasks += [task_xml(i, 2, f"T{i}", wbs=f"1.{i}", start=S, finish=F, baselines=[(0, S, F, 0)],
                       extra="<Estimated>1</Estimated>") for i in range(2, 30)]
    path = write_xml(tmp_path / "p.xml", project_xml(tasks))
    assert main(["validar", "--origen", str(path), "--fecha-corte", "2026-01-31"]) == 0
    out = capsys.readouterr().out
    assert "Fecha de estado: 31/01/2026 (--fecha-corte)" in out
    assert "Fechas y duración: 29 advertencias" in out and "... y 9 más (--todas para verlas todas)." in out
    main(["validar", "--origen", str(path), "--todas"])
    assert "--todas para verlas todas" not in capsys.readouterr().out


def test_cli_validar_missing_plan_reviews_whole_project(tmp_path, capsys):
    path = problem_xml(tmp_path)
    cfg = tmp_path / "config.toml"
    cfg.write_text('[exportacion]\nplanificacion = "INTERNO"\n', encoding="utf-8")
    assert main(["validar", "--origen", str(path), "--config", str(cfg)]) == 1
    out = capsys.readouterr().out
    assert "No se encontró la planificación 'INTERNO'" in out and "Se revisa el proyecto completo." in out
    assert "1.14.2 Mal" in out


def test_cli_validar_several_files_with_repeated_codes(tmp_path, capsys):
    one = problem_xml(tmp_path)
    two = write_xml(tmp_path / "dos.xml", project_xml([
        task_xml(1, 1, "Otra fase", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0)])]))
    cfg = tmp_path / "config.toml"
    cfg.write_text(f"[[proyectos]]\nnombre = \"P1\"\norigen = '{one}'\n\n"
                   f"[[proyectos]]\nnombre = \"P2\"\norigen = '{two}'\n", encoding="utf-8")
    assert main(["validar", "--config", str(cfg)]) == 1
    out = capsys.readouterr().out
    assert "El código EDT '1.0.0' está en los archivos 'P1' y 'P2'" in out
    assert "[P1] El código EDT no depende" in out


def test_part_keeps_scope_for_schedule_checks():
    tasks = tree('''
        1 Fase
          1.1 A
    ''')
    part = validate_part(P(tasks), Config(), "", date(2026, 2, 1))
    assert part.scope is not None and part.code_of(tasks[1]) == "1.1.0"
    assert any(i.topic == TOPIC_PROGRESS for i in part.checks) and any(i.topic == TOPIC_LOGIC for i in part.checks)


def test_generar_writes_topic_and_new_checks(tmp_path):
    path = problem_xml(tmp_path)
    out = tmp_path / "libro.xlsx"
    assert main(["generar", "--origen", str(path), "--salida", str(out)]) == 0
    ws = load_workbook(out)["Validación"]
    assert ws.cell(row=3, column=5).value == "Tema"
    rows = [r for r in ws.iter_rows(min_row=4, values_only=True) if r[0]]
    assert ("Error", "1.14.2", "Mal") == rows[[r[1] for r in rows].index("1.14.2")][:3]
    assert {r[4] for r in rows} >= {TOPIC_WBS}
    assert not [r for r in rows if r[4] in (TOPIC_PROGRESS, TOPIC_LOGIC)]     # solo en validar
