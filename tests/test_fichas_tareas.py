"""Fichas de tareas (comando generar-fichas): elementos exportados, fechas y duración de la línea base, vínculos
trasladados al elemento con ficha, documento Markdown y lectura de los campos nuevos."""
from datetime import date, datetime

from evm.cli import main
from evm.config import Config, load_config
from evm.model import Link, use_current_schedule
from evm.reader_xml import read_xml
from evm.task_cards import build_cards, render
from evm.wbs import assemble_report, build_part

from builders import P, by_name, tree
from mspdi import TEXT2, project_xml, task_xml, write_xml

S, F = datetime(2026, 1, 5, 8), datetime(2026, 3, 27, 17)        # 12 semanas: 60 días laborables
TODAY = date(2026, 10, 24)


def cards_for(tasks, cfg=None, project=None, **project_kw):
    cfg = cfg or Config()
    cfg.validate()
    part = build_part(project or P(tasks, **project_kw), cfg)
    return build_cards(assemble_report([part], cfg), [part], cfg)


def by_code(cards):
    return {c.code: c for c in cards.cards}


def codes(cards):
    return [c.code for c in cards]


def link(task, *predecessors):
    task.predecessors = [Link(p.uid) for p in predecessors]


# ----------------------------------------------------------------------------
# Elementos y vínculos
# ----------------------------------------------------------------------------
def test_links_are_rolled_up_to_the_exported_elements():
    tasks = tree('''
        1 Fase
          1.1 Diseño
            1.1.1 Requisitos
              - Detalle a
              - Detalle b
            1.1.2 Arquitectura
          1.2 Construcción
            1.2.1 Obra
            - Suelta
        2 Cierre
    ''')
    t = by_name(tasks)
    link(t["Detalle b"], t["Detalle a"])                        # interno de 1.1.1: no se muestra
    link(t["Arquitectura"], t["Detalle b"])
    link(t["Obra"], t["Detalle a"], t["Requisitos"])            # el mismo par de elementos: una vez
    link(t["Cierre"], t["Suelta"])                              # tarea sin código: es de 1.2.0
    cards = cards_for(tasks)
    c = by_code(cards)
    assert list(c) == ["1.0.0", "1.1.0", "1.1.1", "1.1.2", "1.2.0", "1.2.1", "2.0.0"]     # sin la fila 1.2.0-D
    assert codes(c["1.1.1"].predecessors) == [] and codes(c["1.1.1"].successors) == ["1.1.2", "1.2.1"]
    assert codes(c["1.1.2"].predecessors) == ["1.1.1"] and codes(c["1.2.1"].predecessors) == ["1.1.1"]
    assert codes(c["1.2.0"].successors) == ["2.0.0"] and codes(c["2.0.0"].predecessors) == ["1.2.0"]
    assert not c["1.0.0"].predecessors and not c["1.1.0"].successors     # resúmenes: solo sus vínculos propios
    assert (cards.links, cards.hidden) == (3, 0)


def test_links_with_tasks_without_card_are_not_shown():
    tasks = tree('''
        INTERNO Interno
          1 Fase
            1.1 Inicio
            1.2 Inactiva
            1.3 Oculta
            1.4 Final
        CLIENTE Cliente
          1 Fase del cliente
            1.1 Tarea del cliente
            1.2 Otra tarea del cliente
    ''')
    t = by_name(tasks)
    t["Inactiva"].active = False
    t["Oculta"].visible = "No"
    link(t["Inicio"], t["Tarea del cliente"])
    link(t["Final"], t["Inactiva"], t["Inicio"])
    link(t["Oculta"], t["Inicio"])
    link(t["Otra tarea del cliente"], t["Tarea del cliente"])    # fuera de las fichas en los dos extremos
    cards = cards_for(tasks, Config(plan="INTERNO"))
    c = by_code(cards)
    assert list(c) == ["1.0.0", "1.1.0", "1.4.0"]
    assert codes(c["1.1.0"].predecessors) == [] and codes(c["1.1.0"].successors) == ["1.4.0"]
    assert (cards.links, cards.hidden) == (1, 3)
    text = render(cards, Config(plan="INTERNO"), TODAY)
    assert "cliente" not in text and "Oculta" not in text and "Inactiva" not in text


def test_dates_and_working_days():
    tasks = tree('''
        1 Fase
          1.1 Base
          1.2 Hito
          1.3 Nueva
          1.4 Sin duración
          1.5 Medio día
    ''')
    t = by_name(tasks)
    t["Base"].baseline_duration, t["Base"].duration = 4800, 9600         # se usa la de la línea base
    t["Hito"].milestone, t["Hito"].baseline_start, t["Hito"].baseline_duration = True, F, 0
    nueva = t["Nueva"]
    nueva.baseline_start = nueva.baseline_finish = None
    nueva.has_baseline = False
    nueva.start, nueva.finish, nueva.duration = datetime(2026, 4, 6, 8), datetime(2026, 4, 10, 17), 2400
    t["Medio día"].baseline_duration = 240
    c = by_code(cards_for(tasks))
    assert (c["1.1.0"].start, c["1.1.0"].finish, c["1.1.0"].days) == (S, F, 10)
    assert c["1.2.0"].days == 0 and c["1.2.0"].milestone
    assert (c["1.3.0"].start, c["1.3.0"].days, c["1.3.0"].current) == (datetime(2026, 4, 6, 8), 5, True)
    assert c["1.4.0"].days == 60 and c["1.0.0"].days == 60        # sin duración en el archivo: calendario
    assert c["1.5.0"].days == 0.5
    assert not any(x.current for code, x in c.items() if code != "1.3.0")
    text = render(cards_for(tasks), Config(), TODAY)
    assert "- Duración: 10 días" in text and "- Duración: 0 días (hito)" in text and "- Duración: 0,5 días" in text
    assert text.count("- Observación: la tarea no tiene línea base") == 1


def test_minutes_per_day_of_the_project():
    tasks = tree('''
        1 Fase
    ''', baseline_duration=4200)
    assert by_code(cards_for(tasks, minutes_per_day=420))["1.0.0"].days == 10
    tasks[0].baseline_duration = 420
    assert "- Duración: 1 día\n" in render(cards_for(tasks, minutes_per_day=420), Config(), TODAY)


def test_project_without_baseline_uses_current_schedule():
    tasks = tree('''
        1 Fase
          1.1 A
    ''', bs=None, bf=None, duration=4800)
    project = P(tasks)
    use_current_schedule(project)
    cards = cards_for(tasks, project=project)
    c = by_code(cards)
    assert c["1.1.0"].days == 10 and not cards.current
    text = render(cards, Config(), TODAY)
    assert "- Fechas planificadas: Programa actual (sin línea base)" in text and "Observación" not in text


# ----------------------------------------------------------------------------
# Documento Markdown
# ----------------------------------------------------------------------------
def test_markdown_document():
    tasks = tree('''
        1 Fase *crítica* #1
          1.1 Diseño [v2] | revisión
            1.1.1 Planos_finales
              1.1.1.1 Detalle visible
            1.1.2 Cálculo
    ''')
    t = by_name(tasks)
    t["Detalle visible"].visible = "Sí"
    t["Planos_finales"].description = "Primera línea\r\n- viñeta\r\n\r\n1. paso\r<b>html</b>"
    link(t["Cálculo"], t["Detalle visible"])
    text = render(cards_for(tasks), Config(), TODAY)
    lines = text.splitlines()
    assert lines[:8] == ["# Fichas de tareas: Programa", "", "- Fechas planificadas: Línea base",
                         "- Duración: en días laborables", "- Nivel de desglose EDT: 3", "- Fichas: 5",
                         "- Fecha de emisión: 24/10/2026", ""]
    assert "## 1.0.0 Fase \\*crítica\\* \\#1" in lines
    assert "### 1.1.0 Diseño \\[v2\\] \\| revisión" in lines
    assert "#### 1.1.1 Planos\\_finales" in lines
    assert "##### 1.1.1.1 Detalle visible" in lines               # Visible = Sí: por debajo del nivel máximo
    start = lines.index("#### 1.1.1 Planos\\_finales")
    assert lines[start + 1:start + 10] == ["", "Primera línea", "", "\\- viñeta", "", "1\\. paso", "",
                                           "\\<b\\>html\\</b\\>", ""]
    assert lines[start + 10:start + 15] == ["- Inicio planificado: 05/01/2026", "- Fin planificado: 27/03/2026",
                                            "- Duración: 60 días", "- Predecesoras: ninguna", "- Sucesoras: ninguna"]
    block = text.split("##### 1.1.1.1 Detalle visible", 1)[1]
    assert "- Sucesoras:\n    - 1.1.2 - Cálculo\n" in block
    assert text.endswith("\n")


def test_plan_in_header():
    tasks = tree('''
        INTERNO Interno
          1 Fase
    ''')
    text = render(cards_for(tasks, Config(plan="INTERNO")), Config(plan="INTERNO"), TODAY)
    assert "- Planificación: INTERNO" in text.splitlines()


# ----------------------------------------------------------------------------
# Lectura de los campos nuevos
# ----------------------------------------------------------------------------
def fields_xml(tmp_path):
    return write_xml(tmp_path / "p.xml", project_xml([
        task_xml(1, 1, "A", wbs="1", start=S, finish=F, baselines=[(0, S, F, 0, "PT80H0M0S")],
                 extra="<Duration>PT160H0M0S</Duration><Notes>Alcance de A</Notes>", ext={TEXT2: "Texto propio"}),
        task_xml(2, 1, "B", wbs="2", start=S, finish=F),
    ], header="<MinutesPerDay>420</MinutesPerDay>", definitions=[(TEXT2, "Text2", "Descripción")]))


def test_xml_reads_baseline_duration_minutes_per_day_and_description(tmp_path):
    path = fields_xml(tmp_path)
    p = read_xml(path, Config())
    a, b = p.tasks
    assert (a.baseline_duration, a.duration, b.baseline_duration, p.minutes_per_day) == (4800, 9600, None, 420)
    assert a.description == "" and not p.warnings                  # por defecto no se lee la descripción
    assert read_xml(path, Config(description_field="Notas")).tasks[0].description == "Alcance de A"
    assert read_xml(path, Config(description_field="Descripción")).tasks[0].description == "Texto propio"
    p = read_xml(path, Config(description_field="Resumen"))
    assert p.tasks[0].description == "" and "campo de descripción 'Resumen'" in p.warnings[0]


def test_com_reads_baseline_duration_hours_per_day_and_notes():
    from test_reader_com_fake import FakeTask, make_session
    a = FakeTask(1, 1, "A", "1")
    b = FakeTask(2, 1, "B", "2", bl=False)
    a.BaselineDuration, a.Notes = 4800, "Nota\rsegunda línea"
    b.BaselineDuration = 0                                        # MS Project da 0 sin línea base
    s = make_session([a, b])
    s.project.HoursPerDay = 7.5
    p = s.read_project(Config(description_field="Notas"))
    ta, tb = p.tasks
    assert (ta.baseline_duration, tb.baseline_duration, p.minutes_per_day) == (4800, None, 450)
    assert ta.description == "Nota\rsegunda línea"
    assert make_session([FakeTask(1, 1, "A", "1")]).read_project(Config()).minutes_per_day == 480


def test_config_description_field(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text('[campos]\ndescripcion = "Notas"\n', encoding="utf-8")
    assert load_config(path).description_field == "Notas" and Config().description_field == ""


# ----------------------------------------------------------------------------
# Comando generar-fichas
# ----------------------------------------------------------------------------
def project_file(path, title, codes_and_names, link_to_last=True):
    """XML con un elemento de nivel 1 y sus tareas de nivel 2; la última depende de la anterior."""
    (top_code, top_name), *rest = codes_and_names
    tasks = [task_xml(1, 1, top_name, wbs=top_code, start=S, finish=F, baselines=[(0, S, F, 0, "PT480H0M0S")])]
    for i, (code, name) in enumerate(rest, start=2):
        extra = (f"<PredecessorLink><PredecessorUID>{i - 1}</PredecessorUID></PredecessorLink>"
                 if link_to_last and i == len(rest) + 1 and i > 2 else "")
        tasks.append(task_xml(i, 2, name, wbs=code, start=S, finish=F, baselines=[(0, S, F, 0, "PT80H0M0S")],
                              extra=extra))
    return write_xml(path, project_xml(tasks, title=title))


def test_cli_generar_fichas(tmp_path, capsys):
    path = project_file(tmp_path / "proyecto.xml", "Programa", [("1", "Fase"), ("1.1", "Diseño"), ("1.2", "Obra")])
    assert main(["generar-fichas", "--origen", str(path)]) == 0
    out = capsys.readouterr().out
    assert "Fichas: 3  |  Nivel máximo: 3  |  Vínculos entre fichas: 1" in out
    text = (tmp_path / "proyecto_fichas.md").read_text(encoding="utf-8")
    assert text.startswith("# Fichas de tareas: Programa\n")
    obra = text.split("### 1.2.0 Obra", 1)[1]
    assert "- Duración: 10 días\n- Predecesoras:\n    - 1.1.0 - Diseño\n- Sucesoras: ninguna" in obra
    assert main(["generar-fichas", "--origen", str(path)]) == 2
    assert "ya existe" in capsys.readouterr().err
    assert main(["generar-fichas", "--origen", str(path), "--sobrescribir"]) == 0
    assert main(["generar-fichas", "--origen", str(path), "--salida", str(tmp_path / "fichas.txt")]) == 2
    assert ".md" in capsys.readouterr().err


def test_cli_generar_fichas_several_files(tmp_path, capsys):
    # Los dos archivos usan los mismos UniqueID: los vínculos se resuelven en cada archivo
    one = project_file(tmp_path / "uno.xml", "Proyecto uno", [("1", "Fase uno"), ("1.1", "Diseño"), ("1.2", "Obra")])
    two = project_file(tmp_path / "dos.xml", "Proyecto nueve",
                       [("9", "Fase nueve"), ("9.1", "Montaje"), ("9.2", "Pruebas")])
    cfg = tmp_path / "config.toml"
    cfg.write_text(f"[[proyectos]]\nnombre = \"P1\"\norigen = '{one}'\n\n"
                   f"[[proyectos]]\nnombre = \"P2\"\norigen = '{two}'\n", encoding="utf-8")
    assert main(["generar-fichas", "--config", str(cfg)]) == 2
    assert "--salida" in capsys.readouterr().err
    salida = tmp_path / "fichas.md"
    assert main(["generar-fichas", "--config", str(cfg), "--salida", str(salida)]) == 0
    out = capsys.readouterr().out
    assert "P1: Proyecto uno" in out and "P2: Proyecto nueve" in out and "Vínculos entre fichas: 2" in out
    text = salida.read_text(encoding="utf-8")
    assert "- Archivos de MS Project:\n    - Proyecto uno: fechas planificadas: Línea base\n" in text
    assert "### 1.2.0 Obra" in text and "### 9.2.0 Pruebas" in text
    assert "    - 1.1.0 - Diseño" in text.split("### 1.2.0 Obra", 1)[1].split("##", 1)[0]
    assert "    - 9.1.0 - Montaje" in text.split("### 9.2.0 Pruebas", 1)[1]
    assert "Montaje" not in text.split("### 1.2.0 Obra", 1)[1].split("##", 1)[0]


def test_cli_generar_fichas_missing_plan(tmp_path, capsys):
    path = project_file(tmp_path / "proyecto.xml", "Programa", [("1", "Fase"), ("1.1", "Diseño")])
    assert main(["generar-fichas", "--origen", str(path), "--planificacion", "CLIENTE"]) == 2
    assert "No se encontró la planificación 'CLIENTE'" in capsys.readouterr().err
    assert not (tmp_path / "proyecto_fichas.md").exists()
