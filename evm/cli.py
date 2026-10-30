"""Interfaz de línea de comandos.

    python evm_export.py generar    --origen proyecto.mpp --salida EVM.xlsx
    python evm_export.py actualizar --origen proyecto.mpp --libro EVM.xlsx --fecha-corte 2026-10-31
    python evm_export.py validar    --origen proyecto.mpp
    python evm_export.py analizar   --libro EVM.xlsx
    python evm_export.py generar-fichas --origen proyecto.mpp --salida Fichas.md
"""
from __future__ import annotations

import argparse
import logging
import sys
from contextlib import contextmanager
from dataclasses import replace
from datetime import date
from pathlib import Path

from evm import __version__, progress
from evm.config import Config, ConfigError, load_config, parse_date, parse_month
from evm.fields import fmt_number
from evm.messages import (SEV_ERROR, SEV_INFO, SEV_WARN, SUMMARY_ID, EvmError, Issue, MissingFileError, Silence,
                          parse_silence)
from evm.wbs import assemble_report, build_part, fmt_money, plural

log = logging.getLogger("evm")

COM_EXTENSIONS = {".mpp", ".mpt", ".mpd"}


def _engine_for(origen: str, motor: str) -> str:
    if motor != "auto":
        return motor
    if origen.lower() == "activo":
        return "com"
    ext = Path(origen).suffix.lower()
    if ext == ".xml":
        return "xml"
    if ext in COM_EXTENSIONS:
        return "com"
    raise ConfigError("CFG-035",
                      f"No se reconoce la extensión '{ext}'. Se debe usar un archivo .mpp (requiere MS Project) "
                      "o .xml (MS Project > Guardar como > XML), o indicar --motor.")


@contextmanager
def open_project(origen: str, cfg: Config, motor: str, visible: bool = False, status: date | None = None,
                 name: str = "", links: bool = True):
    """Devuelve el proyecto leído (ProjectData). Con COM, MS Project queda abierto mientras dure el bloque.

    name: nombre del archivo en [[proyectos]] (etapas de la barra de progreso).
    links: lee los vínculos y las delimitaciones de las tareas (validar y generar-fichas).
    """
    engine = _engine_for(origen, motor)
    if origen.lower() != "activo" and not Path(origen).is_file():
        raise MissingFileError("LEC-001", f"No existe el archivo de origen: {origen}")
    which = f" ({name})" if name else ""
    if engine == "xml":
        from evm.reader_xml import read_xml
        progress.stage(f"leer:{name}", f"Lectura del XML{which}")
        yield read_xml(origen, cfg, links=links)
        return
    from evm.reader_com import ComSession
    progress.stage(f"abrir:{name}", f"Apertura en MS Project{which}")
    with ComSession(origen, visible=visible) as session:
        progress.stage(f"leer:{name}", f"Lectura de MS Project{which}")
        yield session.read_project(cfg, status_date=status, links=links)


def _stages(files: list[tuple[str, str, Config]], motor: str, before: dict[str, float] | None = None,
            compute: float = 0.0, after: dict[str, float] | None = None) -> dict[str, float]:
    """Etapas de la barra de progreso y su peso relativo (tiempo habitual de cada una): por archivo, apertura en
    MS Project (COM), lectura y cálculo; antes y después, las del comando."""
    stages = dict(before or {})
    for name, origin, _cfg in files:
        if _engine_for(origin, motor) == "com":
            stages[f"abrir:{name}"] = 15.0
            stages[f"leer:{name}"] = 40.0
        else:
            stages[f"leer:{name}"] = 10.0
        if compute:
            stages[f"calcular:{name}"] = compute
    stages.update(after or {})
    return stages


def _computing(name: str, text: str) -> None:
    progress.stage(f"calcular:{name}", f"{text} ({name})" if name else text)


def _apply_overrides(cfg: Config, args) -> Config:
    # Opciones indicadas explícitamente (aunque valgan 0 o ""): se validan, no se ignoran
    if getattr(args, "mes_inicio", None) is not None:
        cfg.start_month = parse_month(args.mes_inicio)
    if getattr(args, "meses", None) is not None:
        cfg.months = args.meses
    if getattr(args, "linea_base", None) is not None:
        cfg.baseline = args.linea_base
    if getattr(args, "nivel_max", None) is not None:
        cfg.max_level = args.nivel_max
    if getattr(args, "planificacion", None) is not None:
        cfg.plan = args.planificacion.strip()
    _add_silence(cfg, args)
    cfg.validate()
    return cfg


def _add_silence(cfg: Config, args) -> None:
    """--silenciar (se puede repetir y separar con comas) se suma a [validacion] silenciar."""
    for item in getattr(args, "silenciar", None) or []:
        cfg.silence += [x.strip() for x in item.split(",") if x.strip()]


def _silence(cfg: Config) -> Silence:
    return parse_silence(cfg.silence)


def _sources(cfg: Config, args) -> list[tuple[str, str, Config]]:
    """Archivos de MS Project a leer: [(nombre, ruta, configuración del archivo)].

    Un solo archivo: [("", --origen, cfg)]. Varios archivos ([[proyectos]]): uno por sección, en su orden;
    --origen NOMBRE=RUTA reemplaza la ruta de un archivo.
    """
    given = args.origen or []
    if not cfg.projects:
        if not given:
            raise ConfigError("CFG-036",
                              "Se debe indicar el archivo de MS Project con --origen (o los archivos en secciones "
                              "[[proyectos]] del archivo de configuración).")
        if len(given) > 1:
            raise ConfigError("CFG-037",
                              "Para combinar varios archivos de MS Project en un mismo libro, se deben indicar en "
                              "secciones [[proyectos]] del archivo de configuración (--origen se indica una vez).")
        return [("", given[0], cfg)]
    if args.planificacion is not None or args.linea_base is not None:
        raise ConfigError("CFG-038",
                          "Con varios archivos ([[proyectos]]), la planificación y la línea base se indican en la "
                          "sección [[proyectos]] de cada archivo, no con --planificacion ni --linea-base.")
    origins = {p.name: p.origin for p in cfg.projects}
    for item in given:
        name, sep, path = (x.strip() for x in item.partition("="))
        if not sep or name not in origins or not path:
            raise ConfigError("CFG-039",
                              f"--origen '{item}': con varios archivos ([[proyectos]]) se indica NOMBRE=RUTA para "
                              f"reemplazar la ruta de un archivo, con uno de estos nombres: {', '.join(origins)}.")
        if path.lower() == "activo":
            raise ConfigError("CFG-031",
                              "Con varios archivos ([[proyectos]]) no se admite 'activo': se debe indicar la ruta del "
                              "archivo .mpp o .xml.")
        origins[name] = path
    return [(p.name, origins[p.name], cfg.project_config(p.name)) for p in cfg.projects]


def _reading(name: str, origin: str) -> str:
    return f"Leyendo {name}: {origin} ..." if name else f"Leyendo {origin} ..."


MARKS = {SEV_ERROR: "[X]", SEV_WARN: "[!]", SEV_INFO: "[i]"}


def _issue_line(i: Issue) -> str:
    """'  [!] FEC-003 1.2.3 Montaje: mensaje' (el código de la observación, la fila o tarea y el mensaje)."""
    where = f"{i.code} {i.name}".strip()
    return f"  {MARKS.get(i.severity, '[!]')} {i.id} {where + ': ' if where else ''}{i.message}"


def _print_issues(issues: list[Issue], limit: int = 15) -> None:
    serious = [i for i in issues if i.severity != SEV_INFO]
    infos = [i for i in issues if i.severity == SEV_INFO]
    for i in infos[:limit]:
        print(_issue_line(i))
    if len(infos) > limit:
        print(f"  ... y {len(infos) - limit} informativas más (ver hoja 'Validación').")
    serious.sort(key=lambda i: i.severity != SEV_ERROR)
    for i in serious[:limit]:
        print(_issue_line(i))
    if len(serious) > limit:
        print(f"  ... y {len(serious) - limit} observaciones más (ver hoja 'Validación').")


def cmd_generar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    files = _sources(cfg, args)
    origen = files[0][1]
    if args.salida:
        salida = Path(args.salida)
    elif cfg.projects:
        raise ConfigError("CFG-040",
                          "Con varios archivos ([[proyectos]]) se debe indicar el libro a crear con --salida.")
    elif origen.lower() == "activo":
        salida = Path("EVM.xlsx")
    else:
        salida = Path(origen).with_name(Path(origen).stem + "_EVM.xlsx")
    if salida.suffix.lower() != ".xlsx":
        raise ConfigError("CFG-041", "La salida debe ser un archivo .xlsx")
    _check_folder(salida)
    if salida.exists() and not args.sobrescribir:
        raise ConfigError("CFG-042",
                          f"'{salida}' ya existe. Para reemplazarlo se debe usar --sobrescribir; para conservar "
                          "los EV/AC cargados se debe generar un archivo nuevo con --migrar-desde.")
    if args.migrar_desde and Path(args.migrar_desde).resolve() == salida.resolve():
        raise ConfigError("CFG-043", "--migrar-desde debe ser un archivo distinto de --salida.")

    from evm.workbook import check_writable, write_workbook
    if salida.exists():
        check_writable(salida)                  # libro abierto en Excel: se informa antes de leer MS Project
    parts = []
    after = {"migrar": 10.0} if args.migrar_desde else {}
    with progress.tracking(_stages(files, args.motor, compute=10.0,
                                   after={**after, "escribir": 20.0, "guardar": 15.0})):
        for name, origin, pcfg in files:
            progress.say(_reading(name, origin))
            with open_project(origin, pcfg, args.motor, args.visible, name=name, links=False) as project:
                _computing(name, "Cálculo del % planificado")
                parts.append(build_part(project, pcfg, name))
        report = assemble_report(parts, cfg)
        prior = None
        if args.migrar_desde:
            from evm.migrate import plan_migration, prior_from_book
            progress.stage("migrar", "Lectura del libro anterior")
            prior = prior_from_book(args.migrar_desde, report.project.name)
            report.issues.extend(plan_migration(report, prior, cfg.max_level))
        silence = _silence(cfg)
        unmatched = []
        if prior is not None and prior.unmatched:
            lost_ev = sum(r.total("EV") for r in prior.unmatched)
            lost_ac = sum(r.total("AC") for r in prior.unmatched)
            unmatched = silence.filter([Issue("MIG-017", message=f"{len(prior.unmatched)} filas no se pudieron "
                                                                 f"trasladar (EV {fmt_money(lost_ev)}, AC "
                                                                 f"{fmt_money(lost_ac)}). Ver hoja 'Validación'.")])
        report.issues = silence.apply(report.issues)
        write_workbook(report, cfg, salida, prior)

    leaves = sum(1 for r in report.rows if r.is_input)
    if report.multi_file:
        print(f"Informe: {report.title}")
        for s in report.sources:
            plan = f"  |  Planificación: {s.plan}" if s.plan else ""
            print(f"  {s.name}: {s.project.name}{plan}  |  {s.project.baseline_label}  |  PV: {s.pv_source}")
    else:
        project = report.project
        print(f"Proyecto: {project.name}")
        if report.plan:
            print(f"  Planificación: {report.plan}")
        print(f"  {project.baseline_label}  |  PV: {report.pv_source}")
    print(f"  Filas EDT: {len(report.rows)} ({leaves} de ingreso)  |  {_rows_rule(cfg)}")
    print(f"  Horizonte: {report.horizon.start:%m/%Y} a {report.horizon.end:%m/%Y} ({report.horizon.months} meses)")
    sheets = (f"PV-{report.company_a} y PV-{report.company_b}" if report.multi else "PV")
    print(f"  BAC: se debe ingresar el BAC de cada fila en la columna D de {sheets} (celdas amarillas).")
    _print_issues(report.issues)
    if prior is not None:
        print(f"Migración desde {args.migrar_desde}: {prior.matched} filas con historial trasladado.")
        for i in unmatched:
            print(_issue_line(i))
    print(f"Libro generado: {salida.resolve()}")
    return 0


def _check_folder(salida: Path) -> None:
    """La carpeta de --salida debe existir: se informa antes de leer MS Project, no al guardar."""
    folder = salida.resolve().parent
    if not folder.is_dir():
        raise ConfigError("CFG-051", f"No existe la carpeta del archivo de salida: {folder}. Se debe crear la "
                                     "carpeta o indicar otra ruta en --salida.")


def _rows_rule(cfg) -> str:
    """Cómo se deciden las filas del libro: el nivel máximo o la lista EDT."""
    if cfg.wbs_list is not None:
        return f"Lista EDT: {plural(len(cfg.wbs_list), 'código', 'códigos')} ({Path(cfg.wbs_list_file).name})"
    return f"Nivel máximo: {cfg.max_level}"


def _print_schedule(plan) -> None:
    """Plazo según el cronograma de MS Project (hojas Resumen e Indicadores EDT)."""
    t = plan.total
    if t.finish is None:
        return

    def day(d) -> str:
        return f"{d:%d/%m/%Y}" if d else "-"

    variance = ""
    if t.variance is not None:
        sign = "+" if t.variance > 0 else ""
        variance = f" (desvío {sign}{t.variance} {'día laborable' if abs(t.variance) == 1 else 'días laborables'})"
    print(f"Plazo según MS Project: fin de línea base {day(t.baseline_finish)}  |  fin previsto {day(t.finish)}"
          f"{variance}")
    print(f"  Paquetes críticos sin terminar: {plan.critical} ({plan.late_critical} con fin atrasado)")


def _print_update(res, currency: str, silence: Silence) -> None:
    from evm.indicators import format_table
    from evm.update import indicators_for
    print(f"Fecha de corte: {res.status_date:%d/%m/%Y} (mes {res.month_index + 1} del horizonte)")
    if res.dry_run:
        def m(v):
            return "-" if v is None else fmt_number(v, 0)

        company_cols = sorted({k for r in res.rows for k in r.by_company})
        print(f"{'Código':<10} {'EV del mes':>16} {'EV acumulado':>16}" + "".join(f" {k:>16}" for k in company_cols))
        for r in res.rows:
            print(f"{r.code:<10} {m(r.ev_month):>16} {m(r.ev_cum):>16}"
                  + "".join(f" {m(r.by_company.get(k)):>16}" for k in company_cols))
    notes = silence.filter([n for r in res.rows for n in r.notes])
    if notes:
        print("Observaciones por fila:")
        for n in notes[:25]:
            print(_issue_line(replace(n, name=n.name[:40])))
        if len(notes) > 25:
            print(f"  ... y {len(notes) - 25} más")
    reminder = []
    if res.book is not None and res.book.multi:
        reminder.append(Issue("ACT-026", message=f"El AC se carga a mano en las hojas 'AC-{res.book.company_a}' y "
                                                 f"'AC-{res.book.company_b}'; los indicadores de costo usan los "
                                                 "valores cargados."))
    elif res.book is not None:
        reminder.append(Issue("ACT-026", message="El AC se carga a mano en la hoja 'AC'; los indicadores de costo "
                                                 "usan los valores cargados."))
    _print_notices(silence.apply(res.warnings + reminder))
    _print_deviations(res)
    if res.book is not None:
        print()
        print(format_table(indicators_for(res.book, res.status_date), currency))
    if res.schedule is not None:
        print()
        _print_schedule(res.schedule)


def _print_notices(issues: list[Issue], limit: int = 30) -> None:
    """Advertencias e informativas de actualizar, analizar y generar-fichas (el resumen de las silenciadas, siempre)."""
    shown = [i for i in issues if i.id != SUMMARY_ID]
    for i in shown[:limit]:
        print(_issue_line(i))
    if len(shown) > limit:
        print(f"  ... y {len(shown) - limit} observaciones más.")
    for i in issues:
        if i.id == SUMMARY_ID:
            print(_issue_line(i))


def _print_deviations(res, limit: int = 25) -> None:
    """Paquetes con SPI o CPI acumulado fuera del rango verde del semáforo (hoja Análisis de desviaciones)."""
    th = res.thresholds
    if th is None:
        return
    from evm.workbook import fmt_decimal
    rule = f"SPI o CPI fuera de {fmt_decimal(th.yellow_low)} a {fmt_decimal(th.yellow_high)}"
    found = res.deviations
    if not found:
        print(f"Paquetes fuera de umbral ({rule}): ninguno.")
        return

    def r(v):
        return "-" if v is None else fmt_number(v)

    print(f"Paquetes fuera de umbral ({rule}): {len(found)}"
          + ("" if res.dry_run else " (hoja 'Análisis de desviaciones')"))
    for d in found[:limit]:
        print(f"  {d.code:<10} {d.name[:40]:<40}  SPI {r(d.spi):>5}  CPI {r(d.cpi):>5}")
    if len(found) > limit:
        print(f"  ... y {len(found) - limit} más")


def cmd_analizar(args) -> int:
    cfg = load_config(args.config)
    _add_silence(cfg, args)
    from evm.update import book_meta, run_analysis
    from evm.workbook import check_writable
    book_meta(args.libro)                   # libro inexistente o que no es de este programa: error con su código
    check_writable(args.libro)
    with progress.tracking({"libro": 40.0, "analizar": 10.0, "guardar": 40.0}):
        res = run_analysis(cfg, args.libro, backup=not args.sin_respaldo)
    print(f"Fecha de corte: {res.status_date:%d/%m/%Y} (mes {res.month_index + 1} del horizonte)")
    _print_notices(_silence(cfg).apply(res.warnings))
    _print_deviations(res)
    if res.backup:
        print(f"Respaldo: {res.backup}")
    print(f"Libro actualizado: {Path(args.libro).resolve()}")
    return 0


def cmd_actualizar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    status = parse_date(args.fecha_corte) if args.fecha_corte else None
    from evm.update import apply_book_settings, book_meta, check_sources, run_update
    # Multi-empresa: se usan los nombres y campos con los que se generó el libro
    meta = book_meta(args.libro)
    cfg = apply_book_settings(cfg, meta)
    files = _sources(cfg, args)
    check_sources(meta, [name for name, _origin, _cfg in files])
    if not args.simular:
        from evm.workbook import check_writable
        check_writable(args.libro)              # libro abierto en Excel: se informa antes de leer MS Project
    projects = {}
    after = {"libro": 20.0, "calcular": 10.0} | ({} if args.simular else {"guardar": 20.0})
    with progress.tracking(_stages(files, args.motor, after=after)):
        for name, origin, pcfg in files:
            progress.say(_reading(name, origin))
            with open_project(origin, pcfg, args.motor, args.visible, status, name=name, links=False) as project:
                projects[name] = project
        res = run_update(projects if cfg.projects else projects[""], cfg, args.libro, status,
                         backup=not args.sin_respaldo, force=args.forzar, dry_run=args.simular)
    _print_update(res, str(res.book.meta.get("moneda") or "") if res.book else "", _silence(cfg))
    if args.simular:
        print("Simulación: no se modificó el libro.")
    else:
        if res.backup:
            print(f"Respaldo: {res.backup}")
        print(f"Libro actualizado: {Path(args.libro).resolve()}")
        print("Los indicadores y los gráficos se calculan en el libro (hojas 'Cuadro de Mando', 'Resumen' e "
              "'Indicadores EDT').")
    return 0


def _count_text(issues) -> str:
    counts = [(sum(1 for i in issues if i.severity == sev), one, many)
              for sev, one, many in ((SEV_ERROR, "error", "errores"), (SEV_WARN, "advertencia", "advertencias"),
                                     (SEV_INFO, "informativa", "informativas"))]
    parts = [f"{n} {one if n == 1 else many}" for n, one, many in counts if n]
    return ", ".join(parts) if parts else "sin observaciones"


def _print_validation(issues, limit: int | None) -> None:
    """Observaciones agrupadas por tema (el tema Libro, solo si tiene observaciones)."""
    from evm.messages import TOPIC_BOOK, TOPICS
    order = {SEV_ERROR: 0, SEV_WARN: 1, SEV_INFO: 2}
    for topic in TOPICS:
        # El resumen de las observaciones silenciadas va primero: no queda fuera del límite
        found = sorted((i for i in issues if i.topic == topic),
                       key=lambda i: (i.id != SUMMARY_ID, order.get(i.severity, 1)))
        if not found and topic == TOPIC_BOOK:
            continue
        print()
        print(f"{topic}: {_count_text(found)}")
        shown = found if limit is None else found[:limit]
        for i in shown:
            print(_issue_line(i))
        if len(found) > len(shown):
            print(f"  ... y {len(found) - len(shown)} más (--todas para verlas todas).")


def cmd_validar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    files = _sources(cfg, args)
    status = parse_date(args.fecha_corte) if args.fecha_corte else None
    from evm.validation import combine, validate_part
    parts = []
    with progress.tracking(_stages(files, args.motor, compute=10.0, after={"revisar": 5.0})):
        for name, origin, pcfg in files:
            progress.say(_reading(name, origin))
            with open_project(origin, pcfg, args.motor, args.visible, name=name) as project:
                _computing(name, "Revisión del cronograma")
                parts.append(validate_part(project, pcfg, name, status))
        progress.stage("revisar", "Revisión de los códigos EDT de todos los archivos")
        report, issues = combine(parts, cfg)
    issues = _silence(cfg).apply(issues)

    def when(project) -> str:
        if status is not None:
            return f"Fecha de estado: {status:%d/%m/%Y} (--fecha-corte)"
        if project.status_date:
            return f"Fecha de estado: {project.status_date:%d/%m/%Y}"
        return "Sin fecha de estado"

    reviewed = sum(1 for p in parts for _ in p.scope.tasks())
    if any(p.source.name for p in parts):
        print(f"Validación: {report.title if report else ' + '.join(p.source.project.name for p in parts)}")
        for p in parts:
            plan = f"  |  Planificación: {p.source.plan}" if p.source.plan else ""
            print(f"  {p.source.name}: {p.source.project.name}{plan}  |  {p.source.project.baseline_label}  |  "
                  f"{when(p.source.project)}")
    else:
        project = parts[0].source.project
        print(f"Validación: {project.name}")
        plan = f"Planificación: {parts[0].source.plan}  |  " if parts[0].source.plan else ""
        print(f"  {plan}{project.baseline_label}  |  {when(project)}")
    rows = report.rows if report else [r for p in parts for r in p.rows]
    leaves = sum(1 for r in rows if r.is_input)
    print(f"  Tareas revisadas: {reviewed}  |  Filas EDT: {len(rows)} ({leaves} de ingreso)  |  "
          f"{_rows_rule(cfg)}")
    _print_validation(issues, None if args.todas else 20)
    errors = sum(1 for i in issues if i.severity == SEV_ERROR)
    print()
    print(f"Resultado: {_count_text(issues)}.")
    if errors:
        print("Se deben corregir los errores en MS Project antes de generar o actualizar el libro.")
    print("No se creó ni se modificó ningún archivo.")
    return 1 if errors else 0


def cmd_generar_fichas(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    files = _sources(cfg, args)
    origen = files[0][1]
    if args.salida:
        salida = Path(args.salida)
    elif cfg.projects:
        raise ConfigError("CFG-040",
                          "Con varios archivos ([[proyectos]]) se debe indicar el archivo de las fichas con --salida.")
    elif origen.lower() == "activo":
        salida = Path("Fichas.md")
    else:
        salida = Path(origen).with_name(Path(origen).stem + "_fichas.md")
    if salida.suffix.lower() != ".md":
        raise ConfigError("CFG-041", "Las fichas se guardan en un archivo .md (Markdown).")
    _check_folder(salida)
    if salida.exists() and not args.sobrescribir:
        raise ConfigError("CFG-042", f"'{salida}' ya existe. Para reemplazarlo se debe usar --sobrescribir.")

    parts = []
    with progress.tracking(_stages(files, args.motor, compute=10.0, after={"fichas": 5.0})):
        for name, origin, pcfg in files:
            progress.say(_reading(name, origin))
            with open_project(origin, pcfg, args.motor, args.visible, name=name) as project:
                _computing(name, "Cálculo de las filas EDT")
                parts.append(build_part(project, pcfg, name))
        report = assemble_report(parts, cfg)
        from evm.task_cards import build_cards, write_cards
        progress.stage("fichas", "Escritura de las fichas")
        cards = build_cards(report, parts, cfg)
        write_cards(cards, cfg, salida)

    if report.multi_file:
        print(f"Fichas: {report.title}")
        for s in report.sources:
            plan = f"  |  Planificación: {s.plan}" if s.plan else ""
            print(f"  {s.name}: {s.project.name}{plan}  |  {s.project.baseline_label}")
    else:
        print(f"Proyecto: {report.project.name}")
        if report.plan:
            print(f"  Planificación: {report.plan}")
        print(f"  {report.project.baseline_label}")
    print(f"  Fichas: {len(cards.cards)}  |  {_rows_rule(cfg)}  |  Vínculos entre fichas: {cards.links}")
    silence = _silence(cfg)
    notices = [replace(w, message=f"[{s.name}] {w.message}") if s.name else w
               for s in report.sources for w in s.project.warnings]
    if not cards.cards:
        notices.append(Issue("FIC-001", message="Ninguna tarea tiene código EDT: no hay fichas. Se debe revisar el "
                                                "campo EDT configurado ([campos] edt) y la planificación elegida."))
    if cards.hidden:
        notices.append(Issue("FIC-002", message=f"{plural(cards.hidden, 'vínculo', 'vínculos')} con tareas sin ficha "
                                                "(de otra planificación, inactivas, con Visible = No, fuera de la "
                                                "lista EDT o de otro proyecto): no se muestran."))
    if cards.current:
        notices.append(Issue("FIC-003", message=f"{plural(cards.current, 'ficha', 'fichas')} sin línea base: se "
                                                "usaron las fechas y la duración del programa actual (se indica en "
                                                "la ficha)."))
    errors = sum(1 for i in report.issues if i.severity == SEV_ERROR and not silence.matches(i))
    if errors:
        notices.append(Issue("FIC-004", message=f"El archivo tiene {plural(errors, 'error', 'errores')} (por "
                                                "ejemplo, códigos EDT duplicados): se recomienda revisarlo con "
                                                "'validar'."))
    _print_notices(silence.apply(notices), limit=50)
    print(f"Fichas generadas: {salida.resolve()}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="evm_export",
        description="Exporta datos de MS Project a un libro Excel de Valor Ganado (PV mensual, EV, AC, curva S "
                    "e indicadores calculados en Excel).")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("-v", "--verbose", action="store_true", help="muestra mensajes de diagnóstico")
    sub = p.add_subparsers(dest="comando", required=True)

    def common_source(sp):
        sp.add_argument("--origen", action="append",
                        help="archivo .mpp (requiere MS Project instalado), .xml exportado de MS Project, "
                             "o 'activo' para usar el proyecto abierto en MS Project. Con varios archivos "
                             "([[proyectos]] en la configuración) se omite, o se indica NOMBRE=RUTA para reemplazar "
                             "la ruta de un archivo (se puede repetir)")
        sp.add_argument("--config", help="archivo de configuración TOML (ver config.ejemplo.toml)")
        sp.add_argument("--motor", choices=["auto", "com", "xml"], default="auto",
                        help="forma de lectura (por defecto según la extensión)")
        sp.add_argument("--linea-base", type=int, choices=range(0, 11), metavar="N",
                        help="línea base a usar: 0 = Línea base, 1..10 = Línea base N")
        sp.add_argument("--nivel-max", type=int, help="nivel máximo del código EDT (por defecto, nivel_maximo de la configuración: 3)")
        sp.add_argument("--planificacion", metavar="CODIGO",
                        help="código EDT de la planificación a exportar, p. ej. INTERNO o CLIENTE (reemplaza el "
                             "valor del archivo de configuración; \"\" = todas)")
        sp.add_argument("--visible", action="store_true", help="muestra MS Project durante la lectura (COM)")
        silence_option(sp)

    def silence_option(sp):
        sp.add_argument("--silenciar", action="append", metavar="CODIGOS",
                        help="códigos de las observaciones que no se muestran, separados por comas (p. ej. "
                             "FEC-003,LOG-004 o MET-002:1.2.3 para una fila); se suman a [validacion] silenciar. "
                             "Ver docs/codigos.md")

    g = sub.add_parser("generar", help="crea el libro Excel con el PV mensual a partir de MS Project")
    common_source(g)
    g.add_argument("--salida", help="libro .xlsx a crear (por defecto <origen>_EVM.xlsx)")
    g.add_argument("--mes-inicio", help="primer mes del horizonte, AAAA-MM (por defecto, mes_inicio de la "
                                        "configuración o el primer mes con %% planificado)")
    g.add_argument("--meses", type=int, help="cantidad de meses (por defecto, meses de la configuración: 120)")
    g.add_argument("--migrar-desde", help="libro anterior del que se copian el BAC, el %% ganado y el AC ya cargados")
    g.add_argument("--sobrescribir", action="store_true", help="reemplaza el libro de salida si existe")
    g.set_defaults(func=cmd_generar)

    a = sub.add_parser("actualizar", help="calcula el %% ganado a la fecha de corte y lo carga en el mes correspondiente (el AC se carga a mano)")
    common_source(a)
    a.add_argument("--libro", required=True, help="libro .xlsx generado con 'generar'")
    a.add_argument("--fecha-corte", help="fecha de corte AAAA-MM-DD (por defecto la fecha de estado del proyecto; "
                                         "si el proyecto no la tiene, es obligatoria)")
    a.add_argument("--forzar", action="store_true", help="permite actualizar aunque haya meses posteriores cargados")
    a.add_argument("--sin-respaldo", action="store_true", help="no crea copia de respaldo del libro")
    a.add_argument("--simular", action="store_true", help="muestra el resultado sin modificar el libro")
    a.set_defaults(func=cmd_actualizar)

    v = sub.add_parser("validar", help="revisa el archivo de MS Project (códigos EDT, fechas, duración, pesos, "
                                       "método, empresas, avance y lógica del cronograma) y muestra el resultado, "
                                       "sin crear ningún archivo")
    common_source(v)
    v.add_argument("--fecha-corte", help="fecha de estado para revisar el avance, AAAA-MM-DD (por defecto la fecha de "
                                         "estado del proyecto)")
    v.add_argument("--todas", action="store_true", help="muestra todas las observaciones (por defecto, hasta 20 "
                                                        "por tema)")
    v.set_defaults(func=cmd_validar)

    f = sub.add_parser("generar-fichas", help="crea un documento Markdown (.md) con la ficha de cada tarea "
                                              "exportada: código EDT, nombre, fechas y duración de la línea base, "
                                              "predecesoras y sucesoras")
    common_source(f)
    f.add_argument("--salida", help="archivo .md a crear (por defecto <origen>_fichas.md)")
    f.add_argument("--sobrescribir", action="store_true", help="reemplaza el archivo de salida si existe")
    f.set_defaults(func=cmd_generar_fichas)

    n = sub.add_parser("analizar", help="actualiza la hoja 'Análisis de desviaciones' con los valores del libro, "
                                        "sin leer MS Project (p. ej. después de cargar a mano el AC del mes)")
    n.add_argument("--libro", required=True, help="libro .xlsx generado con 'generar'")
    n.add_argument("--config", help="archivo de configuración: solo se usan [validacion] silenciar y loe_maximo (el "
                                    "libro guarda los umbrales del semáforo)")
    n.add_argument("--sin-respaldo", action="store_true", help="no crea copia de respaldo del libro")
    silence_option(n)
    n.set_defaults(func=cmd_analizar)

    return p


def main(argv: list[str] | None = None) -> int:
    """Códigos de salida: 0 correcto; 1 'validar' encontró errores; 2 error previsto (configuración, archivos,
    opciones); 3 error inesperado del programa; 130 cancelado."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except AttributeError:  # pragma: no cover
            pass
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format="%(levelname)s %(name)s: %(message)s")
    print(f"Exportador EVM {__version__}", flush=True)
    try:
        return args.func(args)
    except EvmError as exc:                 # errores previstos, con su código del catálogo (docs/codigos.md)
        if args.verbose:
            log.exception("Error")
        _error(exc.id, str(exc))
        return 2
    except (FileNotFoundError, PermissionError, ValueError, RuntimeError) as exc:
        if args.verbose:
            log.exception("Error")
        _error("SIS-003", str(exc))
        return 2
    except OSError as exc:  # disco lleno, unidad de red no disponible, ruta no válida...
        if args.verbose:
            log.exception("Error")
        _error("SIS-002", f"no se pudo leer o escribir un archivo: {exc}")
        return 2
    except KeyboardInterrupt:
        print("Cancelado.", file=sys.stderr)
        return 130
    except Exception as exc:  # errores no previstos: mensaje breve; con -v, el detalle completo
        if args.verbose:
            log.exception("Error inesperado")
        else:
            print("Para ver el detalle técnico se debe ejecutar de nuevo con -v.", file=sys.stderr)
        _error("SIS-001", f"error inesperado: {type(exc).__name__}: {exc}")
        return 3


def _error(id: str, text: str) -> None:
    """Error que detiene el programa: 'ERROR CFG-001: mensaje' en la salida de error."""
    print(f"ERROR {id}: {text}", file=sys.stderr)
