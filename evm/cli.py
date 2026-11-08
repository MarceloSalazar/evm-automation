"""Interfaz de línea de comandos.

    python evm_export.py generar    --origen proyecto.mpp --salida EVM.xlsx
    python evm_export.py actualizar --origen proyecto.mpp --libro EVM.xlsx --fecha-corte 2026-10-31
    python evm_export.py validar    --origen proyecto.mpp
    python evm_export.py analizar   --libro EVM.xlsx
    python evm_export.py generar-fichas --origen proyecto.mpp --salida Fichas.md
    python evm_export.py procesar-costes --libro EVM_Depuracion.xlsx
    python evm_export.py generar-cliente --origen interno.mpp --simular
"""
from __future__ import annotations

import argparse
import copy
import logging
import platform
import sys
import time
from contextlib import contextmanager
from dataclasses import replace
from datetime import date
from pathlib import Path

from evm import __version__, progress
from evm.config import Config, ConfigError, load_config, parse_date, parse_month
from evm.fields import fmt_number
from evm.messages import (SEV_ERROR, SEV_INFO, SEV_WARN, SUMMARY_ID, TARGET_DETAIL, TARGET_GENERAL, TARGET_OUT,
                          TARGET_ROW, TARGETS, EvmError, Issue, MissingFileError, Silence, by_priority, grouped,
                          parse_silence)
from evm.wbs import assemble_report, build_part, fmt_money, located, plural

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
                 name: str = "", links: bool = True, review: bool = False, client: bool = False):
    """Devuelve el proyecto leído (ProjectData). Con COM, MS Project queda abierto mientras dure el bloque.

    name: nombre del archivo en [[proyectos]] (etapas de la barra de progreso).
    links: lee los vínculos y las delimitaciones de las tareas (validar y generar-fichas).
    review: con COM, lecturas que solo usa validar (líneas base y calendario de cada tarea; el XML las tiene).
    client: el campo de [cliente] campo y el formato de la posición de los vínculos (generar-cliente).
    """
    engine = _engine_for(origen, motor)
    if origen.lower() != "activo" and not Path(origen).is_file():
        raise MissingFileError("LEC-001", f"No existe el archivo de origen: {origen}")
    which = f" ({name})" if name else ""
    if engine == "xml":
        from evm.reader_xml import read_xml
        progress.stage(f"leer:{name}", f"Lectura del XML{which}")
        yield read_xml(origen, cfg, links=links, client=client)
        return
    from evm.reader_com import ComSession
    progress.stage(f"abrir:{name}", f"Apertura en MS Project{which}")
    with ComSession(origen, visible=visible) as session:
        progress.stage(f"leer:{name}", f"Lectura de MS Project{which}")
        yield session.read_project(cfg, status_date=status, links=links, review=review, client=client)


def _stages(files: list[tuple[str, str, Config]], motor: str, before: dict[str, float] | None = None,
            compute: float = 0.0, after: dict[str, float] | None = None) -> dict[str, float]:
    """Etapas de la barra de progreso y su peso relativo: por archivo, apertura en MS Project (COM), lectura y
    cálculo; antes y después, las del comando.

    Peso = segundos aproximados con un proyecto de unas 2.000 tareas y 120 meses (medidos con XML; la apertura y la
    lectura con MS Project, estimadas). Solo cuenta la proporción entre las etapas: el ETA se ajusta al ritmo real.
    """
    stages = dict(before or {})
    for name, origin, _cfg in files:
        if _engine_for(origin, motor) == "com":
            stages[f"abrir:{name}"] = 8.0
            stages[f"leer:{name}"] = 10.0
        else:
            stages[f"leer:{name}"] = 0.5
        if compute:
            stages[f"calcular:{name}"] = compute
    stages.update(after or {})
    return stages


def _book_stages(months: int, **stages: float) -> dict[str, float]:
    """Etapas del libro (lectura, cálculo, escritura, guardado) con su peso para 120 meses, ajustado a los meses del
    horizonte: con 36 meses llevan la mitad del tiempo."""
    scale = 0.3 + 0.7 * months / 120
    return {k: v * scale for k, v in stages.items()}


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


# Alcance de las observaciones en el resumen de la consola
TARGET_LABELS = {TARGET_GENERAL: "generales", TARGET_ROW: "en filas EDT", TARGET_DETAIL: "en tareas de detalle",
                 TARGET_OUT: "fuera del libro"}


def _issue_line(i: Issue) -> str:
    """'  [!] FEC-003 1.2.3 Montaje: mensaje' (el código de la observación, la fila o tarea y el mensaje). Una tarea
    de detalle indica su fila del libro ('2.7.1.3 Montaje (detalle de 2.7.1)'); una línea agrupada, solo el
    mensaje."""
    where = "" if i.tasks else f"{i.code} {i.name}".strip()
    if where and i.target == TARGET_DETAIL and i.row:
        where += f" (detalle de {i.row})"
    elif where and i.target == TARGET_OUT:
        where += " (fuera del libro)"
    return f"  {MARKS.get(i.severity, '[!]')} {i.id} {where + ': ' if where else ''}{i.message}"


def _targets_text(issues: list[Issue]) -> str:
    """'en filas EDT: 2 errores, 3 advertencias; en tareas de detalle: 40 advertencias' (los alcances que tienen
    observaciones, en el orden de prioridad)."""
    found = {t: [i for i in issues if (i.target or TARGET_GENERAL) == t] for t in TARGETS}
    return "; ".join(f"{TARGET_LABELS[t]}: {_count_text(v)}" for t, v in found.items() if v)


def _print_issues(issues: list[Issue], limit: int = 15) -> None:
    """Las informativas y después las demás, por prioridad (severidad y alcance), con las de las tareas de detalle y
    fuera del libro agrupadas por código y fila del libro."""
    lines = by_priority(grouped(issues, "Una por tarea en la hoja 'Validación detalle'."))
    serious = [i for i in lines if i.severity != SEV_INFO]
    infos = [i for i in lines if i.severity == SEV_INFO]
    for i in infos[:limit]:
        print(_issue_line(i))
    if len(infos) > limit:
        print(f"  ... y {len(infos) - limit} informativas más (ver hoja 'Validación').")
    for i in serious[:limit]:
        print(_issue_line(i))
    if len(serious) > limit:
        print(f"  ... y {len(serious) - limit} observaciones más (ver hoja 'Validación').")


def cmd_generar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    _one_plan(cfg)
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
    after = {"migrar": 9.9} if args.migrar_desde else {}
    with progress.tracking(_stages(files, args.motor, compute=0.3,
                                   after=_book_stages(cfg.months, **after, escribir=6.5, guardar=3.8))):
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
            report.issues.extend(plan_migration(report, prior, cfg.max_level, cfg.ignore_interior_zeros))
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
    issues = located(report.rows, report.issues)
    _print_issues(issues)
    if any(i.severity != SEV_INFO for i in issues):
        print(f"  Por alcance: {_targets_text(issues)}.")
    if prior is not None:
        print(f"Migración desde {args.migrar_desde}: {prior.matched} filas con historial trasladado.")
        for i in unmatched:
            print(_issue_line(i))
    print(f"Libro generado: {salida.resolve()}")
    return 0


def _plans(cfg: Config) -> list[str]:
    """Planificaciones indicadas en planificacion o --planificacion: varias, separadas por comas, solo en validar."""
    return list(dict.fromkeys(p.strip() for p in (cfg.plan or "").split(",") if p.strip()))


def _one_plan(cfg: Config) -> None:
    """generar, actualizar y generar-fichas trabajan con una sola planificación por archivo (CFG-053)."""
    if any("," in (plan or "") for plan in [cfg.plan] + [spec.plan for spec in cfg.projects]):
        raise ConfigError("CFG-053", "Se indicaron varias planificaciones separadas por comas: solo validar las revisa "
                                     "(validar --planificacion INTERNO,CLIENTE). Se debe indicar una sola "
                                     "planificación.")


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
    return f"Nivel máximo: {cfg.max_level}" + (" (sin los ceros intermedios)" if cfg.ignore_interior_zeros else "")


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
    notes = _same_note_rows(silence.filter([n for r in res.rows for n in r.notes]))
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


def _same_note_rows(notes: list[Issue]) -> list[Issue]:
    """Observaciones por fila de actualizar: las de varias filas con el mismo código y el mismo texto (p. ej. ACT-034,
    método no definido) en una línea con la cantidad de filas y sus códigos EDT, en el lugar de la primera."""
    groups: dict[tuple[str, str], list[Issue]] = {}
    for n in notes:
        groups.setdefault((n.id, n.message), []).append(n)
    out: list[Issue] = []
    for n in notes:
        same = groups.get((n.id, n.message))
        if same is None:
            continue                            # ya se mostró en la línea de su grupo
        del groups[(n.id, n.message)]
        if len(same) == 1:
            out.append(n)
            continue
        codes = ", ".join(x.code for x in same[:10]) + (f" y {len(same) - 10} más" if len(same) > 10 else "")
        out.append(replace(n, code="", name="", tasks=len(same),
                           message=f"{n.message} en {len(same)} filas ({codes})."))
    return out


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
    meta = book_meta(args.libro)            # libro inexistente o que no es de este programa: error con su código
    check_writable(args.libro)
    with progress.tracking(_book_stages(meta["meses"], libro=9.7, analizar=0.2, guardar=4.5)):
        res = run_analysis(cfg, args.libro, backup=not args.sin_respaldo)
    print(f"Fecha de corte: {res.status_date:%d/%m/%Y} (mes {res.month_index + 1} del horizonte)")
    _print_notices(_silence(cfg).apply(res.warnings))
    _print_deviations(res)
    if res.backup:
        print(f"Respaldo: {res.backup}")
    print(f"Libro actualizado: {Path(args.libro).resolve()}")
    return 0


def cmd_procesar_costes(args) -> int:
    cfg = load_config(args.config)
    _add_silence(cfg, args)
    if args.calculos:
        cfg.cost_review_calcs = True
    from evm.cost_plan import run_cost_plan
    from evm.update import book_meta
    meta = book_meta(args.libro)
    if not args.simular:
        from evm.workbook import check_writable
        check_writable(args.libro)
    after = {} if args.simular else {"guardar": 4.5}
    with progress.tracking(_book_stages(meta["meses"], libro=9.7, procesar=0.3, **after)):
        res = run_cost_plan(cfg, args.libro, backup=not args.sin_respaldo, dry_run=args.simular)
    decimals = int(res.book.meta.get("decimales") or 0)
    money = lambda v: fmt_number(v, decimals)       # noqa: E731
    if res.calculated is not None:
        print(f"Cálculos del plan de costes: {plural(res.calculated, 'línea', 'líneas')} de la hoja PV-depuración "
              "(sumas, EAC y BAC de las EDT cliente)" + (" (no se escriben en la simulación)" if args.simular else ""))
    print(f"Hoja: {res.sheet}  |  Plan de costes: {plural(res.lines, 'línea', 'líneas')} de EDT cliente, "
          f"{res.counted} con fila de ingreso  |  Importe leído: {money(res.total)}")
    print(f"Plan de costes en {len(res.rows)} de {plural(res.inputs, 'fila', 'filas')} de ingreso, "
          f"{money(res.assigned)} dentro del horizonte:")
    for code, name, eac in res.rows[:15]:
        print(f"  {code:<12} {name[:48]:<48}  EAC {money(eac):>14}")
    if len(res.rows) > 15:
        print(f"  ... y {len(res.rows) - 15} filas más.")
    _print_notices(_silence(cfg).apply(res.warnings))
    print(f"% planificado (hoja '% {res.sheet}') = importe del mes / BAC: el PV de cada mes (BAC × %) es el importe "
          "del plan de costes.")
    if args.simular:
        print("Simulación: no se modificó el libro.")
    else:
        if res.backup:
            print(f"Respaldo: {res.backup}")
        print(f"Libro actualizado: {Path(args.libro).resolve()}")
    return 0


def cmd_actualizar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    _one_plan(cfg)
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
    after = _book_stages(meta["meses"], libro=9.6, calcular=0.4, **({} if args.simular else {"guardar": 4.8}))
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
    """Observaciones agrupadas por tema (el tema Libro, solo si tiene observaciones), por prioridad: severidad y
    alcance. Con límite (sin --todas), las de las tareas de detalle y fuera del libro van en una línea por código y
    fila del libro."""
    from evm.messages import TOPIC_BOOK, TOPICS
    for topic in TOPICS:
        # El resumen de las observaciones silenciadas va primero: no queda fuera del límite
        found = sorted(by_priority([i for i in issues if i.topic == topic]), key=lambda i: i.id != SUMMARY_ID)
        if not found and topic == TOPIC_BOOK:
            continue
        print()
        print(f"{topic}: {_count_text(found)}")
        lines = found if limit is None else grouped(found, "Una por tarea con --todas.")
        shown = lines if limit is None else lines[:limit]
        for i in shown:
            print(_issue_line(i))
        if len(lines) > len(shown):
            print(f"  ... y {len(lines) - len(shown)} más (--todas para verlas todas).")


def _report_path(args, cfg: Config, files) -> Path:
    """Archivo del informe de mejoras (--informe): el indicado o <origen>_mejoras.md."""
    origen = files[0][1]
    if args.informe:
        path = Path(args.informe)
    elif cfg.projects:
        raise ConfigError("CFG-040", "Con varios archivos ([[proyectos]]) se debe indicar el archivo del informe de "
                                     "mejoras: --informe RUTA.md.")
    elif origen.lower() == "activo":
        path = Path("proyecto_mejoras.md")
    else:
        path = Path(origen).with_name(Path(origen).stem + "_mejoras.md")
    if path.suffix.lower() != ".md":
        raise ConfigError("CFG-041", "El informe de mejoras se guarda en un archivo .md (Markdown).")
    _check_folder(path)
    if path.exists() and not args.sobrescribir:
        raise ConfigError("CFG-042", f"'{path}' ya existe. Para reemplazarlo se debe usar --sobrescribir.")
    return path


def _print_unit(unit, cfg: Config, status: date | None, limit: int | None) -> int:
    """Una revisión de validar: encabezado, observaciones por tema, evaluación DCMA y resultado. Devuelve los
    errores."""
    parts = unit.parts

    def when(project) -> str:
        if status is not None:
            return f"Fecha de estado: {status:%d/%m/%Y} (--fecha-corte)"
        if project.status_date:
            return f"Fecha de estado: {project.status_date:%d/%m/%Y}"
        return "Sin fecha de estado"

    reviewed = sum(1 for p in parts for _ in p.scope.tasks())
    if any(p.source.name for p in parts):
        title = unit.report.title if unit.report else " + ".join(p.source.project.name for p in parts)
        print(f"Validación: {title}")
        for p in parts:
            plan = f"  |  Planificación: {p.source.plan}" if p.source.plan else ""
            print(f"  {p.source.name}: {p.source.project.name}{plan}  |  {p.source.project.baseline_label}  |  "
                  f"{when(p.source.project)}")
    else:
        project = parts[0].source.project
        print(f"Validación: {project.name}")
        plan = f"Planificación: {parts[0].source.plan}  |  " if parts[0].source.plan else ""
        print(f"  {plan}{project.baseline_label}  |  {when(project)}")
    rows = unit.report.rows if unit.report else [r for p in parts for r in p.rows]
    leaves = sum(1 for r in rows if r.is_input)
    print(f"  Tareas revisadas: {reviewed}  |  Filas EDT: {len(rows)} ({leaves} de ingreso)  |  "
          f"{_rows_rule(cfg)}")
    _print_validation(unit.issues, limit)
    _print_dcma(parts)
    errors = sum(1 for i in unit.issues if i.severity == SEV_ERROR)
    print()
    print(f"Resultado: {_count_text(unit.issues)}.")
    if unit.issues:
        print(f"  Por alcance: {_targets_text(unit.issues)}.")
    return errors


def _print_dcma(parts) -> None:
    """Evaluación DCMA de 14 puntos de cada archivo: valor, umbral y si cumple."""
    from evm.review import dcma_rows
    for k, part in enumerate(parts):
        rows = dcma_rows([part])
        width = max(len(text) for _p, _n, values, _t in rows for text in [(values[0] or ("-",))[0]])
        print()
        print("Evaluación DCMA de 14 puntos" + (f" [{part.source.name}]" if part.source.name else "") + ":")
        for point, name, values, threshold in rows:
            text, _threshold, ok = values[0] or ("-", threshold, None)
            result = "" if ok is None else ("cumple" if ok else "no cumple")
            print(f"  {point:>2}. {name:<38} {text:<{width}}  {threshold:<7} {result}".rstrip())


def _print_comparison(units) -> None:
    """Varias planificaciones: sus cifras lado a lado y las diferencias entre sus filas."""
    from evm.review import compare, facts, facts_table
    items = [facts(u.parts[0], u.issues, u.title, u.status) for u in units]
    table = facts_table(items)
    first = max(len(name) for name, _ in table)
    widths = [max(len(f.label), *(len(values[k]) for _n, values in table)) for k, f in enumerate(items)]
    print()
    print("Comparación de planificaciones:")
    print("  " + " " * first + "".join(f"  {f.label:>{w}}" for f, w in zip(items, widths)))
    for name, values in table:
        print(f"  {name:<{first}}" + "".join(f"  {v:>{w}}" for v, w in zip(values, widths)))
    for line in compare(units):
        print(f"  {line}")


def cmd_validar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    files = _sources(cfg, args)
    plans = _plans(cfg)
    several = len(plans) > 1
    if (several and cfg.projects) or any("," in (spec.plan or "") for spec in cfg.projects):
        raise ConfigError("CFG-053", "Varias planificaciones separadas por comas solo se revisan con un archivo, sin "
                                     "[[proyectos]]: se debe indicar una sola planificación en cada archivo.")
    status = parse_date(args.fecha_corte) if args.fecha_corte else None
    report_path = _report_path(args, cfg, files) if args.informe is not None else None
    from evm.review import Unit
    from evm.validation import combine, validate_part
    units: list[Unit] = []
    weight = 0.2 * max(1, len(plans))
    with progress.tracking(_stages(files, args.motor, compute=weight, after={"revisar": 0.1})):
        if several:
            # Un archivo, varias planificaciones: se lee una vez y se revisa cada planificación por separado
            name, origin, pcfg = files[0]
            progress.say(_reading(name, origin))
            with open_project(origin, pcfg, args.motor, args.visible, name=name, review=True) as project:
                for plan in plans:
                    _computing(name, f"Revisión de la planificación {plan}")
                    part = validate_part(copy.deepcopy(project), replace(pcfg, plan=plan), name, status)
                    report, issues = combine([part], replace(cfg, plan=plan))
                    units.append(Unit(plan, [part], report, issues, status))
            progress.stage("revisar", "Comparación de las planificaciones")
        else:
            parts = []
            for name, origin, pcfg in files:
                progress.say(_reading(name, origin))
                with open_project(origin, pcfg, args.motor, args.visible, name=name, review=True) as project:
                    _computing(name, "Revisión del cronograma")
                    parts.append(validate_part(project, pcfg, name, status))
            progress.stage("revisar", "Revisión de los códigos EDT de todos los archivos")
            report, issues = combine(parts, cfg)
            units.append(Unit("", parts, report, issues, status))
    silence = _silence(cfg)
    errors = 0
    for k, unit in enumerate(units):
        rows = unit.report.rows if unit.report else [r for p in unit.parts for r in p.rows]
        unit.issues = located(rows, silence.apply(unit.issues))
        if k:
            print()
            print("=" * 100)
        errors += _print_unit(unit, cfg, status, None if args.todas else 20)
    if several:
        _print_comparison(units)
    if errors:
        print("Se deben corregir los errores en MS Project antes de generar o actualizar el libro.")
    if report_path is not None:
        from evm.review import write_report
        engine = ("MS Project (.mpp)" if units[0].parts[0].source.project.backend == "com" else "XML de MS Project")
        write_report(report_path, units, cfg, __version__, engine)
        print(f"Informe de mejoras: {report_path.resolve()} (sin nombres de tareas). No se modificó ningún otro "
              "archivo.")
    else:
        print("No se creó ni se modificó ningún archivo.")
    return 1 if errors else 0


def cmd_generar_fichas(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    _one_plan(cfg)
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
    with progress.tracking(_stages(files, args.motor, compute=0.3, after={"fichas": 0.2})):
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


def _client_paths(args, origen: str) -> tuple[Path | None, Path]:
    """Archivo del cliente (--destino; por defecto <origen>_cliente.mpp) y del informe de control (--informe; por
    defecto <destino>_control.md)."""
    destino = None
    if args.destino:
        destino = Path(args.destino)
    elif origen.lower() != "activo":
        destino = Path(origen).with_name(Path(origen).stem + "_cliente.mpp")
    if destino is not None and destino.suffix.lower() != ".mpp":
        raise ConfigError("CFG-041", "La planificación del cliente se guarda en un archivo .mpp (--destino).")
    if args.informe:
        informe = Path(args.informe)
    elif destino is not None:
        informe = destino.with_name(destino.stem + "_control.md")
    else:
        informe = Path("Cliente_control.md")
    if informe.suffix.lower() != ".md":
        raise ConfigError("CFG-041", "El informe de control se guarda en un archivo .md (Markdown).")
    return destino, informe


def _pct(value: float | None) -> str:
    return "—" if value is None else f"{fmt_number(round(value, 1), trim=True)} %"


def cmd_generar_cliente(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    _one_plan(cfg)
    if cfg.projects:
        raise EvmError("CLI-009", "generar-cliente trabaja con un solo archivo de MS Project: se debe quitar las "
                                  "secciones [[proyectos]] de la configuración e indicar el archivo con --origen.")
    from evm.client_check import predict
    from evm.client_plan import (INTERNAL_NAME, SKIP_LINKED, build_client_plan, require_client_field, require_field,
                                 write_control)
    require_field(cfg)
    files = _sources(cfg, args)
    origen = files[0][1]
    destino, informe = _client_paths(args, origen)
    generate = not args.simular
    if generate:
        if origen.lower() == "activo" or _engine_for(origen, args.motor) != "com" or \
                Path(origen).suffix.lower() not in COM_EXTENSIONS:
            raise EvmError("CLI-010", "Para generar el archivo del cliente, el origen debe ser el archivo .mpp de la "
                                      "planificación interna (se copia y MS Project trabaja sobre la copia). Con un XML "
                                      "o con 'activo' solo se puede simular (--simular).")
        if not Path(origen).is_file():
            raise MissingFileError("LEC-001", f"No existe el archivo de origen: {origen}")
        if destino.resolve() == Path(origen).resolve():
            raise EvmError("CLI-011", "El archivo del cliente (--destino) no puede ser el archivo de origen.")
        _check_folder(destino)
        if destino.exists() and not args.sobrescribir:
            raise ConfigError("CFG-042", f"'{destino}' ya existe. Para reemplazarlo se debe usar --sobrescribir.")
    _check_folder(informe)
    if informe.exists() and not args.sobrescribir:
        raise ConfigError("CFG-042", f"'{informe}' ya existe. Para reemplazarlo se debe usar --sobrescribir.")

    silence = _silence(cfg)
    outcome = None
    if generate:
        from evm.client_writer import generate_client
        # Segundos de cada etapa en una planificación real de unas 1.450 tareas (registro de diagnóstico)
        stages = {"copiar": 0.5, "abrir:": 3.0, "leer:": 78.0, "calcular:": 5.0, "borrar": 9.0, "vincular": 0.5,
                  "anonimizar": 17.0, "fechas": 3.0, "contenido": 3.0, "guardar": 1.0, "comparar:": 30.0,
                  "informe": 0.1}
        with progress.tracking(stages):
            progress.say(f"Generando la planificación del cliente a partir de {origen} ...")
            plan, outcome = generate_client(origen, destino, cfg, visible=args.visible, overwrite=args.sobrescribir)
            project = plan.project
            result = f"Archivo del cliente generado: {destino.resolve()} (de solo lectura)."
            notices = silence.apply(project.warnings + plan.issues() + outcome.issues(plan))
            progress.stage("informe", "Escritura del informe de control")
            write_control(plan, origen, informe, result, notices, outcome=outcome)
    else:
        result = "Simulación: no se generó el archivo del cliente."
        with progress.tracking(_stages(files, args.motor, compute=1.0, after={"informe": 0.1})):
            progress.say(_reading("", origen))
            with open_project(origen, cfg, args.motor, args.visible, client=True) as project:
                require_client_field(project, cfg)
                _computing("", "Cálculo de la planificación del cliente")
                plan = build_client_plan(project, cfg)
                plan.comparison = predict(plan, cfg)
                notices = silence.apply(project.warnings + plan.issues())
                progress.stage("informe", "Escritura del informe de control")
                write_control(plan, origen, informe, result, notices)

    print(f"Proyecto: {project.name}")
    if plan.root is not None:
        renamed = f" (en el archivo del cliente: {plan.new_name})" if plan.new_name else ""
        print(f"  Planificación: {plan.scope.prefix}{renamed}")
    print(f"  Campo: {cfg.client_field}  |  Tareas: {len(plan.tasks)}  |  Se muestran: {len(plan.shown)}  |  "
          f"No se muestran: {len(plan.hidden)}  |  De otras planificaciones: {len(plan.outside)}")
    if plan.internal_activity:
        print(f"  Tareas que no se muestran: {len(plan.anonymous)} como «{INTERNAL_NAME}», {len(plan.dropped)} "
              f"inactivas se quitan  |  Vínculos: {plan.kept_links} se conservan")
    else:
        skipped = sum(1 for s in plan.skipped if s.reason != SKIP_LINKED)
        print(f"  Tareas que no se muestran: se borran (actividad_interna = false)  |  Resúmenes que pasan a ser "
              f"tareas: {len(plan.collapsed)}  |  Vínculos: {plan.kept_links} se conservan, {len(plan.new_links)} "
              f"nuevos" + (f", {skipped} no se crean" if skipped else ""))
    if outcome is not None and not plan.internal_activity:
        print(f"  Fechas: {len(outcome.fixed)} fijadas con «No comenzar antes del», {len(outcome.differences)} "
              f"distintas de la planificación interna  |  Cálculos de MS Project: {outcome.rounds}")
    comparison = outcome.comparison if outcome is not None else plan.comparison
    if comparison is not None:
        book = comparison.book
        rows = ("no se pudo comparar" if book.error else
                f"{book.same} filas iguales, {len(book.moved)} con otro código, {book.deviations} con otras tareas o "
                "valores")
        print(f"  % completado del proyecto: {_pct(comparison.percent_internal)} (interno) y "
              f"{_pct(comparison.percent_client)} (cliente)  |  Libro EVM: {rows}")
        if comparison.checked:
            print(f"  Tareas comparadas: {comparison.checked}, con datos distintos: {len(comparison.tasks)}")
    _print_notices(notices, limit=50)
    print(result)
    print(f"Informe de control: {informe.resolve()}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="evm_export",
        description="Exporta datos de MS Project a un libro Excel de Valor Ganado (PV mensual, EV, AC, curva S "
                    "e indicadores calculados en Excel).")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    p.add_argument("-v", "--verbose", action="store_true", help="muestra mensajes de diagnóstico")
    p.add_argument("--registro", metavar="RUTA", help="escribe el registro de diagnóstico completo (con fecha y hora, "
                                                      "y el detalle de los errores) en el archivo indicado, para "
                                                      "revisar un problema; no incluye nombres de tareas en los pasos "
                                                      "de generar-cliente")
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
                             "valor del archivo de configuración; \"\" = todas). validar admite varias separadas por "
                             "comas (INTERNO,CLIENTE): las revisa por separado y las compara")
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

    v = sub.add_parser("validar", help="revisa el archivo de MS Project (códigos EDT, fechas, duración, línea "
                                       "base, pesos, método, empresas, avance y lógica del cronograma) y muestra el "
                                       "resultado y la evaluación DCMA; con --informe, escribe además un informe de "
                                       "mejoras sin nombres de tareas")
    common_source(v)
    v.add_argument("--fecha-corte", help="fecha de estado para revisar el avance, AAAA-MM-DD (por defecto la fecha de "
                                         "estado del proyecto)")
    v.add_argument("--todas", action="store_true", help="muestra todas las observaciones, una por tarea (por "
                                                        "defecto, hasta 20 líneas por tema, con las de las tareas de "
                                                        "detalle agrupadas por fila del libro)")
    v.add_argument("--informe", nargs="?", const="", default=None, metavar="RUTA",
                   help="escribe un informe de mejoras en Markdown, sin nombres de tareas, de empresas ni de "
                        "calendarios (por defecto <origen>_mejoras.md)")
    v.add_argument("--sobrescribir", action="store_true", help="reemplaza el informe de mejoras si existe")
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
    n.add_argument("--config", help="archivo de configuración: solo se usan [validacion] silenciar y loe_maximo y "
                                    "[exportacion] guardado_rapido (el libro guarda los umbrales del semáforo)")
    n.add_argument("--sin-respaldo", action="store_true", help="no crea copia de respaldo del libro")
    silence_option(n)
    n.set_defaults(func=cmd_analizar)

    c = sub.add_parser("procesar-costes", help="lleva el importe mensual del plan de costes de la hoja PV-depuración "
                                               "(líneas de EDT cliente) al %% planificado de cada fila de la empresa "
                                               "que se depura: el PV (BAC × %%) es el importe del plan")
    c.add_argument("--libro", required=True, help="libro .xlsx generado con depuracion_costes = true")
    c.add_argument("--calculos", action="store_true",
                   help="escribe antes en la hoja PV-depuración las fórmulas de los cálculos del plan: sumas de los "
                        "componentes y de las EDT interna, EAC y BAC de las EDT cliente (como [exportacion] "
                        "depuracion_calculos = true)")
    c.add_argument("--config", help="archivo de configuración: solo se usan [validacion] silenciar y [exportacion] "
                                    "guardado_rapido y depuracion_calculos")
    c.add_argument("--simular", action="store_true", help="muestra el resultado sin modificar el libro")
    c.add_argument("--sin-respaldo", action="store_true", help="no crea copia de respaldo del libro")
    silence_option(c)
    c.set_defaults(func=cmd_procesar_costes)

    k = sub.add_parser("generar-cliente", help="genera la planificación para el cliente: un .mpp con solo las tareas "
                                               "marcadas con Sí en [cliente] campo (y en sus tareas superiores), "
                                               "con la lógica rehecha sin las tareas ocultas, y un informe de "
                                               "control interno")
    common_source(k)
    k.add_argument("--destino", help="archivo .mpp del cliente a crear (por defecto <origen>_cliente.mpp); queda "
                                     "de solo lectura")
    k.add_argument("--informe", help="informe de control en Markdown (por defecto <destino>_control.md)")
    k.add_argument("--simular", action="store_true", help="calcula los cambios y escribe el informe de control sin "
                                                          "generar el archivo del cliente (admite el XML)")
    k.add_argument("--sobrescribir", action="store_true", help="reemplaza el archivo del cliente y el informe de "
                                                               "control si existen")
    k.set_defaults(func=cmd_generar_cliente)

    return p


class _LogFile:
    """--registro: el registro de diagnóstico completo (DEBUG, con fecha y hora) en un archivo. Mientras dura el
    comando, el registro «evm» escribe en el archivo y en la consola (lo mismo que sin --registro: advertencias, o
    todo con -v) y no pasa sus mensajes al registro raíz."""

    def __init__(self, path: str, verbose: bool):
        self.logger = logging.getLogger("evm")
        self.file = logging.FileHandler(path, mode="w", encoding="utf-8")
        self.file.setLevel(logging.DEBUG)
        self.file.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        self.console = logging.StreamHandler(sys.stderr)
        self.console.setLevel(logging.DEBUG if verbose else logging.WARNING)
        self.console.setFormatter(logging.Formatter("%(levelname)s %(name)s: %(message)s"))
        self.saved = (self.logger.level, self.logger.propagate)
        self.logger.setLevel(logging.DEBUG)
        self.logger.propagate = False
        self.logger.addHandler(self.file)
        self.logger.addHandler(self.console)

    def close(self) -> None:
        for handler in (self.file, self.console):
            self.logger.removeHandler(handler)
            handler.close()
        self.logger.setLevel(self.saved[0])
        self.logger.propagate = self.saved[1]


def _log_detail(args, text: str) -> None:
    """Detalle técnico de un error: en la consola con -v; con --registro, en el archivo."""
    if args.verbose:
        log.exception(text)
    elif getattr(args, "registro", None):
        log.debug(text, exc_info=True)


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
    started = time.monotonic()
    log_file = None
    try:
        if args.registro:
            log_file = _LogFile(args.registro, args.verbose)
            log.info("Exportador EVM %s; Python %s; %s", __version__, sys.version.split()[0], platform.platform())
            log.info("Argumentos: %s", sys.argv[1:] if argv is None else argv)
        code = args.func(args)
        if progress.bar_was_shown():          # consola interactiva: la barra se borró al terminar
            print(f"Tiempo total: {progress.clock_text(time.monotonic() - started)}")
        return code
    except EvmError as exc:                 # errores previstos, con su código del catálogo (docs/codigos.md)
        _log_detail(args, "Error")
        _error(exc.id, str(exc))
        return 2
    except (FileNotFoundError, PermissionError, ValueError, RuntimeError) as exc:
        _log_detail(args, "Error")
        _error("SIS-003", str(exc))
        return 2
    except OSError as exc:  # disco lleno, unidad de red no disponible, ruta no válida...
        _log_detail(args, "Error")
        _error("SIS-002", f"no se pudo leer o escribir un archivo: {exc}")
        return 2
    except KeyboardInterrupt:
        print("Cancelado.", file=sys.stderr)
        return 130
    except Exception as exc:  # errores no previstos: mensaje breve; con -v o --registro, el detalle completo
        _log_detail(args, "Error inesperado")
        if args.registro and not args.verbose:
            print(f"El detalle técnico está en el registro de diagnóstico: {args.registro}", file=sys.stderr)
        elif not args.verbose:
            print("Para ver el detalle técnico se debe ejecutar de nuevo con -v (o con --registro RUTA).",
                  file=sys.stderr)
        _error("SIS-001", f"error inesperado: {type(exc).__name__}: {exc}")
        return 3
    finally:
        if log_file is not None:
            log_file.close()


def _error(id: str, text: str) -> None:
    """Error que detiene el programa: 'ERROR CFG-001: mensaje' en la salida de error."""
    print(f"ERROR {id}: {text}", file=sys.stderr)
