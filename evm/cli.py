"""Interfaz de línea de comandos.

    python evm_export.py generar    --origen proyecto.mpp --salida EVM.xlsx
    python evm_export.py actualizar --origen proyecto.mpp --libro EVM.xlsx --fecha-corte 2026-10-31
"""
from __future__ import annotations

import argparse
import logging
import sys
from contextlib import contextmanager
from datetime import date
from pathlib import Path

from evm import __version__
from evm.config import Config, ConfigError, load_config, parse_date, parse_month
from evm.wbs import SEV_INFO, TimephasedPVSource, build_report

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
    raise ConfigError(f"No se reconoce la extensión '{ext}'. Se debe usar un archivo .mpp (requiere MS Project) "
                      "o .xml (MS Project > Guardar como > XML), o indicar --motor.")


@contextmanager
def open_project(origen: str, cfg: Config, motor: str, visible: bool = False, status: date | None = None):
    """Devuelve (ProjectData, PVSource). Con COM, MS Project queda abierto mientras dure el bloque."""
    engine = _engine_for(origen, motor)
    if origen.lower() != "activo" and not Path(origen).is_file():
        raise FileNotFoundError(f"No existe el archivo de origen: {origen}")
    if engine == "xml":
        from evm.reader_xml import read_xml
        project = read_xml(origen, cfg)
        yield project, TimephasedPVSource(
            lambda t: t.timephased_pv,
            "Costo de línea base mensual del XML de MS Project (datos escalonados)",
            cfg.distribution, None, project.calendar)
        return
    from evm.reader_com import ComSession
    with ComSession(origen, visible=visible) as session:
        project = session.read_project(cfg, status_date=status)
        yield project, session.pv_source(project, cfg)


def _apply_overrides(cfg: Config, args) -> Config:
    if getattr(args, "mes_inicio", None):
        cfg.start_month = parse_month(args.mes_inicio)
    if getattr(args, "meses", None):
        cfg.months = args.meses
    if getattr(args, "linea_base", None) is not None:
        cfg.baseline = args.linea_base
    if getattr(args, "nivel_max", None):
        cfg.max_level = args.nivel_max
    if getattr(args, "planificacion", None) is not None:
        cfg.plan = args.planificacion.strip()
    cfg.validate()
    return cfg


def _print_issues(issues, limit: int = 15) -> None:
    serious = [i for i in issues if i.severity != SEV_INFO]
    infos = [i for i in issues if i.severity == SEV_INFO]
    for i in infos:
        print(f"  [i] {i.message}")
    for i in serious[:limit]:
        where = f"{i.code} {i.name}".strip()
        print(f"  [!] {where + ': ' if where else ''}{i.message}")
    if len(serious) > limit:
        print(f"  ... y {len(serious) - limit} observaciones más (ver hoja 'Validación').")


def cmd_generar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    salida = Path(args.salida) if args.salida else Path(args.origen).with_name(Path(args.origen).stem + "_EVM.xlsx")
    if args.origen.lower() == "activo" and not args.salida:
        salida = Path("EVM.xlsx")
    if salida.suffix.lower() != ".xlsx":
        raise ConfigError("La salida debe ser un archivo .xlsx")
    if salida.exists() and not args.sobrescribir:
        raise ConfigError(f"'{salida}' ya existe. Para reemplazarlo se debe usar --sobrescribir; para conservar "
                          "los EV/AC cargados se debe generar un archivo nuevo con --migrar-desde.")
    if args.migrar_desde and Path(args.migrar_desde).resolve() == salida.resolve():
        raise ConfigError("--migrar-desde debe ser un archivo distinto de --salida.")

    from evm.workbook import write_workbook
    print(f"Leyendo {args.origen} ...")
    with open_project(args.origen, cfg, args.motor, args.visible) as (project, source):
        report = build_report(project, cfg, source)
    prior = None
    if args.migrar_desde:
        from evm.migrate import plan_migration, prior_from_book
        prior = prior_from_book(args.migrar_desde, project.name)
        report.issues.extend(plan_migration(report, prior, cfg.max_level))
    write_workbook(report, cfg, salida, prior)

    leaves = sum(1 for r in report.rows if r.is_input)
    print(f"Proyecto: {project.name}")
    if report.plan:
        print(f"  Planificación: {report.plan}")
    print(f"  {project.baseline_label}  |  PV: {report.pv_source}")
    print(f"  Filas EDT: {len(report.rows)} ({leaves} de ingreso)  |  Nivel máximo: {cfg.max_level}")
    print(f"  Horizonte: {report.horizon.start:%m/%Y} a {report.horizon.end:%m/%Y} ({report.horizon.months} meses)")
    print(f"  BAC total: {report.total_bac:,.2f}")
    _print_issues(report.issues)
    if prior is not None:
        lost_ev = sum(r.total("EV") for r in prior.unmatched)
        lost_ac = sum(r.total("AC") for r in prior.unmatched)
        print(f"Migración desde {args.migrar_desde}: {prior.matched} filas con EV/AC trasladados.")
        if prior.unmatched:
            print(f"  [!] {len(prior.unmatched)} filas no se pudieron trasladar (EV {lost_ev:,.2f}, AC {lost_ac:,.2f}). "
                  "Ver hoja 'Validación'.")
    print(f"Libro generado: {salida.resolve()}")
    return 0


def _print_update(res, currency: str) -> None:
    from evm.indicators import format_table
    from evm.update import indicators_for
    print(f"Fecha de corte: {res.status_date:%d/%m/%Y} (mes {res.month_index + 1} del horizonte)")
    if res.dry_run:
        def m(v):
            return "-" if v is None else f"{v:,.0f}"

        company_cols = sorted({k for r in res.rows for k in r.by_company}, key=lambda k: (k[:2] == "AC", k))
        print(f"{'Código':<10} {'EV del mes':>16} {'EV acumulado':>16} {'AC del mes':>16} {'AC acumulado':>16}"
              + "".join(f" {k:>16}" for k in company_cols))
        for r in res.rows:
            print(f"{r.code:<10} {m(r.ev_month):>16} {m(r.ev_cum):>16} {m(r.ac_month):>16} {m(r.ac_cum):>16}"
                  + "".join(f" {m(r.by_company.get(k)):>16}" for k in company_cols))
    shown = [r for r in res.rows if r.note]
    if shown:
        print("Observaciones por fila:")
        for r in shown[:25]:
            print(f"  {r.code} {r.name[:40]}: {r.note}")
        if len(shown) > 25:
            print(f"  ... y {len(shown) - 25} más")
    for w in res.warnings[:30]:
        print(f"  [!] {w}")
    if len(res.warnings) > 30:
        print(f"  ... y {len(res.warnings) - 30} advertencias más.")
    if res.book is not None and res.book.multi and not res.ac_loaded:
        print(f"  [i] El AC se carga a mano en las hojas 'AC-{res.book.company_a}' y 'AC-{res.book.company_b}'; "
              "los indicadores de costo usan los valores cargados.")
    if res.book is not None:
        print()
        print(format_table(indicators_for(res.book, res.status_date), currency))


def cmd_actualizar(args) -> int:
    cfg = _apply_overrides(load_config(args.config), args)
    status = parse_date(args.fecha_corte) if args.fecha_corte else None
    from evm.update import apply_book_settings, book_meta, check_only, run_update
    # Multi-empresa: se usan los nombres y campos con los que se generó el libro
    cfg = apply_book_settings(cfg, book_meta(args.libro))
    check_only(cfg, args.solo)
    print(f"Leyendo {args.origen} ...")
    with open_project(args.origen, cfg, args.motor, args.visible, status) as (project, _source):
        pass
    res = run_update(project, cfg, args.libro, status, args.solo, backup=not args.sin_respaldo,
                     force=args.forzar, dry_run=args.simular)
    _print_update(res, str(res.book.meta.get("moneda") or "") if res.book else "")
    if args.simular:
        print("Simulación: no se modificó el libro.")
    else:
        if res.backup:
            print(f"Respaldo: {res.backup}")
        print(f"Libro actualizado: {Path(args.libro).resolve()}")
        print("Los indicadores y los gráficos se calculan en el libro (hojas 'Cuadro de Mando', 'Resumen' e "
              "'Indicadores EDT').")
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
        sp.add_argument("--origen", required=True,
                        help="archivo .mpp (requiere MS Project instalado), .xml exportado de MS Project, "
                             "o 'activo' para usar el proyecto abierto en MS Project")
        sp.add_argument("--config", help="archivo de configuración TOML (ver config.ejemplo.toml)")
        sp.add_argument("--motor", choices=["auto", "com", "xml"], default="auto",
                        help="forma de lectura (por defecto según la extensión)")
        sp.add_argument("--linea-base", type=int, choices=range(0, 11), metavar="N",
                        help="línea base a usar: 0 = Línea base, 1..10 = Línea base N")
        sp.add_argument("--nivel-max", type=int, help="nivel máximo del código EDT (por defecto 3)")
        sp.add_argument("--planificacion", metavar="CODIGO",
                        help="código EDT de la planificación a exportar, p. ej. INTERNO o CLIENTE (reemplaza el "
                             "valor del archivo de configuración; \"\" = todas)")
        sp.add_argument("--visible", action="store_true", help="muestra MS Project durante la lectura (COM)")

    g = sub.add_parser("generar", help="crea el libro Excel con el PV mensual a partir de MS Project")
    common_source(g)
    g.add_argument("--salida", help="libro .xlsx a crear (por defecto <origen>_EVM.xlsx)")
    g.add_argument("--mes-inicio", help="primer mes del horizonte, AAAA-MM (por defecto el inicio de la línea base)")
    g.add_argument("--meses", type=int, help="cantidad de meses (por defecto 120)")
    g.add_argument("--migrar-desde", help="libro anterior del que se copian los EV/AC ya cargados")
    g.add_argument("--sobrescribir", action="store_true", help="reemplaza el libro de salida si existe")
    g.set_defaults(func=cmd_generar)

    a = sub.add_parser("actualizar", help="calcula EV y AC a la fecha de corte y los carga en el mes correspondiente")
    common_source(a)
    a.add_argument("--libro", required=True, help="libro .xlsx generado con 'generar'")
    a.add_argument("--fecha-corte", help="fecha de corte AAAA-MM-DD (por defecto la fecha de estado del proyecto; "
                                         "si el proyecto no la tiene, es obligatoria)")
    a.add_argument("--solo", choices=["ev", "ac"], help="actualiza solo EV o solo AC")
    a.add_argument("--forzar", action="store_true", help="permite actualizar aunque haya meses posteriores cargados")
    a.add_argument("--sin-respaldo", action="store_true", help="no crea copia de respaldo del libro")
    a.add_argument("--simular", action="store_true", help="muestra el resultado sin modificar el libro")
    a.set_defaults(func=cmd_actualizar)

    return p


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except AttributeError:  # pragma: no cover
            pass
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format="%(levelname)s %(name)s: %(message)s")
    try:
        return args.func(args)
    except (ConfigError, FileNotFoundError, PermissionError, ValueError, RuntimeError) as exc:
        if args.verbose:
            log.exception("Error")
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Cancelado.", file=sys.stderr)
        return 130
    except Exception as exc:  # errores no previstos: mensaje breve; con -v, el detalle completo
        if args.verbose:
            log.exception("Error inesperado")
        else:
            print("Para ver el detalle técnico se debe ejecutar de nuevo con -v.", file=sys.stderr)
        print(f"ERROR inesperado: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
