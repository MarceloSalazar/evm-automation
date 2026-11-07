"""Informe de mejoras de validar (--informe) y datos que validar muestra al final: evaluación DCMA de 14 puntos y
comparación de planificaciones.

El informe es un documento Markdown con las observaciones agrupadas por prioridad (alta: las que cambian el libro;
media: las de la calidad del cronograma, que solo revisa validar; baja: las informativas), cada una con la causa y
qué hacer del catálogo de mensajes. No incluye nombres de tareas, de empresas ni de calendarios, ni el título del
proyecto: las tareas se identifican por su código EDT o, sin código, por su UniqueID. Los textos de las observaciones
generales se incluyen sin esos nombres.
"""
from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from evm.config import Config
from evm.fields import fmt_number
from evm.messages import CATALOG, SEV_ERROR, SEV_INFO, SEV_WARN, W_V, Issue
from evm.methods import is_complete
from evm.model import ProjectData, Task
from evm.wbs import KIND_RESIDUAL, Report, ReportPart, ReportRow, code_key, plural

PRIORITY_HIGH, PRIORITY_MEDIUM, PRIORITY_LOW = "Alta", "Media", "Baja"
PRIORITIES = [PRIORITY_HIGH, PRIORITY_MEDIUM, PRIORITY_LOW]
PRIORITY_TEXT = {
    PRIORITY_HIGH: "errores y advertencias que cambian el libro (filas, % planificado, % ganado u horizonte)",
    PRIORITY_MEDIUM: "advertencias de la calidad del cronograma (línea base, avance y lógica), que solo revisa validar",
    PRIORITY_LOW: "observaciones informativas",
}


@dataclass
class Unit:
    """Una revisión de validar: un archivo (o los de [[proyectos]]) con una planificación."""

    label: str                          # planificación revisada ("" = la de la configuración)
    parts: list[ReportPart]
    report: Report | None
    issues: list[Issue]                 # observaciones ya silenciadas y con su alcance
    status: date | None = None          # --fecha-corte (None = la fecha de estado de cada archivo)

    @property
    def title(self) -> str:
        return self.label or (self.parts[0].source.plan if len(self.parts) == 1 else "") or "Proyecto"


def priority(issue: Issue) -> str:
    """Alta: errores, y advertencias de los controles que también hace generar (cambian el libro). Media: advertencias
    de los controles que solo hace validar (línea base, avance y lógica del cronograma). Baja: informativas."""
    if issue.severity == SEV_ERROR:
        return PRIORITY_HIGH
    if issue.severity == SEV_INFO:
        return PRIORITY_LOW
    entry = CATALOG.get(issue.id)
    return PRIORITY_MEDIUM if entry is not None and entry.location == W_V else PRIORITY_HIGH


def _num(v: float, decimals: int = 1) -> str:
    return fmt_number(v, decimals, trim=True)


def _day(d: datetime | date | None) -> str:
    return f"{d:%d/%m/%Y}" if d else "-"


def _active_chain(t: Task) -> bool:
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


# ----------------------------------------------------------------------------
# Datos de cada revisión
# ----------------------------------------------------------------------------
@dataclass
class Facts:
    """Cifras de una revisión (un archivo con una planificación), para el resumen y la comparación."""

    label: str
    tasks: int = 0                      # tareas de la planificación (todas)
    active: int = 0
    inactive: int = 0
    hidden: int = 0                     # tareas activas con Visible = No (no se revisan)
    reviewed: int = 0
    rows: int = 0
    inputs: int = 0
    baselines: dict[int, int] = field(default_factory=dict)
    not_rolled_up: int = 0              # tareas resumen con la línea base sin resumir (FEC-014)
    links_in: int = 0                   # vínculos entre tareas de la planificación
    links_out: int = 0                  # vínculos con tareas de otras planificaciones
    critical: str = "-"                 # % de tareas de detalle sin terminar que son críticas
    baseline_finish: datetime | None = None
    finish: datetime | None = None
    status: date | None = None
    baseline_label: str = ""
    errors: int = 0
    warnings: int = 0
    infos: int = 0


def facts(part: ReportPart, issues: list[Issue], label: str = "", status: date | None = None) -> Facts:
    from evm.validation import plan_members
    project: ProjectData = part.source.project
    scope = part.scope
    members = plan_members(project, scope) if scope is not None else None
    every = [t for t in project.tasks if t.outline_level >= 1 and (members is None or t.uid in members)]
    out = Facts(label or part.source.plan or part.source.name or "Proyecto")
    out.tasks = len(every)
    out.active = sum(1 for t in every if _active_chain(t))
    out.inactive = out.tasks - out.active
    reviewed = list(scope.tasks()) if scope is not None else []
    out.reviewed = len(reviewed)
    out.hidden = max(0, out.active - out.reviewed)
    out.rows = len(part.rows)
    out.inputs = sum(1 for r in part.rows if r.is_input)
    for t in reviewed:
        for n in t.baselines:
            out.baselines[n] = out.baselines.get(n, 0) + 1
    out.baselines = dict(sorted(out.baselines.items()))
    out.not_rolled_up = sum(1 for i in issues if i.id == "FEC-014")
    by_uid = {t.uid: t for t in project.tasks}
    inside = {t.uid for t in every}
    for t in project.tasks:
        for link in t.predecessors:
            p = by_uid.get(link.uid)
            if p is None or not _active_chain(t) or not _active_chain(p):
                continue
            a, b = t.uid in inside, p.uid in inside
            if a and b:
                out.links_in += 1
            elif a or b:
                out.links_out += 1
    leaves = [t for t in reviewed if not t.children]
    open_tasks = [t for t in leaves if not is_complete(t)]
    if open_tasks:
        out.critical = f"{_num(100.0 * sum(1 for t in open_tasks if t.critical) / len(open_tasks))} %"
    bl = [t.baseline_finish for t in leaves if t.has_baseline and t.baseline_finish]
    cur = [t.finish for t in leaves if t.finish]
    out.baseline_finish = max(bl) if bl else None
    out.finish = max(cur) if cur else None
    out.status = status or (project.status_date.date() if project.status_date else None)
    out.baseline_label = project.baseline_label
    out.errors = sum(1 for i in issues if i.severity == SEV_ERROR)
    out.warnings = sum(1 for i in issues if i.severity == SEV_WARN)
    out.infos = sum(1 for i in issues if i.severity == SEV_INFO)
    return out


def facts_table(items: list[Facts]) -> list[tuple[str, list[str]]]:
    """Filas (dato, valor de cada revisión) del resumen y de la comparación de planificaciones."""
    def bl(f: Facts) -> str:
        if not f.baselines:
            return "-"
        return " · ".join(f"{'LB' if n == 0 else f'LB{n}'}: {fmt_number(c, 0)}" for n, c in f.baselines.items())

    n = lambda v: fmt_number(v, 0)      # noqa: E731
    return [
        ("Tareas de la planificación (activas)", [f"{n(f.tasks)} ({n(f.active)})" for f in items]),
        ("Tareas inactivas", [n(f.inactive) for f in items]),
        ("Tareas activas con Visible = No", [n(f.hidden) for f in items]),
        ("Tareas revisadas", [n(f.reviewed) for f in items]),
        ("Filas del libro (de ingreso)", [f"{n(f.rows)} ({n(f.inputs)})" for f in items]),
        ("Tareas con comienzo en cada línea base", [bl(f) for f in items]),
        ("Tareas resumen con la línea base sin resumir", [n(f.not_rolled_up) for f in items]),
        ("Vínculos entre sus tareas", [n(f.links_in) for f in items]),
        ("Vínculos con otras planificaciones", [n(f.links_out) for f in items]),
        ("Tareas sin terminar que son críticas", [f.critical for f in items]),
        ("Fecha de estado", [_day(f.status) for f in items]),
        ("Fin de la línea base / fin previsto", [f"{_day(f.baseline_finish)} / {_day(f.finish)}" for f in items]),
        ("Errores / advertencias / informativas", [f"{f.errors} / {f.warnings} / {f.infos}" for f in items]),
    ]


def _row_dates(row: ReportRow) -> tuple[date | None, date | None, date | None]:
    finishes = [t.finish for t in row.leaves if t.finish] if row.is_input else []
    if not row.is_input:
        stack = list(row.children)
        while stack:
            c = stack.pop()
            finishes += [t.finish for t in c.leaves if t.finish]
            stack += c.children
    return (row.start.date() if row.start else None, row.finish.date() if row.finish else None,
            max(finishes).date() if finishes else None)


def _short(codes: list[str], limit: int = 10) -> str:
    return ", ".join(codes[:limit]) + (f" y {len(codes) - limit} más" if len(codes) > limit else "")


def compare(units: list[Unit]) -> list[str]:
    """Diferencias entre las planificaciones revisadas (una parte cada una): códigos EDT de las filas del libro que
    están solo en una, y filas con el mismo código con otra línea base (fechas de plan) u otro fin previsto."""
    rows = [{code_key(r.code): r for r in u.parts[0].rows if r.kind != KIND_RESIDUAL} for u in units]
    out: list[str] = []
    for k, (u, mine) in enumerate(zip(units, rows)):
        others = set().union(*(set(r) for j, r in enumerate(rows) if j != k))
        only = [mine[c].code for c in mine if c not in others]
        out.append(f"Filas solo en {u.title}: {len(only)}" + (f" ({_short(only)})" if only else ""))
    first, *rest = rows
    for j, other in enumerate(rest, start=1):
        common = [c for c in first if c in other]
        baseline = [first[c].code for c in common if _row_dates(first[c])[:2] != _row_dates(other[c])[:2]]
        forecast = [first[c].code for c in common if _row_dates(first[c])[2] != _row_dates(other[c])[2]]
        pair = f"{units[0].title} y {units[j].title}"
        out.append(f"Filas comunes a {pair}: {len(common)}; con otras fechas de plan (línea base): {len(baseline)}"
                   + (f" ({_short(baseline)})" if baseline else ""))
        out.append(f"Filas comunes a {pair} con otro fin previsto: {len(forecast)}"
                   + (f" ({_short(forecast)})" if forecast else ""))
    return out


# ----------------------------------------------------------------------------
# Sin nombres
# ----------------------------------------------------------------------------
def sanitizer(projects: list[ProjectData], cfg: Config, keep: set[str]) -> Callable[[str], str]:
    """Quita de un texto los nombres de las tareas, del proyecto, de los calendarios de tarea y de las empresas: lo
    que va entre «» y cada nombre que aparece como palabra completa. keep: textos que se conservan (p. ej. los
    códigos de las planificaciones, que pueden ser también el nombre de su tarea)."""
    names = {t.name.strip() for p in projects for t in p.tasks if t.name}
    names |= {p.name.strip() for p in projects if p.name}
    names |= {t.calendar for p in projects for t in p.tasks if t.calendar}
    companies = {}
    if cfg.multi_enabled:
        companies = {cfg.company_a: "empresa A", cfg.company_b: "empresa B", cfg.company_both: "ambas empresas"}
    names = {n for n in names if len(n) >= 4 and n not in keep and n not in companies}
    pattern = (re.compile(r"(?<!\w)(" + "|".join(re.escape(n) for n in sorted(names, key=len, reverse=True))
                          + r")(?!\w)") if names else None)

    def clean(text: str) -> str:
        text = re.sub(r"«[^»]*»", "«…»", text)
        for name, generic in companies.items():
            if name:
                text = re.sub(r"(?<!\w)" + re.escape(name) + r"(?!\w)", generic, text)
        return pattern.sub("…", text) if pattern is not None else text

    return clean


# ----------------------------------------------------------------------------
# Informe Markdown
# ----------------------------------------------------------------------------
def _where(found: list[Issue], limit: int = 20) -> str:
    """Tareas de las observaciones: su código EDT o, sin código, su UniqueID."""
    names = list(dict.fromkeys((i.code or (f"UID {i.uid}" if i.uid is not None else "")) for i in found))
    names = [n for n in names if n]
    return ", ".join(f"`{n}`" for n in names[:limit]) + (f" y {len(names) - limit} más" if len(names) > limit else "")


def _md(text: str) -> str:
    """Texto del catálogo o de una observación para una tabla o una lista de Markdown."""
    return text.replace("|", "\\|").replace("\n", " ")


def _improvements(issues: list[Issue], clean: Callable[[str], str], level: str) -> list[str]:
    out: list[str] = []
    by_priority: dict[str, dict[str, list[Issue]]] = {p: {} for p in PRIORITIES}
    for i in issues:
        by_priority[priority(i)].setdefault(i.id, []).append(i)
    for p in PRIORITIES:
        codes = by_priority[p]
        out += [f"{level} Prioridad {p.lower()}", ""]
        if not codes:
            out += ["Sin observaciones.", ""]
            continue
        out += [f"*{PRIORITY_TEXT[p][0].upper() + PRIORITY_TEXT[p][1:]}.*", ""]
        # Los errores primero; después, por código (cada resumen junto a su detalle: FEC-013 y FEC-014...)
        order = sorted(codes.items(), key=lambda kv: (kv[1][0].severity != SEV_ERROR, kv[0]))
        for code, found in order:
            entry = CATALOG.get(code)
            title = entry.title if entry else code
            tasks = [i for i in found if i.code or i.uid is not None]
            general = [i for i in found if not (i.code or i.uid is not None)]
            out.append(f"{level}# {code} · {title}")
            out.append("")
            severity = found[0].severity
            count = f"{len(tasks)} {'tarea' if len(tasks) == 1 else 'tareas'}" if tasks else "general"
            out.append(f"- **{severity}** · {count}")
            if tasks:
                out.append(f"- **Dónde:** {_where(tasks)}")
            for i in general[:3]:
                out.append(f"- **Observación:** {_md(clean(i.message))}")
            if entry is not None:
                out.append(f"- **Causa:** {_md(entry.cause)}")
                out.append(f"- **Qué hacer:** {_md(entry.action)}")
            out.append("")
    return out


def _dcma_cell(value: tuple[str, str, bool | None] | None) -> str:
    if value is None:
        return "-"
    text, _threshold, ok = value
    mark = "" if ok is None else (" · cumple" if ok else " · **no cumple**")
    return _md(text) + mark


def dcma_rows(parts: list[ReportPart]) -> list[tuple[int, str, list[tuple[str, str, bool | None] | None], str]]:
    """Filas de la evaluación DCMA: (punto, nombre, valor de cada revisión, umbral)."""
    from evm.validation import DCMA_POINTS, DCMA_THRESHOLDS
    return [(k, DCMA_POINTS[k], [p.dcma.get(k) for p in parts], DCMA_THRESHOLDS[k]) for k in sorted(DCMA_POINTS)]


def write_report(path: str | Path, units: list[Unit], cfg: Config, version: str, engine: str) -> None:
    """Escribe el informe de mejoras (Markdown, UTF-8)."""
    projects = [p.source.project for u in units for p in u.parts]
    keep = {u.title for u in units} | {p.source.plan for u in units for p in u.parts if p.source.plan}
    clean = sanitizer(projects, cfg, keep)
    columns = [(u, p) for u in units for p in u.parts]
    labels = [u.title if len(u.parts) == 1 else (p.source.name or u.title) for u, p in columns]
    all_facts = [facts(p, [i for i in u.issues if not i.file or i.file == p.source.name], label)
                 for (u, p), label in zip(columns, labels)]

    lines = ["# Informe de mejoras del cronograma", "",
             "| | |", "|---|---|",
             f"| **Revisión** | `validar` del Exportador EVM {version}, el {date.today():%d/%m/%Y} |",
             f"| **Archivo** | {engine} |",
             f"| **Revisado** | {', '.join(labels)} |",
             f"| **Línea base** | {projects[0].baseline_label} |", "",
             "> El informe no incluye nombres de tareas, de empresas ni de calendarios, ni el título del proyecto: las "
             "tareas se identifican por su código EDT o, sin código, por su UniqueID (UID). Los textos entre «» se "
             "omiten (`«…»`).", "",
             "---", "", "## 1. Resumen", "",
             "| Dato | " + " | ".join(labels) + " |", "|---|" + "---:|" * len(labels)]
    for name, values in facts_table(all_facts):
        lines.append(f"| {name} | " + " | ".join(values) + " |")
    lines.append("")
    top = [(u.title, i) for u in units for i in u.issues if priority(i) == PRIORITY_HIGH]
    counts: dict[tuple[str, str], int] = {}
    for title, i in top:
        counts[(title, i.id)] = counts.get((title, i.id), 0) + 1
    if counts:
        lines += ["**Mejoras de prioridad alta** (las que más observaciones tienen):", ""]
        for (title, code), n in sorted(counts.items(), key=lambda kv: -kv[1])[:8]:
            entry = CATALOG.get(code)
            where = f" en {title}" if len(units) > 1 else ""
            lines.append(f"- `{code}` {entry.title if entry else ''}: {n} {'observación' if n == 1 else 'observaciones'}"
                         f"{where}.")
        lines.append("")

    lines += ["---", "", "## 2. Mejoras por prioridad", "",
              "Cada código se explica en [Códigos de errores y advertencias]"
              "(https://github.com/MarceloSalazar/evm-automation/blob/main/docs/codigos.md).", ""]
    for u in units:
        level = "###"
        if len(units) > 1:
            lines += [f"### {u.title}", ""]
            level = "####"
        lines += _improvements(u.issues, clean, level)

    lines += ["---", "", "## 3. Evaluación DCMA de 14 puntos", "",
              "| Punto | " + " | ".join(labels) + " | Umbral |", "|---|" + "---|" * len(labels) + "---|"]
    for k, name, values, threshold in dcma_rows([p for _u, p in columns]):
        lines.append(f"| {k}. {name} | " + " | ".join(_dcma_cell(v) for v in values) + f" | {threshold} |")
    lines.append("")

    if len(units) > 1:
        lines += ["---", "", "## 4. Comparación de planificaciones", ""]
        lines += [f"- {_md(clean(text))}" for text in compare(units)]
        lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")
