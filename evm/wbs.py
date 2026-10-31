"""Construcción de la estructura EDT del informe (hasta el nivel máximo o lo que indique el campo Visible) y
% planificado mensual de cada fila."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import date, datetime

from evm import progress
from evm.config import Config, ConfigError
from evm.fields import fmt_number
from evm.messages import (SEV_ERROR, SEV_INFO, SEV_WARN, TOPIC_BOOK, TOPIC_COMPANY, TOPIC_DATES,  # noqa: F401
                          TOPIC_LOGIC, TOPIC_METHOD, TOPIC_PROGRESS, TOPIC_READ, TOPIC_VISIBLE, TOPIC_WBS,
                          TOPIC_WEIGHTS, TOPICS, Issue)
from evm.methods import is_undefined
from evm.model import MonthKey, ProjectData, Task
from evm.timeline import Horizon, first_of_month

KIND_LEAF = "hoja"          # fila de ingreso de datos (paquete de trabajo / cuenta de control)
KIND_SUMMARY = "resumen"    # fila subtotal (fórmulas)
KIND_RESIDUAL = "residual"  # tareas sin fila propia que cuelgan de una tarea resumen ("Otras tareas")


@dataclass
class ReportRow:
    code: str
    name: str
    level: int
    kind: str
    task: Task
    method: str = ""
    start: datetime | None = None
    finish: datetime | None = None
    parent: "ReportRow | None" = field(default=None, repr=False)
    children: list["ReportRow"] = field(default_factory=list, repr=False)
    vector: list[float] = field(default_factory=list, repr=False)    # % planificado del mes (todas las empresas)
    hidden_tasks: int = 0       # tareas de nivel inferior agregadas en esta fila
    source: str = ""            # varios archivos: nombre del archivo de la tarea ([[proyectos]])
    # Tareas de detalle que forman la fila y % planificado de cada empresa
    # (clave "" en un libro de una sola empresa; "A" y "B" en multi-empresa)
    leaves: list[Task] = field(default_factory=list, repr=False)
    pct: dict[str, dict[MonthKey, float]] = field(default_factory=dict, repr=False)
    pct_vec: dict[str, list[float]] = field(default_factory=dict, repr=False)

    @property
    def is_input(self) -> bool:
        return self.kind in (KIND_LEAF, KIND_RESIDUAL)


@dataclass
class Source:
    """Archivo de MS Project incluido en el informe."""

    name: str                   # nombre del archivo en [[proyectos]]; "" = informe de un solo archivo
    project: ProjectData
    plan: str = ""              # código EDT de la planificación exportada ("" = todo el proyecto)
    pv_source: str = ""         # descripción de la fuente del PV mensual


@dataclass
class Report:
    project: ProjectData        # varios archivos: el primero (los datos de cada uno están en sources)
    horizon: Horizon
    rows: list[ReportRow]
    roots: list[ReportRow]
    issues: list[Issue]
    total_vector: list[float]
    pv_source: str = ""
    plan: str = ""              # código EDT de la planificación exportada ("" = todo el proyecto)
    multi: bool = False         # multi-empresa activado
    company_a: str = ""
    company_b: str = ""
    sources: list[Source] = field(default_factory=list)
    title: str = ""             # título del informe
    currency: str = ""          # moneda de MS Project (se usa si la configuración no indica 'moneda')
    # Depuración de costes: hoja PV que lleva las filas «(coste)» ("A" o "B" en multi-empresa, "" en un libro de una
    # sola empresa); None = sin depuración
    cost_review: str | None = None

    @property
    def multi_file(self) -> bool:
        """True si el informe combina archivos de [[proyectos]] (las claves de las filas llevan su nombre)."""
        return any(s.name for s in self.sources)


@dataclass
class ReportPart:
    """Filas de un archivo de MS Project, antes de combinar los archivos y fijar el horizonte."""

    source: Source
    rows: list[ReportRow]
    roots: list[ReportRow]
    issues: list[Issue]         # observaciones de la lectura y de la estructura EDT
    checks: list[Issue]         # controles finales (tareas omitidas, BAC, códigos duplicados...)
    first_date: date | None     # inicio del horizonte si ninguna fila tiene % planificado
    scope: "Scope | None" = None                # tareas exportadas (sin las inactivas ni las de Visible = No)
    code_of: Callable[[Task], str] | None = None    # código EDT de una tarea, como en el libro
    list_matched: set[str] | None = None        # lista EDT: códigos de la lista que tienen tarea (None = sin lista)


# ----------------------------------------------------------------------------
# Árbol de tareas
# ----------------------------------------------------------------------------
def build_tree(tasks: list[Task]) -> list[Task]:
    """Arma la jerarquía a partir del orden de las tareas y su nivel de esquema."""
    roots: list[Task] = []
    stack: list[Task] = []
    for t in tasks:
        t.parent = None
        t.children = []
        while stack and stack[-1].outline_level >= t.outline_level:
            stack.pop()
        if stack:
            t.parent = stack[-1]
            stack[-1].children.append(t)
        else:
            roots.append(t)
        stack.append(t)
    return roots


def iter_subtree(task: Task):
    yield task
    for c in task.children:
        yield from iter_subtree(c)


def prune_inactive(roots: list[Task]) -> tuple[list[Task], dict[int, list[Task]], int]:
    """Quita las tareas inactivas (con toda su rama).

    Devuelve (raíces activas, {uid del padre: hijos excluidos}, cantidad de tareas excluidas).
    """
    removed_children: dict[int, list[Task]] = {}
    count = 0

    def prune(nodes: list[Task], parent: Task | None) -> list[Task]:
        nonlocal count
        kept = []
        for n in nodes:
            if not n.active:
                count += sum(1 for _ in iter_subtree(n))
                if parent is not None:
                    removed_children.setdefault(parent.uid, []).append(n)
                continue
            n.children = prune(n.children, n)
            kept.append(n)
        return kept

    return prune(roots, None), removed_children, count


def pad_code(code: str, levels: int) -> str:
    """Completa un código EDT numérico con ceros: '1' -> '1.0.0', '1.2' -> '1.2.0'.

    Los códigos no numéricos (p. ej. 'INTERNO') no se modifican.
    """
    code = (code or "").strip()
    if not code:
        return code
    parts = code.split(".")
    if not all(p.strip().isdigit() for p in parts) or len(parts) >= levels:
        return code
    return ".".join(parts + ["0"] * (levels - len(parts)))


def plural(n: int, singular: str, plural_form: str) -> str:
    """'1 tarea', '3 tareas'."""
    return f"{n} {singular if n == 1 else plural_form}"


def task_name(task: Task, code_of) -> str:
    """Nombre de la tarea para las observaciones; sin código EDT, con su Id de MS Project para encontrarla."""
    return task.name if code_of(task) else f"{task.name} (Id {task.id})"


def fmt_money(v: float) -> str:
    return fmt_number(v)


# ----------------------------------------------------------------------------
# Códigos EDT y planificación
# ----------------------------------------------------------------------------
def _code_parts(code: str) -> list[str]:
    """Segmentos de un código EDT sin los ceros finales (relleno): '1.2.0.1.0' -> ['1', '2', '0', '1']."""
    parts = [p.strip() for p in (code or "").strip().split(".")]
    while len(parts) > 1 and not parts[-1].strip("0"):
        parts.pop()
    return parts


def code_level(code: str, ignore_interior_zeros: bool = False) -> int:
    """Nivel de un código EDT: cantidad de segmentos sin contar los ceros finales.

    '2.0.0' -> 1, '2.1.0' -> 2, '2.1.1' -> 3, '2.1.1.4' -> 4, 'INTERNO' -> 1. Un 0 intermedio cuenta como un
    nivel sin elemento: '1.2.0.1' -> 4 (actividad de 1.2). Con ignore_interior_zeros ([exportacion]
    ignorar_ceros_intermedios) tampoco cuentan los ceros intermedios: '1.2.0.1' -> 3 (hijo de 1.2, como 1.2.1.0).
    """
    parts = _code_parts(code)
    if ignore_interior_zeros and all(p.isdigit() for p in parts):
        return 1 + sum(1 for p in parts[1:] if p.strip("0"))
    return len(parts)


def has_interior_zero(code: str) -> bool:
    """El código EDT numérico tiene un 0 intermedio (su nivel depende de ignorar_ceros_intermedios)."""
    parts = _code_parts(code)
    return all(p.isdigit() for p in parts) and any(not p.strip("0") for p in parts[1:])


def code_key(code: str) -> str:
    """Código EDT para compararlo con otro: sin los ceros finales, que son relleno ('1.2', '1.2.0' y '1.2.0.0' ->
    '1.2'). Los ceros intermedios cuentan ('1.2.0.1' y '1.2.1' son distintos). Una fila -D conserva el sufijo."""
    code = (code or "").strip()
    if code.endswith("-D"):
        return code_key(code[:-2]) + "-D"
    return ".".join(_code_parts(code))


class PlanNotFound(ConfigError):
    """La planificación indicada en la configuración no existe en el archivo."""


@dataclass
class Scope:
    """Tareas a exportar: todo el proyecto o la planificación elegida (p. ej. INTERNO)."""

    roots: list[Task]
    plan: Task | None = None
    prefix: str = ""            # código EDT de la planificación

    def raw_code(self, task: Task) -> str:
        """Código EDT de la tarea, sin el prefijo de la planificación si lo tuviera."""
        return strip_prefix(task.wbs or "", self.prefix)

    def tasks(self):
        for r in self.roots:
            yield from iter_subtree(r)


def strip_prefix(code: str, prefix: str) -> str:
    """Código EDT sin el prefijo de la planificación ('INTERNO.1.2.0' -> '1.2.0'), si lo tiene."""
    code = (code or "").strip()
    p = prefix
    if p and len(code) > len(p) + 1 and code[:len(p)].upper() == p.upper() and code[len(p)] in ".-_ /":
        code = code[len(p) + 1:].strip()
    return code


def select_scope(roots: list[Task], plan_code: str) -> Scope:
    """Elige la planificación por su código EDT (o, si no lo hay, por el nombre de la tarea)."""
    plan_code = (plan_code or "").strip()
    if not plan_code:
        return Scope(list(roots))
    from evm.fields import same_name
    every = [t for r in roots for t in iter_subtree(r)]
    found = sorted((t for t in every if same_name(t.wbs or "", plan_code)), key=lambda t: t.outline_level)
    if not found:
        found = sorted((t for t in every if same_name(t.name or "", plan_code)), key=lambda t: t.outline_level)
    if not found:
        tops = [f"{(t.wbs or '').strip()} ({t.name})" if (t.wbs or "").strip() else t.name for t in roots]
        raise PlanNotFound(
            "EDT-010", f"No se encontró la planificación '{plan_code}': ninguna tarea tiene ese código EDT. Tareas "
            f"de nivel superior del proyecto: {', '.join(tops[:15]) or 'ninguna'}.")
    plan = found[0]
    return Scope(list(plan.children), plan, (plan.wbs or "").strip() or plan_code)


@dataclass
class Visibility:
    """Filas del libro según el código EDT, el nivel máximo y el campo Visible."""

    rows: dict[int, bool]               # UniqueID -> tiene fila propia
    issues: list[Issue]
    shown: int = 0                      # tareas con Visible = Sí y fila propia por debajo del nivel máximo
    # Visible = No: ramas que no forman parte del PV, el EV ni el AC
    excluded: list[Task] = field(default_factory=list)
    # Tareas cuyas subtareas tienen todas Visible = No: la tarea las representa (con toda su rama)
    represented: list[Task] = field(default_factory=list)
    listed: bool = False                # las filas las decide la lista EDT (lista_edt)
    matched: set[str] = field(default_factory=set)     # códigos de la lista EDT que tienen tarea

    def has_row(self, task: Task) -> bool:
        return self.rows.get(task.uid, False)

    def excluded_uids(self) -> set[int]:
        """Tareas de las ramas excluidas (con sus subtareas)."""
        return {n.uid for t in self.excluded for n in iter_subtree(t)}


def row_visibility(scope: Scope, cfg: Config, code_of=None) -> Visibility:
    """Decide qué tareas tienen fila propia en el libro y cuáles quedan fuera del cálculo.

    - Visible vacío: fila propia si la tarea tiene código EDT de nivel no mayor que el nivel máximo.
    - Visible = Sí: fila propia aunque su código sea más detallado que el nivel máximo (necesita código EDT).
    - Visible = No: la tarea y su rama no forman parte del PV, el EV ni el AC. Si todas las subtareas de una tarea
      tienen Visible = No, esa tarea las representa: no tienen fila propia, pero su rama se suma en la fila de la
      tarea (que, si tiene fila, pasa a ser el paquete de trabajo). Prevalece sobre un Sí en la rama.
    """
    if cfg.wbs_list is not None:
        return list_visibility(scope, cfg)
    from evm.fields import parse_visible
    code_of = code_of or scope.raw_code
    out = Visibility({}, [])
    values: dict[int, bool | None] = {}
    if cfg.visible_field:               # visible = "" en la configuración: el campo no se usa
        for t in scope.tasks():
            value, valid = parse_visible(t.visible)
            values[t.uid] = value
            if not valid:
                out.issues.append(Issue("VIS-001", code_of(t), t.name,
                                        f"Visible = '{t.visible}' no se reconoce (se admite Sí o No): se usa el "
                                        "nivel máximo."))

    def hidden(t: Task) -> bool:
        return values.get(t.uid) is False

    def label(t: Task) -> str:
        return f"{code_of(t) or '(sin código)'} «{t.name}»"

    def drop(t: Task) -> None:
        """Rama excluida: ninguna tarea tiene fila; se avisa de los Sí que quedan fuera."""
        out.excluded.append(t)
        for n in iter_subtree(t):
            out.rows[n.uid] = False
            if n is not t and values.get(n.uid) is True:
                out.issues.append(Issue("VIS-002", code_of(n), n.name,
                                        f"Visible = Sí, pero depende de {label(t)}, con Visible = No: queda fuera "
                                        "del cálculo con esa rama."))

    def walk(group: list[Task], parent: Task | None, hidden_by: Task | None) -> None:
        every = bool(group) and all(hidden(t) for t in group)
        if every and parent is not None:
            out.represented.append(parent)
        for t in group:
            raw = scope.raw_code(t)
            value = values.get(t.uid)
            if value is False and not every:
                drop(t)
                continue
            inner = hidden_by or (t if value is False else None)
            if value is True and hidden_by is not None:
                out.issues.append(Issue("VIS-003", code_of(t), t.name,
                                        f"Visible = Sí, pero depende de {label(hidden_by)}, con Visible = No: se "
                                        "suma con su rama."))
            elif value is True and not raw:
                out.issues.append(Issue("VIS-004", code_of(t), t.name,
                                        "Visible = Sí en una tarea sin código EDT: no puede tener fila propia; se "
                                        "suma en su fila superior."))
            level = code_level(raw, cfg.ignore_interior_zeros)
            has_row = bool(raw) and inner is None and (value is True or level <= cfg.max_level)
            if has_row and level > cfg.max_level:
                out.shown += 1
            out.rows[t.uid] = has_row
            walk(t.children, t, inner)

    roots = scope.roots
    if roots and all(hidden(t) for t in roots):
        out.issues.append(Issue("VIS-005", "", "",
                                "Todos los elementos de nivel superior tienen Visible = No: se ignora en ellos (el "
                                "libro no tendría filas)."))
        for t in roots:
            values[t.uid] = None
    walk(roots, None, None)
    return out


def list_visibility(scope: Scope, cfg: Config) -> Visibility:
    """Filas según la lista EDT (lista_edt): tiene fila cada tarea cuyo código está en la lista. Las demás se suman
    en su elemento listado más cercano. La lista reemplaza al nivel máximo y al campo Visible. Los códigos se
    comparan sin el prefijo de la planificación ni los ceros finales (1.2.0 en la lista es la tarea 1.2.0.0)."""
    def book_code(t: Task) -> str:
        raw = scope.raw_code(t)
        return pad_code(raw, cfg.max_level) if cfg.pad_wbs else raw

    wanted: dict[str, list[str]] = {}           # código sin los ceros finales -> líneas de la lista
    for entry in cfg.wbs_list or []:
        wanted.setdefault(code_key(strip_prefix(entry, scope.prefix)), []).append(entry)
    out = Visibility({}, [], listed=True)
    tasks = list(scope.tasks())
    for t in tasks:
        code = code_key(scope.raw_code(t))
        listed = bool(code) and code in wanted
        out.rows[t.uid] = listed
        if listed:
            out.matched.update(wanted[code])
    with_value = sum(1 for t in tasks if (t.visible or "").strip()) if cfg.visible_field else 0
    if with_value:
        out.issues.append(Issue("VIS-006", "", "",
                                f"{plural(with_value, 'tarea tiene', 'tareas tienen')} valor en el campo Visible: con "
                                "lista_edt, las filas del libro las decide la lista EDT y el campo Visible no se "
                                "usa."))
    # Elementos listados cuya tarea superior tiene código EDT y no está en la lista
    in_scope = {t.uid for t in tasks}
    for t in tasks:
        parent = t.parent
        if not out.rows[t.uid] or parent is None or parent.uid not in in_scope or out.rows.get(parent.uid):
            continue
        if not book_code(parent):
            continue
        above = parent.parent
        while above is not None and above.uid in in_scope and not out.rows.get(above.uid):
            above = above.parent
        where = (f"queda debajo de {book_code(above)} «{above.name}»" if above is not None and above.uid in in_scope
                 else "queda en el primer nivel del libro")
        out.issues.append(Issue("EDT-015", book_code(t), t.name,
                                f"Está en la lista EDT y su tarea superior {book_code(parent)} «{parent.name}» no: la "
                                f"fila {where}."))
    return out


def prune_hidden(scope: Scope, visibility: Visibility) -> dict[int, list[Task]]:
    """Quita del árbol las ramas excluidas por Visible = No (como las tareas inactivas). Devuelve
    {UniqueID de la tarea superior: ramas quitadas}."""
    cut: dict[int, list[Task]] = {}
    out_ids = {t.uid for t in visibility.excluded}
    if not out_ids:
        return cut
    scope.roots = [r for r in scope.roots if r.uid not in out_ids]
    stack = list(scope.roots)
    while stack:
        t = stack.pop()
        removed = [c for c in t.children if c.uid in out_ids]
        if removed:
            cut[t.uid] = removed
            t.children = [c for c in t.children if c.uid not in out_ids]
        stack.extend(t.children)
    return cut


# ----------------------------------------------------------------------------
# % planificado
# ----------------------------------------------------------------------------
PERCENT_PV_SOURCE = ("% planificado de MS Project (duración de la línea base de las tareas y peso de los hitos) "
                     "× BAC ingresado en el libro")


def _has_weight(t: Task) -> bool:
    """Tarea que puede aportar % planificado: tarea con fechas o hito con peso."""
    if t.milestone:
        return (t.milestone_weight or 0) > 0
    return bool((t.baseline_start or t.start) and (t.baseline_finish or t.finish))


def _percent_curves(project: ProjectData, cfg: Config, rows: list[ReportRow], code_of, share) -> list[Issue]:
    """% planificado de cada fila de ingreso (y de cada empresa) y sus observaciones."""
    from evm import percent as pct
    issues: list[Issue] = []
    calendar = pct.calendar_for(project, cfg)
    cache: dict[int, pct.TaskTime] = {}
    alone: dict[str, list[str]] = {}            # empresa -> filas en las que no participa
    if cfg.multi_enabled:
        roles = [(pct.ROLE_A, cfg.company_a, share), (pct.ROLE_B, cfg.company_b, lambda t: 1.0 - share(t))]
    else:
        roles = [(pct.ROLE_ALL, "", None)]
    for i, row in enumerate(rows, start=1):
        progress.step(i, len(rows), "filas")
        if not row.is_input:
            continue
        times = [cache.setdefault(t.uid, pct.task_time(t, calendar, cfg.distribution)) for t in row.leaves]
        full = pct.full_weights(times)
        for role, company, fn in roles:
            rw = pct.company_weights(times, fn, company, full)
            if rw.fallback:
                alone.setdefault(company, []).append(row.code)
            row.pct[role] = pct.curve(times, rw)
            issues += [replace(note, code=row.code, name=row.name) for note in rw.notes]
        if not any(row.pct.values()):
            issues.append(Issue("HIT-005", row.code, row.name,
                                "Ninguna tarea de la fila tiene fechas ni es un hito con peso: no tiene % "
                                "planificado y su BAC no se puede repartir en el tiempo."))

    leaves = list(cache.values())
    undated = [x for x in leaves if not x.dist]
    current = [x for x in leaves if x.dist and not x.from_baseline] if project.baseline_found else []
    if current:
        issues.append(Issue("FEC-010", "", "",
                            f"{plural(len(current), 'tarea no tiene', 'tareas no tienen')} fechas de línea base "
                            f"(p. ej. se {'agregó' if len(current) == 1 else 'agregaron'} después de guardarla): "
                            "su % planificado se calculó con sus fechas actuales. Si corresponde, se debe "
                            "actualizar la línea base de esas tareas."))
        for x in current[:200]:
            issues.append(Issue("FEC-011", code_of(x.task), task_name(x.task, code_of),
                                "Sin fechas de línea base: se usaron sus fechas actuales."))
    for x in undated[:200]:
        issues.append(Issue("FEC-012", code_of(x.task), task_name(x.task, code_of),
                            "Sin fechas de línea base ni fechas actuales: no aporta % planificado."))
    for company, codes in alone.items():
        shown = ", ".join(codes[:10]) + (f" y {len(codes) - 10} más" if len(codes) > 10 else "")
        issues.append(Issue("EMP-002", "", "",
                            f"{company} no participa en las tareas de {plural(len(codes), 'fila', 'filas')} "
                            f"({shown}): si se le asigna BAC, su PV se reparte con el % planificado de la fila "
                            "completa."))
    return issues


# ----------------------------------------------------------------------------
# Informe
# ----------------------------------------------------------------------------
def build_report(project: ProjectData, cfg: Config) -> Report:
    """Informe de un solo archivo de MS Project."""
    return assemble_report([build_part(project, cfg)], cfg)


def build_part(project: ProjectData, cfg: Config, name: str = "") -> ReportPart:
    """Filas EDT y % planificado (por mes calendario) de un archivo de MS Project.

    name: nombre del archivo en [[proyectos]] ("" = informe de un solo archivo).
    """
    from evm.validation import data_checks, wbs_checks
    issues: list[Issue] = list(project.warnings)
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)

    # Tareas inactivas (MS Project 2010+): se excluyen con toda su rama
    if not cfg.include_inactive:
        roots, _, removed = prune_inactive(roots)
        if removed:
            issues.append(Issue("EDT-016", "", "",
                                f"Se {'excluyó' if removed == 1 else 'excluyeron'} "
                                f"{plural(removed, 'tarea inactiva', 'tareas inactivas')}: no "
                                f"{'forma' if removed == 1 else 'forman'} parte del % planificado."))

    # Planificación (p. ej. INTERNO o CLIENTE): solo las tareas que dependen de ella
    scope = select_scope(roots, cfg.plan)
    if scope.plan is not None:
        issues.append(Issue("EDT-017", scope.prefix, scope.plan.name,
                            f"Planificación '{scope.prefix}': se exportan las {sum(1 for _ in scope.tasks())} "
                            f"tareas que dependen de '{scope.plan.name}'."))

    if not project.baseline_found:
        issues.append(Issue(
            "FEC-009", "", "",
            "El proyecto no tiene guardada la línea base seleccionada: el % planificado se calculó con las "
            "fechas ACTUALES de las tareas. Se debe guardar la línea base en MS Project (Proyecto > Establecer "
            "línea base) y volver a exportar."))

    def code_of(t: Task) -> str:
        raw = scope.raw_code(t)
        return pad_code(raw, cfg.max_level) if cfg.pad_wbs else raw

    # Códigos EDT de todas las tareas exportadas, también las de las ramas con Visible = No
    wbs_issues = wbs_checks(scope, code_of, cfg)

    # Fila propia: código EDT de nivel no mayor que el nivel máximo, salvo lo que indique el campo Visible. Las ramas
    # con Visible = No se excluyen del cálculo, como las tareas inactivas
    visibility = row_visibility(scope, cfg, code_of)
    issues += visibility.issues
    included = visibility.has_row
    prune_hidden(scope, visibility)

    # Multi-empresa: proporción de la empresa A en cada tarea
    shares: dict[int, float] = {}
    company_notes: dict[str, list[Task]] = {}
    if cfg.multi_enabled:
        from evm.companies import resolve_shares
        shares, company_notes = resolve_shares(scope.roots, cfg)

    def share(t: Task) -> float:
        return shares.get(t.uid, 0.5)

    # --- Filas del informe -----------------------------------------------------
    rows: list[ReportRow] = []
    report_roots: list[ReportRow] = []

    has_included: dict[int, bool] = {}

    def mark(t: Task) -> bool:
        found = False
        for c in t.children:
            found = mark(c) or included(c) or found
        has_included[t.uid] = found
        return found

    for r in scope.roots:
        mark(r)

    def subtree_dates(t: Task) -> tuple[datetime | None, datetime | None]:
        if t.baseline_start or t.baseline_finish:
            return t.baseline_start, t.baseline_finish
        starts = [n.baseline_start for n in iter_subtree(t) if n.baseline_start]
        finishes = [n.baseline_finish for n in iter_subtree(t) if n.baseline_finish]
        return (min(starts) if starts else None, max(finishes) if finishes else None)

    def gather(node: Task) -> tuple[list[Task], list[Task], list[Task]]:
        """Recorre los hijos hasta encontrar tareas con fila propia.

        Devuelve (tareas con fila, ramas completas sin fila, tareas intermedias sin fila).
        De las intermedias se recorren sus hijos.
        """
        with_row: list[Task] = []
        loose_full: list[Task] = []
        loose_own: list[Task] = []
        for c in node.children:
            if included(c):
                with_row.append(c)
            elif has_included[c.uid]:
                loose_own.append(c)
                a, b, d = gather(c)
                with_row += a
                loose_full += b
                loose_own += d
            else:
                loose_full.append(c)
        return with_row, loose_full, loose_own

    aggregated = 0      # tareas sin fila propia sumadas en una fila
    beyond = "que no están en la lista EDT" if visibility.listed else f"de nivel mayor a {cfg.max_level}"

    def visit(t: Task, parent: ReportRow | None) -> None:
        nonlocal aggregated
        s, f = subtree_dates(t)
        row_tasks, loose_full, loose_own = gather(t)
        # En el nivel máximo la fila suma toda su rama, salvo que una tarea de la rama tenga Visible = Sí
        is_leaf_row = not row_tasks or (not visibility.listed
                                        and code_level(scope.raw_code(t), cfg.ignore_interior_zeros) >= cfg.max_level
                                        and not any(code_level(scope.raw_code(c), cfg.ignore_interior_zeros)
                                                    > cfg.max_level for c in row_tasks))
        row = ReportRow(code=code_of(t), name=t.name, level=(parent.level + 1) if parent else 1,
                        kind=KIND_LEAF if is_leaf_row else KIND_SUMMARY, task=t,
                        start=s, finish=f, parent=parent)
        rows.append(row)
        if parent:
            parent.children.append(row)
        else:
            report_roots.append(row)

        if is_leaf_row:
            row.leaves = [n for n in iter_subtree(t) if not n.children]
            row.method = (t.method or "").strip() or cfg.default_method
            row.hidden_tasks = sum(1 for _ in iter_subtree(t)) - 1
            aggregated += row.hidden_tasks
            return

        for c in row_tasks:
            visit(c, row)
        # Tareas sin código EDT (o de nivel mayor al máximo) que cuelgan de la tarea resumen: forman la fila -D,
        # con su propio BAC, si tienen duración o son hitos con peso.
        loose = [n for b in loose_full for n in iter_subtree(b) if not n.children]
        aggregated += sum(1 for b in loose_full for _ in iter_subtree(b)) + len(loose_own)
        if any(_has_weight(n) for n in loose):
            # Fechas de la línea base de sus propias tareas (no las de toda la tarea resumen)
            starts = [n.baseline_start for n in loose if n.baseline_start]
            finishes = [n.baseline_finish for n in loose if n.baseline_finish]
            res = ReportRow(code=f"{row.code}-D", name=f"Otras tareas de «{t.name}»",
                            level=row.level + 1, kind=KIND_RESIDUAL, task=t, method=cfg.default_method,
                            start=min(starts) if starts else s, finish=max(finishes) if finishes else f, parent=row)
            res.leaves = loose
            rows.append(res)
            row.children.append(res)
            issues.append(Issue("EDT-018", row.code, t.name,
                                f"Tareas sin código EDT o {beyond}: forman la fila {row.code}-D, con su propio BAC."))

    omitted: list[Task] = []

    def top(nodes: list[Task]) -> None:
        for n in nodes:
            if included(n):
                visit(n, None)
            elif has_included[n.uid]:
                omitted.append(n)           # contenedor sin código: se omite, sus hijos se recorren
                top(n.children)
            else:
                omitted.extend(iter_subtree(n))

    top(scope.roots)

    issues += _percent_curves(project, cfg, rows, code_of, share)

    if cfg.multi_enabled:
        # Avisos de empresa, solo de las tareas que forman parte del libro
        omitted_ids = {n.uid for n in omitted}
        notes = {k: [t for t in v if t.uid not in omitted_ids] for k, v in company_notes.items()}
        missing = notes["sin_empresa"]
        missing_ids = {t.uid for t in missing}
        top_missing = [t for t in missing if t.parent is None or t.parent.uid not in missing_ids]
        if top_missing:
            one = len(missing) == 1
            issues.append(Issue("EMP-003", "", "",
                                f"{plural(len(missing), 'tarea', 'tareas')} sin empresa en la columna "
                                f"'{cfg.company_field}' (ni en sus tareas superiores): se "
                                f"{'trata' if one else 'tratan'} como '{cfg.company_both}' "
                                f"({fmt_number(cfg.split_default, trim=True)} % para {cfg.company_a})."))
            for t in top_missing[:50]:
                issues.append(Issue("EMP-004", code_of(t), task_name(t, code_of),
                                    f"Sin empresa: se trata como '{cfg.company_both}'"
                                    + (" (incluye sus subtareas sin empresa)." if t.children else ".")))
        for t in notes["desconocida"][:50]:
            issues.append(Issue("EMP-005", code_of(t), task_name(t, code_of),
                                f"Empresa '{t.company}' no reconocida (se esperaba '{cfg.company_a}', "
                                f"'{cfg.company_b}' o '{cfg.company_both}'): se trata como '{cfg.company_both}'."))
        for t in notes["reparto"][:50]:
            issues.append(Issue("EMP-006", code_of(t), task_name(t, code_of),
                                f"'{cfg.split_field}' = {fmt_number(t.split_a, trim=True)} fuera de rango (debe ser "
                                "mayor que 0 y como máximo 100): se usa "
                                f"{fmt_number(cfg.split_default, trim=True)} %."))

    # --- Controles del archivo ------------------------------------------------
    checks: list[Issue] = []
    # Es una advertencia si las tareas omitidas aportarían % planificado
    omitted_sev = SEV_WARN if any(_has_weight(n) for n in omitted if not n.children) else SEV_INFO
    if omitted and visibility.listed:
        one = len(omitted) == 1
        out_ids = {n.uid for n in omitted}
        coded = [n for n in omitted if scope.raw_code(n) and (n.parent is None or n.parent.uid not in out_ids)]
        shown = ", ".join(code_of(n) for n in coded[:10]) + (f" y {len(coded) - 10} más" if len(coded) > 10 else "")
        checks.append(Issue("EDT-019", "", "",
                            f"Se {'omitió' if one else 'omitieron'} {plural(len(omitted), 'tarea', 'tareas')} que no "
                            f"{'está' if one else 'están'} en la lista EDT ni {'depende' if one else 'dependen'} de "
                            f"un elemento de la lista: no {'forma' if one else 'forman'} parte de ninguna fila del "
                            f"libro ni del PV y el EV" + (f" ({shown})." if shown else "."), omitted_sev))
    elif omitted:
        one = len(omitted) == 1
        checks.append(Issue("EDT-020", "", "",
                            f"Se {'omitió' if one else 'omitieron'} {plural(len(omitted), 'tarea', 'tareas')} sin "
                            f"código EDT que no depende{'' if one else 'n'} de ningún elemento EDT: no "
                            f"forma{'' if one else 'n'} parte de ninguna fila del libro.", omitted_sev))
    if not rows:
        checks.append(Issue("EDT-021", "", "",
                            "Ninguna tarea tiene código EDT: el libro no tiene filas. Se debe revisar el campo EDT "
                            "configurado ([campos] edt) y la planificación elegida."))
    checks += wbs_issues
    checks += data_checks(project, scope, rows, cfg, code_of)

    if aggregated:
        beyond_edt = "que no están en la lista EDT" if visibility.listed else f"de nivel EDT mayor a {cfg.max_level}"
        checks.append(Issue("EDT-022", "", "",
                            f"{plural(aggregated, 'tarea', 'tareas')} sin código EDT o {beyond_edt} se "
                            f"{'agregó' if aggregated == 1 else 'agregaron'} en su elemento EDT superior."))
    if visibility.excluded:
        n = len(visibility.excluded)
        names = ", ".join(code_of(t) or t.name for t in visibility.excluded[:10])
        more = f" y {n - 10} más" if n > 10 else ""
        detail = (f"no {'forma' if n == 1 else 'forman'} parte del % planificado ni del % ganado; el BAC y el AC "
                  f"que se ingresan en el libro no {'la' if n == 1 else 'las'} deben incluir")
        checks.append(Issue("VIS-007", "", "",
                            f"{plural(n, 'tarea', 'tareas')} con Visible = No (con sus subtareas) "
                            f"{'queda' if n == 1 else 'quedan'} fuera del PV, el EV y el AC ({names}{more}): "
                            f"{detail}."))
        for t in visibility.excluded[:200]:
            checks.append(Issue("VIS-008", code_of(t), t.name, "Visible = No: fuera del PV, el EV y el AC."))
    for t in visibility.represented:
        checks.append(Issue("VIS-009", code_of(t), t.name,
                            "Todas sus subtareas tienen Visible = No: "
                            + ("la fila es el paquete de trabajo, con toda su rama." if included(t)
                               else "se suman con la tarea en su fila superior.")))
    if visibility.shown:
        checks.append(Issue("VIS-010", "", "",
                            f"{plural(visibility.shown, 'tarea tiene', 'tareas tienen')} Visible = Sí y fila propia "
                            f"por debajo del nivel máximo ({cfg.max_level})."))
    for row in rows:
        row.source = name
    dates = [t.baseline_start or t.start for t in scope.tasks() if (t.baseline_start or t.start)]
    ref = min(dates) if dates else project.project_start
    first_date = (ref.date() if isinstance(ref, datetime) else ref) if ref else None
    return ReportPart(Source(name, project, scope.prefix, PERCENT_PV_SOURCE), rows, report_roots, issues, checks,
                      first_date, scope, code_of, visibility.matched if visibility.listed else None)


def list_issues(parts: list[ReportPart], cfg: Config) -> list[Issue]:
    """Lista EDT: códigos de la lista que no tienen tarea en ningún archivo (no tienen fila en el libro). Si no
    existe ninguno, el libro no tendría filas: es un error de configuración."""
    if cfg.wbs_list is None:
        return []
    matched = set().union(*(p.list_matched or set() for p in parts))
    missing = [e for e in cfg.wbs_list if e not in matched]
    if not missing:
        return []
    if len(missing) == len(cfg.wbs_list):
        raise ConfigError("EDT-012", f"Ningún código de la lista EDT ({cfg.wbs_list_file}) existe en MS Project: "
                                     "se deben revisar la lista, el campo EDT ([campos] edt) y la planificación "
                                     "elegida.")
    one = len(missing) == 1
    shown = ", ".join(missing[:15]) + (f" y {len(missing) - 15} más" if len(missing) > 15 else "")
    out = [Issue("EDT-013", "", "",
                 f"{plural(len(missing), 'código', 'códigos')} de la lista EDT no {'existe' if one else 'existen'} en "
                 f"MS Project, o {'su tarea está' if one else 'sus tareas están'} inactiva{'' if one else 's'} o fuera "
                 f"de la planificación elegida: no {'tiene' if one else 'tienen'} fila en el libro ({shown}).")]
    out += [Issue("EDT-014", code, "", "Código de la lista EDT sin tarea en MS Project: no tiene fila en el libro.")
            for code in missing[:200]]
    return out


def _code_order(row: ReportRow) -> tuple:
    """Orden de los elementos de nivel superior de varios archivos: por código EDT numérico (1.0.0 < 9.0.0 <
    10.0.0); los códigos no numéricos van al final, en su orden."""
    segments = row.code.split(".")
    if all(x.strip().isdigit() for x in segments):
        return (0, tuple(int(x) for x in segments))
    return (1, ())


def _block(row: ReportRow):
    """La fila y sus descendientes, en el orden del libro."""
    yield row
    for c in row.children:
        yield from _block(c)


def tagged_issues(part: ReportPart, issues: list[Issue]) -> list[Issue]:
    """Con varios archivos, cada observación indica el archivo: '[P2] ...'."""
    name = part.source.name
    if not name:
        return list(issues)
    return [replace(i, message=f"[{name}] {i.message}") for i in issues]


def assemble_report(parts: list[ReportPart], cfg: Config) -> Report:
    """Combina las filas de uno o varios archivos, fija el horizonte y calcula los vectores mensuales."""
    issues = [i for part in parts for i in tagged_issues(part, part.issues)]
    if len(parts) > 1:
        # Las hojas identifican cada fila por su código EDT: un código no puede estar en dos archivos
        owner: dict[str, str] = {}
        for part in parts:
            for row in part.rows:
                other = owner.setdefault(row.code, part.source.name)
                if other != part.source.name:
                    raise ConfigError(
                        "EDT-011",
                        f"El código EDT '{row.code}' está en los archivos '{other}' y '{part.source.name}'. Con "
                        "varios archivos de MS Project, cada código EDT debe estar en un solo archivo (p. ej. 1.0.0 a "
                        "8.0.0 en uno y 9.0.0 en otro): se deben revisar los códigos EDT o la planificación de cada "
                        "archivo en [[proyectos]].")
        roots = sorted((r for part in parts for r in part.roots), key=_code_order)
        rows = [x for root in roots for x in _block(root)]
    else:
        roots, rows = list(parts[0].roots), list(parts[0].rows)

    # --- Horizonte -------------------------------------------------------------
    # Sin mes de inicio: el primer mes con % planificado en las filas del libro (las tareas omitidas no cuentan)
    if cfg.start_month:
        start = first_of_month(cfg.start_month)
    else:
        keys = [k for row in rows if row.is_input for c in row.pct.values() for k, v in c.items() if v]
        if keys:
            y, m = min(keys)
            start = date(y, m, 1)
        else:
            firsts = [p.first_date for p in parts if p.first_date]
            start = first_of_month(min(firsts) if firsts else datetime.now().date())
    horizon = Horizon(start, cfg.months)

    # --- Vectores mensuales ---------------------------------------------------
    pct_before: set[str] = set()       # filas con % planificado fuera del horizonte
    pct_after: set[str] = set()
    for row in rows:
        if not row.is_input:
            continue
        for role, values in row.pct.items():
            vec, before, after = horizon.to_vector(values)
            vec[0] += before                # se acumula en el mes 1 para que las curvas acumuladas sean correctas
            row.pct_vec[role] = vec
            if before > 1e-9:
                pct_before.add(row.code)
            if after > 1e-9:
                pct_after.add(row.code)
        row.vector = [sum(v[i] for v in row.pct_vec.values()) for i in range(horizon.months)]

    def fill_summary(row: ReportRow) -> None:
        if row.is_input:
            return
        for c in row.children:
            fill_summary(c)
        row.vector = [sum(c.vector[i] for c in row.children) for i in range(horizon.months)]

    for r in roots:
        fill_summary(r)
    total_vector = [sum(r.vector[i] for r in roots) for i in range(horizon.months)]

    if pct_before:
        issues.append(Issue("CFG-044", "", "",
                            f"{plural(len(pct_before), 'fila tiene', 'filas tienen')} % planificado anterior al mes "
                            f"1 ({horizon.start:%m/%Y}): se acumuló en la columna del mes 1."))
    if pct_after:
        issues.append(Issue("CFG-045", "", "",
                            f"{plural(len(pct_after), 'fila tiene', 'filas tienen')} % planificado posterior al "
                            f"último mes del horizonte ({horizon.end:%m/%Y}): no aparece en las columnas mensuales "
                            "y su PV no llega al BAC. Se debe aumentar 'meses'."))

    # --- Controles -------------------------------------------------------------
    for part in parts:
        issues += tagged_issues(part, part.checks)
    issues += list_issues(parts, cfg)
    sheets = (f"PV-{cfg.company_a} y PV-{cfg.company_b}" if cfg.multi_enabled else "PV")
    issues.insert(0, Issue("LIB-001", "", "",
                           "El PV y el EV de cada fila son el BAC × el % planificado y el % ganado (hojas % PV y "
                           "% EV), que se calculan con la duración de la línea base de las tareas y el peso de los "
                           "hitos. Se debe ingresar el BAC de cada fila en las celdas amarillas de la columna D de "
                           f"{'las hojas' if cfg.multi_enabled else 'la hoja'} {sheets}."))
    undefined = [r for r in rows if r.is_input and is_undefined(r.method)]
    if undefined:
        issues.append(Issue("MET-005", "", "",
                            f"{plural(len(undefined), 'fila', 'filas')} sin método de medición definido "
                            "(columna C)."))

    sources = [p.source for p in parts]
    symbols = {s.name: s.project.currency_symbol for s in sources if s.project.currency_symbol}
    currency = next(iter(symbols.values()), "")
    if not cfg.currency and len(set(symbols.values())) > 1:
        issues.append(Issue("CFG-046", "", "",
                            "Los archivos tienen monedas distintas en MS Project ("
                            + ", ".join(f"{k}: {v}" for k, v in symbols.items()) + f"): el libro muestra {currency}. "
                            "Se debe indicar la moneda en [exportacion] moneda."))
    if len(sources) > 1 or sources[0].name:
        pv_source = "  |  ".join(f"{s.name}: {s.pv_source}" for s in sources)
        title = cfg.title or " + ".join(dict.fromkeys(s.project.name for s in sources))
        plan = ""
    else:
        pv_source, plan = sources[0].pv_source, sources[0].plan
        title = cfg.title or sources[0].project.name

    report = Report(project=sources[0].project, horizon=horizon, rows=rows, roots=roots, issues=issues,
                    total_vector=total_vector, pv_source=pv_source, plan=plan,
                    multi=cfg.multi_enabled, company_a=cfg.company_a if cfg.multi_enabled else "",
                    company_b=cfg.company_b if cfg.multi_enabled else "", sources=sources, title=title,
                    currency=currency, cost_review=cfg.cost_review_role)
    if report.cost_review is not None:
        from evm.cost_review import review_issues
        report.issues[1:1] = review_issues(report)      # después del aviso del cálculo del PV y el EV
    return report
