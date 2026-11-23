"""Construcción de la estructura EDT del informe (hasta el nivel máximo o lo que indique el campo Visible) y
% planificado mensual de cada fila."""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import date, datetime

from evm import progress
from evm.config import Config, ConfigError
from evm.fields import fmt_number
from evm.messages import (SEV_ERROR, SEV_INFO, SEV_WARN, TARGET_DETAIL, TARGET_GENERAL,  # noqa: F401
                          TARGET_OUT, TARGET_ROW, TOPIC_BOOK, TOPIC_COMPANY, TOPIC_DATES, TOPIC_LOGIC, TOPIC_METHOD,
                          TOPIC_PROGRESS, TOPIC_READ, TOPIC_VISIBLE, TOPIC_WBS, TOPIC_WEIGHTS, TOPICS, Issue)
from evm.methods import MILESTONES, WEIGHTED, is_undefined, parse_method
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
    # Nivel del código EDT (code_level), que decide el color de la fila: al menos uno más que el de la fila
    # superior (los hijos de una planificación como INTERNO, con códigos desde 1.0.0, son de nivel 2); 0 en las -D
    wbs_level: int = 0
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
    weights: dict[str, object] = field(default_factory=dict, repr=False)    # percent.RowWeights (% ganado) por empresa
    # Tareas que forman el % planificado (pv_selection): las de detalle con línea base o la tarea desglosada después;
    # None = las de detalle (leaves; --pv-fechas-actuales); [] = la fila no tiene % planificado (ninguna tarea con
    # línea base). pv_out: tareas de detalle sin línea base que quedan fuera del PV
    pv_leaves: list[Task] | None = field(default=None, repr=False)
    pv_out: list[Task] = field(default_factory=list, repr=False)
    # Reparto de la fila entre sus tareas (percent.RowTree)
    tree: object = field(default=None, repr=False)
    # % ganado de cada mes hasta la fecha de estado que carga 'generar' (history.earned_history), de cada empresa
    ev_hist: dict[str, list[float]] = field(default_factory=dict, repr=False)

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
    # Textos de los coeficientes de la tabla Recargos ([depuracion_precio.recargos]) de un libro con depuración
    cost_rates: tuple[str, ...] = ()

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
    scope: "Scope | None" = None                # tareas exportadas (sin las inactivas ni las de Peso = 0)
    code_of: Callable[[Task], str] | None = None    # código EDT de una tarea, como en el libro
    list_matched: set[str] | None = None        # lista EDT: códigos de la lista que tienen tarea (None = sin lista)
    # validar: evaluación DCMA de 14 puntos, punto -> (valor, umbral, cumple; None = no se evalúa)
    dcma: dict = field(default_factory=dict)
    aggregated: int = 0                         # tareas sin fila propia sumadas en la fila de su elemento superior
    # El campo de método configurado no se encontró o está vacío en todas las tareas (LEC-010 o LEC-019): sus filas
    # no se cuentan en MET-005
    method_unread: bool = False


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


def _active_branch(t: Task) -> bool:
    """La tarea y todas sus tareas superiores están activas."""
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


def _under(t: Task, ancestor: Task) -> bool:
    """La tarea depende de ancestor (en el árbol completo de MS Project)."""
    n = t.parent
    while n is not None:
        if n is ancestor:
            return True
        n = n.parent
    return False


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
            f"de nivel superior del proyecto: {', '.join(tops[:15]) or 'ninguna'}. Si el archivo tiene una sola "
            "planificación, se debe dejar vacía [exportacion] planificacion (planificacion = \"\") o quitar la clave.")
    plan = found[0]
    return Scope(list(plan.children), plan, (plan.wbs or "").strip() or plan_code)


@dataclass
class Visibility:
    """Filas del libro según el código EDT, el nivel máximo y el campo Visible."""

    rows: dict[int, bool]               # UniqueID -> tiene fila propia
    issues: list[Issue]
    shown: int = 0                      # tareas con Visible = Sí y fila propia por debajo del nivel máximo
    # Peso = 0: ramas que no forman parte del PV, el EV ni el AC (desde la 2.7.0, Visible = No ya no excluye)
    excluded: list[Task] = field(default_factory=list)
    listed: bool = False                # las filas las decide la lista EDT (lista_edt) o el campo Visible
    by_visible: bool = False            # las filas las decide el campo Visible (alguna tarea con Sí)
    matched: set[str] = field(default_factory=set)     # códigos de la lista EDT que tienen tarea
    # Tareas con fila propia y Peso = 0: la fila se conserva con PV y EV 0 (sus tareas sin fila no cuentan)
    zero_rows: list[Task] = field(default_factory=list)

    def has_row(self, task: Task) -> bool:
        return self.rows.get(task.uid, False)

    def excluded_uids(self) -> set[int]:
        """Tareas de las ramas excluidas (con sus subtareas)."""
        return {n.uid for t in self.excluded for n in iter_subtree(t)}


def row_visibility(scope: Scope, cfg: Config, code_of=None) -> Visibility:
    """Filas del libro (_rows_of) y ramas con Peso = 0, que no cuentan en el libro (_zero_weight_branches)."""
    out = _rows_of(scope, cfg, code_of)
    _zero_weight_branches(scope, out, code_of or scope.raw_code)
    return out


def _zero_weight_branches(scope: Scope, out: Visibility, code_of) -> None:
    """Peso = 0 escrito en el campo de peso (desde la 2.7.0): la tarea se hace, pero no cuenta en el libro EVM.

    - Una tarea sin fila propia con Peso = 0 queda fuera del PV, el EV y el AC con toda su rama (como una tarea
      inactiva): no recibe parte del peso de su tarea superior ni forma la fila -D (HIT-015). Si su rama tiene filas
      propias (Visible = Sí), esas filas se conservan y solo se quitan sus demás tareas.
    - Una tarea con fila propia y Peso = 0 conserva su fila (su código es una fila del libro), con PV y EV 0: sus
      tareas sin fila no cuentan; sus filas inferiores conservan su valor (HIT-016).
    - Si la rama quitada tiene subtareas con peso, esos pesos no se usan (HIT-014, advertencia).
    """
    from evm.percent import task_weight, zero_weight
    below: dict[int, bool] = {}

    def rows_below(t: Task) -> bool:
        found = False
        for c in t.children:
            found = rows_below(c) or out.has_row(c) or found
        below[t.uid] = found
        return found

    for r in scope.roots:
        rows_below(r)
    cut: dict[int, Task] = {}          # ramas que no cuentan: tarea superior de cada una
    emptied: list[Task] = []            # filas cuyas tareas tienen todas Peso = 0

    def walk(t: Task, under: bool) -> bool:
        """True si la tarea queda fuera con toda su rama."""
        if out.has_row(t):
            results = [walk(c, zero_weight(t)) for c in t.children]
            if zero_weight(t):
                out.zero_rows.append(t)
            elif results and all(results):
                out.zero_rows.append(t)
                emptied.append(t)
            return False
        zero = under or zero_weight(t)
        if zero and not below[t.uid]:
            cut[t.uid] = t
            return True
        results = [walk(c, zero) for c in t.children]
        if results and all(results) and not below[t.uid]:
            # Todas sus subtareas tienen Peso = 0: tampoco cuenta (sin subtareas, contaría con sus propias fechas)
            for c in t.children:
                cut.pop(c.uid, None)
            cut[t.uid] = t
            return True
        return False

    for r in scope.roots:
        walk(r, False)
    cut = list(cut.values())
    if cut:
        out.excluded.extend(cut)
        n = sum(1 for t in cut for _ in iter_subtree(t))
        names = ", ".join(code_of(t) or f"Id {t.id}" for t in cut[:10]) + (f" y {len(cut) - 10} más" if len(cut) > 10
                                                                          else "")
        out.issues.append(Issue("HIT-015", "", "",
                                f"{plural(n, 'tarea no cuenta', 'tareas no cuentan')} en el libro por tener Peso = 0 "
                                f"(ella o su tarea superior): {names}. Se hacen igualmente, pero quedan fuera del PV, "
                                "el EV y el AC, y no forman filas -D."))
        for t in cut:
            lost = [x for x in iter_subtree(t) if x is not t and task_weight(x) > 0]
            if lost:
                listed = ", ".join(code_of(x) or f"Id {x.id}" for x in lost[:5]) + (
                    f" y {len(lost) - 5} más" if len(lost) > 5 else "")
                out.issues.append(Issue("HIT-014", code_of(t), task_name(t, code_of),
                                        f"Tarea con Peso = 0 o dentro de una con Peso = 0: no cuenta en el libro, así "
                                        f"que no se usan los pesos de sus subtareas ({listed}).", severity=SEV_WARN,
                                        uid=t.uid))
    for t in out.zero_rows:
        rows = below.get(t.uid, False)
        why = ("Todas las tareas de la fila tienen Peso = 0" if t.uid in {x.uid for x in emptied} else "Fila con Peso = 0")
        out.issues.append(Issue("HIT-016", code_of(t), t.name,
                                f"{why}: se conserva (su código es una fila del libro), pero sus tareas sin fila propia "
                                "no cuentan, así que su PV y su EV son 0"
                                + ("; sus filas inferiores conservan su valor." if rows else "."), uid=t.uid))


def _rows_of(scope: Scope, cfg: Config, code_of=None) -> Visibility:
    """Decide qué tareas tienen fila propia en el libro (desde la 2.7.0).

    - Si alguna tarea tiene Visible = Sí, las filas las decide el campo, como una lista EDT en el propio archivo:
      tiene fila cada tarea con Sí y código EDT; las demás (vacío o No) no tienen fila y se suman en su fila superior
      con Sí (VIS-011). El nivel máximo solo completa los códigos con ceros.
    - Si ninguna tiene Sí: fila propia si la tarea tiene código EDT de nivel no mayor que el nivel máximo; una tarea
      con Visible = No (y su rama) no tiene fila y se suma en su fila superior.

    Visible = No ya no excluye tareas del cálculo (VIS-012): para eso se inactivan en MS Project.
    """
    if cfg.wbs_list is not None:
        return list_visibility(scope, cfg)
    from evm.fields import parse_visible
    code_of = code_of or scope.raw_code
    values: dict[int, bool | None] = {}
    issues: list[Issue] = []
    if cfg.visible_field:               # visible = "" en la configuración: el campo no se usa
        for t in scope.tasks():
            value, valid = parse_visible(t.visible)
            values[t.uid] = value
            if not valid:
                issues.append(Issue("VIS-001", code_of(t), t.name,
                                    f"Visible = '{t.visible}' no se reconoce (se admite Sí o No): se trata como vacío.",
                                    uid=t.uid))
    tasks = list(scope.tasks())
    no = [t for t in tasks if values.get(t.uid) is False]
    if no:
        names = ", ".join(code_of(t) or t.name for t in no[:10]) + (f" y {len(no) - 10} más" if len(no) > 10 else "")
        issues.append(Issue("VIS-012", "", "",
                            f"{plural(len(no), 'tarea tiene', 'tareas tienen')} Visible = No ({names}): no "
                            f"{'tiene' if len(no) == 1 else 'tienen'} fila propia y "
                            f"{'cuenta' if len(no) == 1 else 'cuentan'} en su fila superior (antes de la versión 2.7.0 "
                            "quedaban fuera del cálculo). Para excluir una tarea del cálculo, se debe inactivar en MS "
                            "Project."))
    if any(values.get(t.uid) is True for t in tasks):
        out = Visibility({}, issues, listed=True, by_visible=True)
        for t in tasks:
            raw = scope.raw_code(t)
            out.rows[t.uid] = bool(raw) and values.get(t.uid) is True
            if values.get(t.uid) is True and not raw:
                out.issues.append(Issue("VIS-004", code_of(t), t.name,
                                        "Visible = Sí en una tarea sin código EDT: no puede tener fila propia; se "
                                        "suma en su fila superior.", uid=t.uid))
        shown = sum(out.rows.values())
        out.issues.append(Issue("VIS-011", "", "",
                                f"Filas del libro según el campo Visible: {plural(shown, 'tarea', 'tareas')} con Sí "
                                f"{'tiene' if shown == 1 else 'tienen'} fila propia; las demás (vacío o No) se suman "
                                "en su fila superior."))
        return out
    out = Visibility({}, issues)
    kept: list[Task] = []               # Visible = No sin fila superior en la que sumarse: conservan su fila

    def walk(group: list[Task], under_no: bool, row_above: bool) -> None:
        for t in group:
            raw = scope.raw_code(t)
            eligible = bool(raw) and code_level(raw, cfg.ignore_interior_zeros) <= cfg.max_level
            inner = under_no or values.get(t.uid) is False
            if inner and eligible and not row_above:
                kept.append(t)
                inner = False
            out.rows[t.uid] = eligible and not inner
            walk(t.children, inner, row_above or out.rows[t.uid])

    walk(scope.roots, False, False)
    if kept:
        names = ", ".join(code_of(t) or t.name for t in kept[:10]) + (f" y {len(kept) - 10} más" if len(kept) > 10
                                                                      else "")
        out.issues.append(Issue("VIS-012", "", "",
                                f"{plural(len(kept), 'tarea', 'tareas')} con Visible = No sin fila superior en la que "
                                f"sumarse ({names}): {'conserva' if len(kept) == 1 else 'conservan'} su fila, para que "
                                "su rama cuente en el libro."))
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
# Sin la línea base seleccionada (FEC-009): el % planificado sale de las fechas actuales
CURRENT_PV_SOURCE = ("% planificado de MS Project con las fechas ACTUALES de las tareas (sin línea base) y el peso de "
                     "los hitos × BAC ingresado en el libro")


def leaf_span(leaves: list[Task], start: datetime | None = None,
              finish: datetime | None = None) -> tuple[datetime | None, datetime | None]:
    """Comienzo y fin de una fila (columnas Fecha inicio y Fecha fin Plan de la hoja PV): el comienzo más temprano y
    el fin más tardío de las tareas de detalle que forman su % planificado, con las mismas fechas (las de la línea
    base o, si no las tienen, las actuales; un fin anterior al comienzo se invierte, como en el reparto). Así las
    fechas coinciden con los meses del % PV aunque la línea base de la tarea resumen no esté resumida o la rama tenga
    tareas con Visible = No. start y finish: las fechas si ninguna tarea de detalle tiene fechas."""
    from evm.percent import task_dates
    starts: list[datetime] = []
    finishes: list[datetime] = []
    for t in leaves:
        a, b, _ = task_dates(t)
        a, b = a or b, b or a
        if a is None or b is None:
            continue
        starts.append(min(a, b))
        finishes.append(max(a, b))
    if not starts:
        return start, finish
    return min(starts), max(finishes)


def has_baseline_dates(t: Task) -> bool:
    """La tarea tiene fechas de línea base (sin línea base en el proyecto, use_current_schedule les da las actuales)."""
    return t.baseline_start is not None or t.baseline_finish is not None


def pv_selection(leaves: list[Task], top: Task, include_top: bool, current_dates: bool = False,
                 allowed: set[int] | None = None) -> tuple[list[Task] | None, list[Task]]:
    """Tareas que forman el % planificado de una fila y tareas de detalle que quedan fuera (desde la 2.7.0).

    El PV es el trabajo planificado: las tareas de detalle con línea base y, si una tarea con línea base se desglosó
    después en subtareas sin línea base, esa tarea, como un bloque con las fechas de su línea base. Una tarea de
    detalle sin línea base que no depende de ninguna de ellas (p. ej. un retrabajo o trabajo no previsto, agregado
    después de guardar la línea base) queda fuera del PV: cuenta en el % ganado y en la fecha de fin. Si ninguna tarea
    de la fila tiene línea base, la fila no tiene % planificado: ([], tareas de detalle). Con current_dates (generar
    --pv-fechas-actuales, el criterio anterior a la 2.7.0) el PV se forma con las tareas de detalle, con sus fechas
    actuales si no tienen línea base: (None, []). Un proyecto sin línea base usa en todas las tareas sus fechas
    actuales como línea base (use_current_schedule).

    top: tarea de la fila (en una fila -D, la tarea resumen, que no forma parte de ella: include_top = False);
    allowed: UniqueID de las tareas que forman la fila (None = todo el árbol podado).
    """
    if current_dates:
        return None, []
    from evm.percent import no_count
    memo: dict[int, bool] = {}

    def baselined_below(t: Task) -> bool:
        if t.uid not in memo:
            # Todas las subtareas que cuentan, también las que tienen fila propia (en una fila -D): igual en 'generar'
            # (árbol podado) y en 'actualizar' (allowed solo tiene las tareas de la fila)
            memo[t.uid] = any(not no_count(c) and (has_baseline_dates(c) or baselined_below(c))
                              for c in t.children)
        return memo[t.uid]

    units: list[Task] = []
    seen: set[int] = set()
    out: list[Task] = []
    for leaf in leaves:
        unit = leaf if has_baseline_dates(leaf) else None
        p = leaf.parent if unit is None and leaf is not top else None
        while unit is None and p is not None and (include_top or p is not top) \
                and (allowed is None or p.uid in allowed):
            if has_baseline_dates(p):
                unit = p if not baselined_below(p) else None
                break
            if p is top:
                break
            p = p.parent
        if unit is None:
            out.append(leaf)
        elif unit.uid not in seen:
            seen.add(unit.uid)
            units.append(unit)
    return units, out


def _weight_issues(row: ReportRow, tree, code_of) -> list[Issue]:
    """Observaciones del reparto de una fila entre sus tareas (percent.RowTree): pesos que no suman 100, tareas sin
    parte y pesos que no se usan. Los hitos sin peso solo se informan en las filas «Ponderado» (HIT-010): en las
    demás es normal que un hito no tenga valor propio."""
    from evm import percent as pct
    issues: list[Issue] = []
    kind = parse_method(row.method).kind
    weighted, milestones = kind == WEIGHTED, kind == MILESTONES
    where = f"la fila {row.code}"
    if weighted and not tree.has_weights and tree.nodes:      # una fila de detalle no tiene tareas que pesar
        issues.append(Issue("HIT-012", row.code, row.name, "Método «Ponderado» y ninguna tarea con peso: % planificado "
                                                           "por duración y % ganado con el método de cada tarea."))
    if tree.negative:
        n = len(tree.negative)
        issues.append(Issue("HIT-006", row.code, row.name,
                            f"{plural(n, 'tarea', 'tareas')} con peso negativo ("
                            + ", ".join(code_of(x) or task_name(x, code_of) for x in tree.negative[:5])
                            + (f" y {n - 5} más" if n > 5 else "") + "): se toma como si no tuviera peso."))
    for t, why in tree.unweighted:
        reason = ("sus tareas hermanas con peso suman 100 o más" if why == pct.NO_SHARE_FULL else
                  "el resto de su tarea superior se reparte por duración y la suya es 0")
        lost = [x for x in iter_subtree(t) if x is not t and pct.task_weight(x) > 0]
        if lost:
            names = ", ".join(code_of(x) or task_name(x, code_of) for x in lost[:5]) + (
                f" y {len(lost) - 5} más" if len(lost) > 5 else "")
            issues.append(Issue("HIT-014", code_of(t), task_name(t, code_of),
                                f"Tarea resumen sin peso en {where}: {reason}, y no forma parte del % planificado "
                                "ni del % ganado. "
                                + ("El peso de su subtarea" if len(lost) == 1 else f"Los pesos de sus {len(lost)} subtareas")
                                + f" ({names}) no se usan: se debe indicar el peso de la tarea resumen.",
                                uid=t.uid))
            continue
        if pct.is_zero_duration(t):
            if weighted and tree.has_weights:       # sin ningún peso ya lo informa HIT-012
                issues.append(Issue("HIT-010", code_of(t), task_name(t, code_of),
                                    f"Hito sin peso en la fila «Ponderado» {row.code}: no forma parte del % planificado "
                                    "ni del % ganado.", uid=t.uid))
            continue
        if not weighted and why == pct.NO_SHARE_DURATION:
            continue                        # sin duración (p. ej. una tarea resumen de hitos sin peso)
        if milestones:
            continue                        # «Hitos ponderados»: lo esperado es que las tareas no cuenten
        subtasks = len(tree.branch_leaves(t)) if t.children else 0
        issues.append(Issue("HIT-011", code_of(t), task_name(t, code_of),
                            f"Sin peso en {where}"
                            + (f" (con sus {subtasks} subtareas)" if subtasks > 1 else "")
                            + f": {reason}, y no forma parte del % planificado ni del % ganado.", uid=t.uid))
    for g in tree.groups:
        own = g.task is None or g.task is row.task
        code, name = (row.code, row.name) if own else (code_of(g.task), task_name(g.task, code_of))
        place = "la fila" if own else f"«{g.task.name}»"
        total = fmt_number(g.total, 2, trim=True)
        if g.total > pct.WEIGHT_TOTAL:
            issues.append(Issue("HIT-013", code, name,
                                f"Los pesos de las subtareas de {place} suman {total}, más de 100: se escalan a 100 y "
                                "sus subtareas sin peso no cuentan. Se deben corregir los pesos para que sumen como "
                                "máximo 100.", SEV_ERROR, uid=None if own else g.task.uid))
        else:
            issues.append(Issue("HIT-013", code, name,
                                f"Los pesos de las subtareas de {place} suman {total} y no hay subtareas sin peso con "
                                "duración que reciban el resto: se escalan a 100.", SEV_WARN,
                                uid=None if own else g.task.uid))
    for t, zero in tree.equal_zero:
        place = "la fila" if t is None or t is row.task else f"«{t.name}»"
        names = ", ".join(code_of(x) or task_name(x, code_of) for x in zero[:5])
        issues.append(Issue("HIT-009", row.code, row.name,
                            f"Ninguna subtarea de {place} tiene peso y {names} no "
                            f"{'tiene' if len(zero) == 1 else 'tienen'} duración, pero "
                            f"{'contiene' if len(zero) == 1 else 'contienen'} tareas con peso: todas las subtareas "
                            f"de {place} pesan igual. Si no deben pesar igual, se indica su peso."))
    if tree.equal:
        places = ", ".join("la fila" if t is None or t is row.task else f"«{t.name}»" for t in tree.equal[:5])
        issues.append(Issue("HIT-009", row.code, row.name,
                            f"Las tareas de {places} no tienen duración ni peso: pesan igual."))
    for t in tree.pv_lost:
        issues.append(Issue("HIT-017", code_of(t), task_name(t, code_of),
                            f"Sin peso en {where} y sus tareas con línea base no tienen duración: no recibe parte del % "
                            "planificado, aunque sí del % ganado (sus tareas sin línea base tienen duración). Se debe "
                            "indicar su peso.", uid=t.uid))
    return issues


def pv_span(pv_leaves: list[Task] | None, leaves: list[Task], start: datetime | None,
            finish: datetime | None) -> tuple[datetime | None, datetime | None]:
    """Fechas del plan de una fila (columnas E y F de la hoja PV): las de sus tareas del % planificado; sin % planificado
    ([]), vacías."""
    if pv_leaves is not None and not pv_leaves:
        return None, None
    return leaf_span(leaves if pv_leaves is None else pv_leaves, start, finish)


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
    planned: dict[int, pct.TaskTime] = {}       # tareas que forman algún % planificado
    excluded: list[tuple[Task, ReportRow]] = []  # tareas de detalle sin línea base fuera del PV
    no_pv: list[ReportRow] = []                 # filas sin % planificado (ninguna tarea con línea base)
    alone: dict[str, list[str]] = {}            # empresa -> filas en las que no participa
    if cfg.multi_enabled:
        roles = [(pct.ROLE_A, cfg.company_a, share), (pct.ROLE_B, cfg.company_b, lambda t: 1.0 - share(t))]
    else:
        roles = [(pct.ROLE_ALL, "", None)]

    def time_of(t: Task) -> pct.TaskTime:
        if t.uid not in cache:
            cache[t.uid] = pct.task_time(t, calendar, cfg.distribution)
        return cache[t.uid]

    for i, row in enumerate(rows, start=1):
        progress.step(i, len(rows), "filas")
        if not row.is_input:
            continue
        own = row.pv_leaves is not None                       # tareas del PV distintas de las de detalle
        tree = pct.RowTree(row.task, row.leaves, time_of, row.pv_leaves, parse_method(row.method),
                           residual=row.kind == KIND_RESIDUAL)
        row.tree = tree
        planned.update((t.uid, time_of(t)) for t in (row.leaves if not own else row.pv_leaves))
        excluded.extend((t, row) for t in row.pv_out)
        if own and not row.pv_leaves:
            no_pv.append(row)
        for role, company, fn in roles:
            rw, pv_pairs = tree.company(fn)
            if rw.fallback:
                alone.setdefault(company, []).append(row.code)
            row.pct[role] = pct.weighted_curve(pv_pairs)
            row.weights[role] = rw
        issues += _weight_issues(row, tree, code_of)
        if not any(row.pct.values()) and (not own or row.pv_leaves):
            issues.append(Issue("HIT-005", row.code, row.name,
                                "Ninguna tarea de la fila tiene fechas ni es un hito con peso: no tiene % "
                                "planificado y su BAC no se puede repartir en el tiempo."))

    leaves = list(planned.values())
    undated = [x for x in leaves if not x.dist]
    current = [x for x in leaves if x.dist and not x.from_baseline] if project.baseline_found else []
    if excluded and project.baseline_found:
        rows_out = sorted({r.code for _t, r in excluded})
        shown = ", ".join(rows_out[:10]) + (f" y {len(rows_out) - 10} más" if len(rows_out) > 10 else "")
        issues.append(Issue("FEC-019", "", "",
                            f"{plural(len(excluded), 'tarea sin línea base no forma', 'tareas sin línea base no forman')} "
                            f"parte del % planificado de su fila ({shown}): cuentan en el % ganado y en la fecha de "
                            "fin, pero no cambian el PV aprobado. Es lo esperado en un retrabajo o en trabajo no "
                            "previsto; si son trabajo aprobado, se debe guardar su línea base (Tareas seleccionadas y "
                            "Resumir líneas base). Con generar --pv-fechas-actuales entran en el PV con sus fechas "
                            "actuales." + (f" {plural(len(no_pv), 'fila no tiene', 'filas no tienen')} ninguna "
                                           f"tarea con línea base ({', '.join(r.code for r in no_pv[:10])}"
                                           f"{f' y {len(no_pv) - 10} más' if len(no_pv) > 10 else ''}): sin % "
                                           "planificado; su BAC debe ser 0." if no_pv else "")))
        for t, r in excluded[:200]:
            issues.append(Issue("FEC-019", code_of(t), task_name(t, code_of),
                                f"Sin línea base: fuera del % planificado de la fila {r.code}; cuenta en el % ganado.",
                                uid=t.uid))
    if current:
        issues.append(Issue("FEC-010", "", "",
                            f"{plural(len(current), 'tarea no tiene', 'tareas no tienen')} fechas de línea base "
                            f"(p. ej. se {'agregó' if len(current) == 1 else 'agregaron'} después de guardarla): "
                            "su % planificado se calculó con sus fechas actuales. Si corresponde, se debe "
                            "actualizar la línea base de esas tareas."))
        for x in current[:200]:
            issues.append(Issue("FEC-010", code_of(x.task), task_name(x.task, code_of),
                                "Sin fechas de línea base: se usaron sus fechas actuales.", uid=x.task.uid))
    for x in undated[:200]:
        issues.append(Issue("FEC-012", code_of(x.task), task_name(x.task, code_of),
                            "Sin fechas de línea base ni fechas actuales: no aporta % planificado.", uid=x.task.uid))
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


def _without_baseline(project: ProjectData, name: str) -> Issue:
    """FEC-009: el proyecto no tiene la línea base que usa el % planificado. Si tiene otras guardadas (con XML; con
    un .mpp, solo en validar), se indican, con la clave para usarlas."""
    n = project.baseline_number
    used = "la Línea base" if n == 0 else f"la Línea base {n}"
    counts: dict[int, int] = {}
    for t in project.tasks:
        for k in t.baselines:
            if k != n:
                counts[k] = counts.get(k, 0) + 1
    other = ""
    if counts:
        best = max(counts, key=lambda k: (counts[k], -k))
        items = [f"{'Línea base' if k == 0 else f'Línea base {k}'} ({fmt_number(c, 0)} "
                 f"{'tarea' if c == 1 else 'tareas'})" for k, c in sorted(counts.items())]
        saved = items[0] if len(items) == 1 else ", ".join(items[:-1]) + " y " + items[-1]
        key = (f"linea_base = {best} en su sección [[proyectos]]" if name else
               f"[exportacion] linea_base = {best} o --linea-base {best}")
        if len(counts) == 1:
            other = f" Tiene guardada la {saved}: si es la aprobada, se puede usar con {key}, después de revisar que "
            other += "esté completa."
        else:
            other = (f" Tiene guardadas {saved}: si una de ellas es la aprobada, se puede usar con linea_base (p. ej. "
                     f"{key}), después de revisar que esté completa.")
    return Issue("FEC-009", "", "",
                 f"El proyecto no tiene guardada {used}, que usa el % planificado: el % planificado se calculó con las "
                 "fechas ACTUALES de las tareas, y el PV sigue al programa actual. Se debe guardar la línea base en "
                 f"MS Project (Proyecto > Establecer línea base) y volver a exportar.{other}")


def build_part(project: ProjectData, cfg: Config, name: str = "") -> ReportPart:
    """Filas EDT y % planificado (por mes calendario) de un archivo de MS Project.

    name: nombre del archivo en [[proyectos]] ("" = informe de un solo archivo).
    """
    from evm.validation import data_checks, wbs_checks
    issues: list[Issue] = list(project.warnings)
    tasks = [t for t in project.tasks if t.outline_level >= 1]
    roots = build_tree(tasks)

    # Tareas inactivas (MS Project 2010+): se excluyen con toda su rama
    inactive: list[Task] = []           # tareas inactivas cuya tarea superior está activa (con su rama)
    if not cfg.include_inactive:
        inactive = [t for t in tasks if not t.active and (t.parent is None or _active_branch(t.parent))]
        roots, _, _ = prune_inactive(roots)

    # Planificación (p. ej. INTERNO o CLIENTE): solo las tareas que dependen de ella. Las tareas inactivas se cuentan
    # solo las de la planificación
    scope = select_scope(roots, cfg.plan)
    removed = sum(sum(1 for _ in iter_subtree(t)) for t in inactive if scope.plan is None or _under(t, scope.plan))
    if removed:
        issues.append(Issue("EDT-016", "", "",
                            f"Se {'excluyó' if removed == 1 else 'excluyeron'} "
                            f"{plural(removed, 'tarea inactiva', 'tareas inactivas')}"
                            f"{'' if scope.plan is None else f' de la planificación {scope.prefix}'}: no "
                            f"{'forma' if removed == 1 else 'forman'} parte del % planificado."))
    if scope.plan is not None:
        from evm.fields import same_name
        plan_name = (scope.plan.name or "").strip()
        named = f" («{plan_name}»)" if plan_name and not same_name(plan_name, scope.prefix) else ""
        by_name = "" if (scope.plan.wbs or "").strip() else " (tarea sin código EDT, elegida por su nombre)"
        issues.append(Issue("EDT-017", "", "",
                            f"Planificación '{scope.prefix}'{named}{by_name}: se exportan las "
                            f"{sum(1 for _ in scope.tasks())} tareas que dependen de ella."))

    if not project.baseline_found:
        issues.append(_without_baseline(project, name))

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
    zero_rows = {t.uid for t in visibility.zero_rows}
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
        """Fechas de la línea base de la tarea (o de su rama): solo si sus tareas de detalle no tienen fechas."""
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
    beyond = ("sin Visible = Sí" if visibility.by_visible else "que no están en la lista EDT" if visibility.listed
              else f"de nivel mayor a {cfg.max_level}")

    def visit(t: Task, parent: ReportRow | None) -> None:
        nonlocal aggregated
        s, f = subtree_dates(t)
        row_tasks, loose_full, loose_own = gather(t)
        # En el nivel máximo la fila suma toda su rama, salvo que una tarea de la rama tenga Visible = Sí
        is_leaf_row = not row_tasks or (not visibility.listed
                                        and code_level(scope.raw_code(t), cfg.ignore_interior_zeros) >= cfg.max_level
                                        and not any(code_level(scope.raw_code(c), cfg.ignore_interior_zeros)
                                                    > cfg.max_level for c in row_tasks))
        wbs_level = max(code_level(scope.raw_code(t), cfg.ignore_interior_zeros), parent.wbs_level + 1 if parent else 1)
        row = ReportRow(code=code_of(t), name=t.name, level=(parent.level + 1) if parent else 1,
                        kind=KIND_LEAF if is_leaf_row else KIND_SUMMARY, task=t,
                        start=s, finish=f, wbs_level=wbs_level, parent=parent)
        rows.append(row)
        if parent:
            parent.children.append(row)
        else:
            report_roots.append(row)

        if is_leaf_row:
            # Fila con Peso = 0 (HIT-016): sus subtareas se quitaron del árbol y la fila no tiene tareas que cuenten
            row.leaves = [] if t.uid in zero_rows else [n for n in iter_subtree(t) if not n.children]
            row.pv_leaves, row.pv_out = pv_selection(row.leaves, t, True, cfg.pv_current_dates)
            row.start, row.finish = pv_span(row.pv_leaves, row.leaves, s, f)
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
            # Fechas de sus propias tareas (no las de toda la tarea resumen)
            pv_leaves, pv_out = pv_selection(loose, t, False, cfg.pv_current_dates)
            start, finish = pv_span(pv_leaves, loose, s, f)
            res = ReportRow(code=f"{row.code}-D", name=f"Otras tareas de «{t.name}»",
                            level=row.level + 1, kind=KIND_RESIDUAL, task=t, method=cfg.default_method,
                            start=start, finish=finish, parent=row)
            res.leaves = loose
            res.pv_leaves, res.pv_out = pv_leaves, pv_out
            rows.append(res)
            row.children.append(res)
            issues.append(Issue("EDT-018", row.code, t.name,
                                f"Tareas sin código EDT o {beyond}: forman la fila {row.code}-D, con su propio BAC."))
        # Fila subtotal: desde el comienzo de la primera de sus filas hasta el fin de la última
        starts = [c.start for c in row.children if c.start]
        finishes = [c.finish for c in row.children if c.finish]
        if starts and finishes:
            row.start, row.finish = min(starts), max(finishes)
        elif row.children and not cfg.pv_current_dates:
            row.start = row.finish = None       # ninguna fila con % planificado (sin línea base)

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
        # Sin el campo de empresa, o con el campo vacío en todas las tareas, ya lo informa LEC-010 o LEC-019
        unread = ("campo_empresa" not in project.fields_found
                  or not any(t.company for t in scope.tasks()))
        missing = [] if unread else notes["sin_empresa"]
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
                issues.append(Issue("EMP-003", code_of(t), task_name(t, code_of),
                                    f"Sin empresa: se trata como '{cfg.company_both}'"
                                    + (" (incluye sus subtareas sin empresa)." if t.children else "."), uid=t.uid))
        for t in notes["desconocida"][:50]:
            issues.append(Issue("EMP-005", code_of(t), task_name(t, code_of),
                                f"Empresa '{t.company}' no reconocida (se esperaba '{cfg.company_a}', "
                                f"'{cfg.company_b}' o '{cfg.company_both}'): se trata como '{cfg.company_both}'.",
                                uid=t.uid))
        for t in notes["reparto"][:50]:
            issues.append(Issue("EMP-006", code_of(t), task_name(t, code_of),
                                f"'{cfg.split_field}' = {fmt_number(t.split_a, trim=True)} fuera de rango (debe ser "
                                "mayor que 0 y como máximo 100): se usa "
                                f"{fmt_number(cfg.split_default, trim=True)} %.", uid=t.uid))

    # --- Controles del archivo ------------------------------------------------
    checks: list[Issue] = []
    # Es una advertencia si las tareas omitidas aportarían % planificado
    omitted_sev = SEV_WARN if any(_has_weight(n) for n in omitted if not n.children) else SEV_INFO
    if omitted and visibility.listed:
        one = len(omitted) == 1
        out_ids = {n.uid for n in omitted}
        coded = [n for n in omitted if scope.raw_code(n) and (n.parent is None or n.parent.uid not in out_ids)]
        shown = ", ".join(code_of(n) for n in coded[:10]) + (f" y {len(coded) - 10} más" if len(coded) > 10 else "")
        where = ("no tienen Visible = Sí ni dependen de una tarea con Sí" if visibility.by_visible else
                 f"no {'está' if one else 'están'} en la lista EDT ni {'depende' if one else 'dependen'} de un elemento "
                 "de la lista")
        if visibility.by_visible and one:
            where = "no tiene Visible = Sí ni depende de una tarea con Sí"
        checks.append(Issue("EDT-019", "", "",
                            f"Se {'omitió' if one else 'omitieron'} {plural(len(omitted), 'tarea', 'tareas')} que "
                            f"{where}: no {'forma' if one else 'forman'} parte de ninguna fila del libro ni del PV y el "
                            "EV" + (f" ({shown})." if shown else "."), omitted_sev))
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
        beyond_edt = ("sin Visible = Sí" if visibility.by_visible else "que no están en la lista EDT"
                      if visibility.listed else f"de nivel EDT mayor a {cfg.max_level}")
        checks.append(Issue("EDT-022", "", "",
                            f"{plural(aggregated, 'tarea', 'tareas')} sin código EDT o {beyond_edt} se "
                            f"{'agregó' if aggregated == 1 else 'agregaron'} en su elemento EDT superior."))
    for row in rows:
        row.source = name
    dates = [t.baseline_start or t.start for t in scope.tasks() if (t.baseline_start or t.start)]
    ref = min(dates) if dates else project.project_start
    first_date = (ref.date() if isinstance(ref, datetime) else ref) if ref else None
    pv_source = PERCENT_PV_SOURCE if project.baseline_found else CURRENT_PV_SOURCE
    method_unread = bool(cfg.method_field) and ("metodo" not in project.fields_found
                                                or not any(t.method for t in scope.tasks()))
    return ReportPart(Source(name, project, scope.prefix, pv_source), rows, report_roots, issues, checks,
                      first_date, scope, code_of, visibility.matched if visibility.listed else None,
                      aggregated=aggregated, method_unread=method_unread)


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
    out += [Issue("EDT-013", code, "", "Código de la lista EDT sin tarea en MS Project: no tiene fila en el libro.")
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


def locate_issues(part: ReportPart, issues: list[Issue]) -> list[Issue]:
    """Las observaciones de un archivo con su alcance y la fila del libro de su tarea: Fila EDT (la tarea es una fila
    de las hojas PV, EV y AC), Detalle (se suma en una fila: de nivel mayor al máximo, fuera de la lista EDT o sin
    código, también en la fila -D), Fuera del libro (no forma parte de ninguna fila: omitida, inactiva, con Visible =
    No o de otra planificación) y General (sin tarea). La tarea se busca por su UniqueID o, si la observación no lo
    indica, por su código EDT."""
    where: dict[int, tuple[str, str]] = {}          # UniqueID -> (alcance, fila del libro)
    for row in part.rows:
        if row.kind == KIND_RESIDUAL:
            for n in row.leaves:
                where.setdefault(n.uid, (TARGET_DETAIL, row.code))
            continue
        where[row.task.uid] = (TARGET_ROW, row.code)
        if row.kind == KIND_LEAF:
            for n in iter_subtree(row.task):
                where.setdefault(n.uid, (TARGET_DETAIL, row.code))
    exported = list(part.scope.tasks()) if part.scope is not None else []
    in_scope = {t.uid for t in exported}
    tasks = {t.uid: t for t in part.source.project.tasks}
    code_of = part.code_of or (lambda t: t.wbs or "")
    by_code: dict[str, Task] = {}
    for t in exported:
        if code_of(t):
            by_code.setdefault(code_key(code_of(t)), t)
    row_codes = {code_key(r.code): r.code for r in part.rows}

    def of_task(t: Task) -> tuple[str, str]:
        if t.uid in where:
            return where[t.uid]
        if t.uid in in_scope:           # sin fila propia, de una tarea resumen (p. ej. sin peso y sin fila -D)
            p = t.parent
            while p is not None:
                if where.get(p.uid, ("",))[0] == TARGET_ROW:
                    return TARGET_DETAIL, where[p.uid][1]
                p = p.parent
        return TARGET_OUT, ""

    out = []
    for i in issues:
        if i.target:
            out.append(i)
            continue
        key = code_key(i.code) if i.code else ""
        if i.uid is not None and i.uid in tasks:
            target, row = of_task(tasks[i.uid])
        elif key in row_codes:
            target, row = TARGET_ROW, row_codes[key]
        elif key in by_code:
            target, row = of_task(by_code[key])
        else:
            target, row = TARGET_GENERAL, ""
        out.append(replace(i, target=target, row=row))
    return out


def located(rows: list[ReportRow], issues: list[Issue]) -> list[Issue]:
    """Las observaciones con alcance. Las que no lo tienen (las del libro, la depuración de costes y --migrar-desde)
    son de una fila EDT si su código EDT es una fila del libro, y generales si no."""
    codes = {code_key(r.code): r.code for r in rows}
    out = []
    for i in issues:
        key = code_key(i.code) if i.code else ""
        if not i.target:
            i = replace(i, target=TARGET_ROW, row=codes[key]) if key in codes else replace(i, target=TARGET_GENERAL)
        out.append(i)
    return out


def tagged_issues(part: ReportPart, issues: list[Issue]) -> list[Issue]:
    """Las observaciones de un archivo con su alcance (locate_issues). Con varios archivos, cada observación indica
    el archivo: '[P2] ...'."""
    issues = locate_issues(part, issues)
    name = part.source.name
    if not name:
        return issues
    return [replace(i, message=f"[{name}] {i.message}", file=name) for i in issues]


def _after_horizon(rows: list[ReportRow], codes: set[str], horizon: Horizon) -> Issue:
    """CFG-045: filas con % planificado después del último mes del horizonte, con el % del BAC al que llega su PV
    (la empresa que menos llega, en multi-empresa), el último mes con % planificado y los meses necesarios."""
    reached: list[tuple[float, str]] = []
    last: MonthKey | None = None
    for row in rows:
        if not row.is_input:
            continue
        for values in row.pct.values():
            keys = [k for k, v in values.items() if v > 1e-9]
            if keys:
                last = max(last, max(keys)) if last else max(keys)
        if row.code in codes:
            share = min((sum(v) for v in row.pct_vec.values() if any(v)), default=0.0)
            reached.append((share, row.code))
    reached.sort()
    shown = ", ".join(f"{code} llega al {fmt_number(100 * share, 1, trim=True)} %" for share, code in reached[:10])
    more = f" y {len(reached) - 10} más" if len(reached) > 10 else ""
    needed = ""
    if last is not None:
        months = (last[0] - horizon.start.year) * 12 + last[1] - horizon.start.month + 1
        needed = (f" El último mes con % planificado es {last[1]:02d}/{last[0]}: se necesitan al menos {months} meses "
                  f"desde {horizon.start:%m/%Y} (hoy son {horizon.months}).")
    return Issue("CFG-045", "", "",
                 f"{plural(len(codes), 'fila tiene', 'filas tienen')} % planificado posterior al último mes del "
                 f"horizonte ({horizon.end:%m/%Y}): no aparece en las columnas mensuales y su PV no llega al BAC "
                 f"({shown}{more}).{needed} Se debe aumentar 'meses' (o --meses) y volver a generar.")


def _forecast_after_horizon(rows: list[ReportRow], horizon: Horizon) -> Issue | None:
    """CFG-052: el fin previsto de las tareas de las filas del libro (cronograma actual) es posterior al último mes
    del horizonte, en el que ya no se pueden cargar el EV ni el AC."""
    finishes = [t.finish for row in rows if row.is_input for t in row.leaves if t.finish is not None]
    if not finishes:
        return None
    last = max(finishes)
    if (last.year, last.month) <= (horizon.end.year, horizon.end.month):
        return None
    months = (last.year - horizon.start.year) * 12 + last.month - horizon.start.month + 1
    return Issue("CFG-052", "", "",
                 f"El fin previsto del cronograma ({last:%m/%Y}) es posterior al último mes del horizonte "
                 f"({horizon.end:%m/%Y}): el EV y el AC de los meses posteriores no se podrán cargar en el libro. Se "
                 f"necesitan al menos {months} meses desde {horizon.start:%m/%Y} (hoy son {horizon.months}): se debe "
                 "aumentar 'meses' (o --meses) y volver a generar.")


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
        issues.append(_after_horizon(rows, pct_after, horizon))
    late = _forecast_after_horizon(rows, horizon)
    if late is not None:
        issues.append(late)

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
    # Sin las filas de los archivos cuyo campo de método no se leyó: ya lo informa LEC-010 o LEC-019
    unread = {p.source.name for p in parts if p.method_unread}
    undefined = [r for r in rows if r.is_input and is_undefined(r.method) and r.source not in unread]
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
    status = {s.name: s.project.status_date.date() if s.project.status_date else None for s in sources}
    if len(sources) > 1 and len(set(status.values())) > 1:
        dates = ", ".join(f"{k}: {f'{d:%d/%m/%Y}' if d else 'sin fecha de estado'}" for k, d in status.items())
        issues.append(Issue("CFG-056", "", "",
                            f"Los archivos tienen fechas de estado distintas ({dates}): el avance de cada uno "
                            "corresponde a otro día. «actualizar» se detiene (ACT-005 o ACT-004) si no se indica "
                            "--fecha-corte. Se debe registrar el avance de todos los archivos a la misma fecha y fijar "
                            "esa fecha de estado en cada uno (Proyecto > Fecha de estado)."))
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
                    currency=currency, cost_review=cfg.cost_review_role,
                    cost_rates=tuple(cfg.cost_price_rates.values()) if cfg.cost_review_role is not None else ())
    if report.cost_review is not None:
        from evm.cost_review import review_issues
        report.issues[1:1] = review_issues(report)      # después del aviso del cálculo del PV y el EV
    return report
