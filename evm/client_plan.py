"""Planificación para el cliente (comando generar-cliente): qué tareas de la planificación interna recibe el cliente y
cómo queda la lógica del cronograma sin las tareas que no se muestran.

Reglas (docs/generar-cliente.md):

- Se muestra una tarea si el campo de [cliente] campo vale Sí (Si o S, en mayúsculas o minúsculas) en ella y en todas
  sus tareas superiores de la planificación. Con No, vacío u otro valor, la tarea no se muestra, con toda su rama.
- La tarea de la planificación (por ejemplo INTERNO) se conserva y se renombra ([cliente] nombre_planificacion). Las
  tareas de otras planificaciones se quitan, con sus vínculos.

Con [cliente] actividad_interna = true (por defecto), las tareas que no se muestran se conservan como «Otras tareas»:
sin nombre ni código EDT, con sus fechas, su duración, su avance, sus líneas base y sus vínculos. Así el % completado de
cada tarea, de los resúmenes y del proyecto, y el libro EVM, son los de la planificación interna. Solo se quitan las
inactivas, que no cuentan en ningún cálculo.

Con actividad_interna = false, las tareas que no se muestran se borran:

- Un resumen que se muestra sin ninguna subtarea mostrada pasa a ser una tarea (resumen convertido): conserva sus
  fechas, su duración, su avance y sus líneas base, y representa a sus subtareas en la lógica.
- Los vínculos entre tareas mostradas se conservan. Los de las subtareas de un resumen convertido pasan al resumen.
  Si entre dos tareas mostradas hay tareas ocultas, se crea un vínculo puente equivalente: su posición es la suma de
  las posiciones de los vínculos del camino más la duración de las tareas ocultas que el camino recorre de comienzo a
  fin (menos la de las que recorre de fin a comienzo).
- Un vínculo nuevo no debe retrasar a su sucesora respecto de la planificación interna: si lo haría, su posición se
  ajusta a la que dan las fechas actuales. Entre dos tareas solo puede haber un vínculo: de los posibles, se elige el
  que más restringe a la sucesora.
- No se crean vínculos entre tareas que ya están vinculadas, entre un resumen y sus subtareas ni los que cerrarían un
  ciclo.

Esta parte no usa MS Project: calcula los cambios y escribe el informe de control. Los cambios en el .mpp los hace
evm.client_writer con MS Project, en una copia del archivo.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from evm.config import Config, ConfigError
from evm.fields import builtin_kind, fmt_number, is_spanish_wbs_name, plain_text
from evm.messages import EvmError, Issue
from evm.model import LINK_FF, LINK_FS, LINK_SF, LINK_SS, Link, ProjectData, Task
from evm.percent import calendar_for
from evm.wbs import Scope, build_tree, iter_subtree, plural, select_scope
from evm.workcalendar import WorkCalendar

START, FINISH = "C", "F"
# Puntos que une cada tipo de vínculo: (punto de la predecesora, punto de la sucesora)
POINTS = {LINK_FF: (FINISH, FINISH), LINK_FS: (FINISH, START), LINK_SF: (START, FINISH), LINK_SS: (START, START)}
KIND_OF = {points: kind for kind, points in POINTS.items()}
LINK_NAMES = {LINK_FF: "FF", LINK_FS: "FC", LINK_SF: "CF", LINK_SS: "CC"}

# Formatos de la posición en tiempo laborable (<LagFormat> del XML y LagType de COM; +32 = estimada): minutos, horas,
# días, semanas y meses. Una posición transcurrida o en % no se puede sumar a una duración: el vínculo que la tiene en
# su camino toma la posición de las fechas actuales.
_WORKING_FORMATS = {3, 5, 7, 9, 11}

ORIGIN_SUMMARY = "resumen"      # vínculo de una subtarea que pasa a su resumen convertido
ORIGIN_BRIDGE = "puente"        # vínculo que reemplaza un camino que pasa por tareas ocultas

SKIP_LINKED = "ya están vinculadas"
SKIP_SUMMARY = "resumen y subtarea"
SKIP_CYCLE = "cerraría un ciclo"

# actividad_interna: nombre de las tareas que no se muestran en el archivo del cliente
INTERNAL_NAME = "Otras tareas"

MARK_YES, MARK_NO, MARK_EMPTY, MARK_OTHER = "Sí", "No", "vacío", "otro valor"
# Valores del campo del cliente (sin tildes ni mayúsculas): solo Sí, Si y S muestran la tarea. Un valor que no es Sí
# ni No (p. ej. un error de escritura) tampoco la muestra, pero se avisa (CLI-003)
_YES = {"si", "s"}
_NO = {"no", "n"}

# Pasos de búsqueda de los caminos por tareas ocultas desde cada tarea (protección ante cronogramas con ciclos)
_MAX_STEPS = 200_000


@dataclass
class NewLink:
    """Vínculo que se crea en la planificación del cliente."""

    pred: Task
    succ: Task
    kind: int
    lag: float                      # minutos laborables (negativa = adelanto)
    origin: str                     # ORIGIN_SUMMARY o ORIGIN_BRIDGE
    hidden: int = 0                 # puentes: tareas ocultas del camino
    adjusted: bool = False          # la posición se ajustó para no retrasar a la sucesora
    from_dates: bool = False        # el camino tiene una posición transcurrida o en %: posición de las fechas
    slack: float | None = None      # minutos entre la fecha que impone el vínculo y la actual de la sucesora


@dataclass
class SkippedLink:
    pred: Task
    succ: Task
    reason: str                     # SKIP_LINKED, SKIP_SUMMARY o SKIP_CYCLE


@dataclass
class SpanChange:
    """Resumen que se muestra con fechas distintas: las fija una subtarea que no se muestra."""

    task: Task
    start: datetime | None          # fechas de sus subtareas que se muestran
    finish: datetime | None


@dataclass
class ProgressChange:
    """Tarea que se muestra con otro % completado que en la planificación interna."""

    task: Task                      # tarea de la planificación interna (su % completado es el interno)
    client: float                   # % completado en el archivo del cliente (previsto o leído de MS Project)
    hidden: int = 0                 # resúmenes: tareas de detalle ocultas, con duración, que ya no cuentan en su %
    converted: int = 0              # resúmenes: resúmenes convertidos cuyo peso cambia (duración ≠ la de sus tareas)

    @property
    def reason(self) -> str:
        parts = []
        if self.hidden:
            parts.append(plural(self.hidden, "tarea oculta", "tareas ocultas"))
        if self.converted:
            parts.append(plural(self.converted, "resumen convertido", "resúmenes convertidos") + " (cuentan por su "
                         "duración, no por la de sus subtareas)")
        return "; ".join(parts) or "—"


@dataclass
class DateChange:
    """Tarea del cliente cuyas fechas, después de recalcular en MS Project, no eran las de la planificación interna."""

    task: Task                      # tarea de la planificación interna (fechas originales)
    start: datetime | None          # fechas en el archivo del cliente
    finish: datetime | None
    note: str = ""                  # por qué no se fijó (tarea comenzada, otra delimitación...)


@dataclass
class FileContent:
    """Lo que no son tareas en el archivo del cliente: calendarios, macros, propiedades del archivo y las vistas,
    tablas, filtros y grupos (evm.client_writer). El archivo del cliente es una copia del .mpp: sin este paso llevaría
    todos los calendarios y las macros de la planificación interna."""

    calendars: list[tuple[str, str]] = field(default_factory=list)          # se conservan: (nombre, uso)
    removed_calendars: list[str] = field(default_factory=list)              # sin uso: se quitaron
    failed_calendars: list[tuple[str, str]] = field(default_factory=list)   # sin uso, no se quitaron: (nombre, error)
    macros: list[str] | None = None             # módulos de macros del archivo de origen (None = no se pudo leer)
    removed_macros: list[str] = field(default_factory=list)                 # módulos quitados
    macros_left: list[str] | None = None        # módulos del archivo del cliente guardado (None = no se pudo leer)
    macro_method: str = ""                      # cómo se quitaron: proyecto VBA u Organizador
    macro_error: str = ""                       # por qué no se pudo usar el proyecto VBA
    properties: list[str] = field(default_factory=list)                     # propiedades del archivo con valor
    lists: dict[str, list[str]] = field(default_factory=dict)               # vistas, tablas, filtros y grupos
    checked: bool = False                       # se revisó el contenido (paso de MS Project)

    @property
    def remaining_macros(self) -> list[str]:
        """Módulos que quedan en el archivo del cliente: los leídos del archivo guardado o, si no se pudo leer, los
        del origen que no se quitaron."""
        if self.macros_left is not None:
            return self.macros_left
        return [m for m in (self.macros or []) if m not in self.removed_macros]


@dataclass
class WriteResult:
    """Resultado de la escritura del archivo del cliente con MS Project (evm.client_writer)."""

    path: str = ""                                              # archivo generado
    source_open: bool = False                                   # el origen estaba abierto en MS Project
    fixed: list[Task] = field(default_factory=list)             # fijadas con «No comenzar antes del»
    differences: list[DateChange] = field(default_factory=list)  # fechas que siguen distintas
    rejected: list[tuple[NewLink, str]] = field(default_factory=list)   # vínculos que MS Project no aceptó
    removed_links: list[tuple[int, int]] = field(default_factory=list)  # creados por MS Project al borrar tareas
    restore_problems: list[tuple[Task, str]] = field(default_factory=list)  # resúmenes convertidos sin restaurar
    field_name_removed: bool = True                             # se quitó el nombre (alias) del campo del cliente
    rounds: int = 0                                             # cálculos de MS Project hasta fijar las fechas
    summary_percent: dict[int, float] = field(default_factory=dict)  # % de cada resumen según MS Project (UID)
    progress_tasks: list[ProgressChange] = field(default_factory=list)  # tareas que no conservan su % completado
    # actividad_interna: tareas que quedan como «Otras tareas», datos que MS Project no dejó quitar (campo -> tareas y
    # primer error), campos calculados con fórmula que no se vacían (campo -> tareas), si se quitaron los códigos
    # EDT y si se aplicó el filtro de la vista (None = no se pidió)
    anonymized: int = 0
    not_cleared: dict[str, tuple[int, str]] = field(default_factory=dict)
    calculated: dict[str, int] = field(default_factory=dict)
    codes_removed: bool = True
    filter_applied: bool | None = None
    content: FileContent = field(default_factory=FileContent)    # calendarios, macros y demás contenido del archivo
    comparison: "Comparison | None" = None   # archivo del cliente frente a la planificación interna (evm.client_check)

    def issues(self, plan: "ClientPlan") -> list[Issue]:
        """Observaciones de la escritura del archivo del cliente."""
        out: list[Issue] = []
        if self.source_open:
            out.append(Issue("CLI-012", message="El archivo de origen está abierto en MS Project: se usó la versión "
                                                "guardada en el disco (los cambios sin guardar no están en el archivo "
                                                "del cliente)."))
        if self.restore_problems:
            what = plural(len(self.restore_problems), "resumen convertido", "resúmenes convertidos")
            out.append(Issue("CLI-019", message=f"{what} con datos que MS Project no aceptó (avance o línea base): "
                                                "ver el informe de control."))
        if self.removed_links:
            what = plural(len(self.removed_links), "par", "pares")
            out.append(Issue("CLI-016", message=f"MS Project vinculó {what} de tareas al borrar las tareas ocultas: "
                                                "esos vínculos se quitaron."))
        if self.rejected:
            what = plural(len(self.rejected), "vínculo nuevo", "vínculos nuevos")
            out.append(Issue("CLI-015", message=f"MS Project no aceptó {what}: ver el informe de control."))
        if not self.field_name_removed:
            out.append(Issue("CLI-017", message=f"No se pudo quitar el nombre del campo '{plan.field_name}' en el "
                                                "archivo del cliente (sus valores sí se vaciaron)."))
        if self.fixed:
            what = plural(len(self.fixed), "tarea", "tareas")
            out.append(Issue("CLI-013", message=f"{what} con el comienzo fijado con «No comenzar antes del», para que "
                                                "tengan la fecha de la planificación interna."))
        if self.differences:
            what = plural(len(self.differences), "tarea tiene", "tareas tienen")
            out.append(Issue("CLI-014", message=f"{what} fechas distintas de la planificación interna: ver el informe "
                                                "de control."))
        if self.progress_tasks:
            what = plural(len(self.progress_tasks), "tarea no conserva", "tareas no conservan")
            out.append(Issue("CLI-021", message=f"{what} su % completado de la planificación interna en el archivo "
                                                "del cliente: ver el informe de control."))
        if self.not_cleared:
            fields = ", ".join(f"{name} ({n})" for name, (n, _error) in sorted(self.not_cleared.items()))
            out.append(Issue("CLI-027", message=f"MS Project no dejó quitar algunos datos de las «{INTERNAL_NAME}» "
                                                f"(campo y tareas): {fields}. Ver el informe de control."))
        if self.calculated:
            fields = ", ".join(f"{name} ({n})" for name, n in sorted(self.calculated.items()))
            out.append(Issue("CLI-030", message=f"Campos calculados con una fórmula en las «{INTERNAL_NAME}» (campo y "
                                                f"tareas): {fields}. No se pueden vaciar: MS Project calcula su valor "
                                                "con los datos que se conservan."))
        if self.filter_applied is False:
            out.append(Issue("CLI-028", message="No se pudo crear el filtro de la vista ([cliente] filtrar_vista): el "
                                                "archivo del cliente se abre con todas las tareas, también las "
                                                f"«{INTERNAL_NAME}»."))
        c = self.content
        removed = []
        if c.removed_calendars:
            removed.append(plural(len(c.removed_calendars), "calendario sin uso", "calendarios sin uso"))
        if c.removed_macros:
            removed.append("las macros (" + ", ".join(c.removed_macros) + ")")
        if removed:
            out.append(Issue("CLI-032", message="Se quitaron del archivo del cliente " + " y ".join(removed) + ": ver "
                                                "el informe de control."))
        left = []
        if c.failed_calendars:
            left.append(plural(len(c.failed_calendars), "calendario sin uso", "calendarios sin uso"))
        if c.remaining_macros:
            left.append("macros (" + ", ".join(c.remaining_macros) + ")")
        if left:
            out.append(Issue("CLI-033", message="MS Project no dejó quitar del archivo del cliente " + " ni ".join(left)
                                                + ": se pueden quitar con el Organizador (Archivo > Información > "
                                                  "Organizador). Ver el informe de control."))
        if self.comparison is not None:
            out += self.comparison.issues()
        return out


@dataclass
class ClientPlan:
    """Cambios para obtener la planificación del cliente a partir de la planificación interna."""

    project: ProjectData
    scope: Scope
    field_name: str
    new_name: str                               # nombre de la tarea de la planificación en el archivo del cliente
    tasks: list[Task]                           # tareas de la planificación (sin su tarea raíz), en orden
    shown: set[int]                             # UID de las tareas que se muestran
    marks: Counter                              # valor del campo de cada tarea: MARK_YES, MARK_NO...
    unknown: list[Task] = field(default_factory=list)           # valor no reconocido (no se muestran)
    orphans: list[Task] = field(default_factory=list)           # Sí debajo de una tarea sin Sí (las de más arriba)
    orphan_count: int = 0                                       # tareas con Sí que no se muestran
    unmarked: list[Task] = field(default_factory=list)          # activas, sin valor, debajo de una que se muestra
    outside: list[Task] = field(default_factory=list)           # tareas fuera de la planificación (se quitan)
    collapsed: list[Task] = field(default_factory=list)         # resúmenes que pasan a ser tareas
    kept_links: int = 0                         # vínculos entre tareas que se muestran (se conservan)
    kept_pairs: set[tuple[int, int]] = field(default_factory=set)  # (UID predecesora, UID sucesora) de esos vínculos
    removed_links: int = 0                      # vínculos con tareas ocultas o dentro de un resumen convertido
    outside_links: int = 0                      # vínculos con tareas fuera de la planificación
    new_links: list[NewLink] = field(default_factory=list)
    skipped: list[SkippedLink] = field(default_factory=list)
    span_changes: list[SpanChange] = field(default_factory=list)
    progress_changes: list[ProgressChange] = field(default_factory=list)   # resúmenes con otro % completado
    # actividad_interna: las tareas que no se muestran quedan como «Otras tareas» (False = se borran)
    internal_activity: bool = False
    dropped: list[Task] = field(default_factory=list)           # actividad_interna: ocultas inactivas (se quitan)
    codes_removed: bool = True                                  # actividad_interna: «Otras tareas» sin código EDT
    comparison: "Comparison | None" = None      # previsión del % del proyecto y del libro EVM (evm.client_check)

    @property
    def root(self) -> Task | None:
        return self.scope.plan

    @property
    def hidden(self) -> list[Task]:
        return [t for t in self.tasks if t.uid not in self.shown]

    @property
    def anonymous(self) -> list[Task]:
        """actividad_interna: tareas que no se muestran y quedan como «Otras tareas»."""
        if not self.internal_activity:
            return []
        gone = {t.uid for t in self.dropped}
        return [t for t in self.hidden if t.uid not in gone]

    @property
    def hidden_branches(self) -> list[Task]:
        """La tarea de más arriba de cada rama que no se muestra."""
        return [t for t in self.hidden if t.parent is None or t.parent is self.root or t.parent.uid in self.shown]

    @property
    def remaining(self) -> set[int]:
        """UID de las tareas del archivo del cliente (con la tarea de la planificación)."""
        root = {self.root.uid} if self.root is not None else set()
        if self.internal_activity:
            gone = {t.uid for t in self.dropped}
            return {t.uid for t in self.tasks if t.uid not in gone} | root
        return set(self.shown) | root

    @property
    def delete_roots(self) -> list[Task]:
        """Tareas que se borran en MS Project (al borrar una tarea se borra su rama): las de otras planificaciones de
        primer nivel y la tarea de más arriba de cada rama oculta (con actividad_interna, de cada rama oculta e
        inactiva)."""
        top = [t for t in self.outside if t.parent is None]
        if self.internal_activity:
            gone = {t.uid for t in self.dropped}
            return top + [t for t in self.dropped if t.parent is None or t.parent.uid not in gone]
        return top + self.hidden_branches

    def ref(self, t: Task) -> str:
        """Tarea en el informe, sin su nombre: código EDT y UID (o número de esquema y UID, si no tiene código)."""
        code = self.scope.raw_code(t)
        return f"{code} (UID {t.uid})" if code else f"UID {t.uid} (esquema {t.outline_number or '?'})"

    def issues(self) -> list[Issue]:
        """Observaciones de la consola y del informe de control."""
        out: list[Issue] = []
        if self.unknown:
            values = Counter(t.client_mark.strip() for t in self.unknown)
            shown = ", ".join(f"«{v}» ({n})" for v, n in values.most_common(5))
            out.append(Issue("CLI-003", message=f"{plural(len(self.unknown), 'tarea tiene', 'tareas tienen')} un valor "
                                                f"no reconocido en '{self.field_name}': {shown}. No se muestran."))
        if self.orphans:
            what = plural(self.orphan_count, "tarea con Sí no se muestra", "tareas con Sí no se muestran")
            where = ", ".join(self.ref(t) for t in self.orphans[:10]) + (" y más" if len(self.orphans) > 10 else "")
            out.append(Issue("CLI-004", message=f"{what} porque una tarea superior no tiene Sí: {where}."))
        if self.unmarked:
            n = len(self.unmarked)
            branch = sum(sum(1 for _ in iter_subtree(t)) for t in self.unmarked)
            what = plural(n, "tarea no tiene", "tareas no tienen")
            sub = f" ({branch} con sus subtareas)" if branch > n else ""
            if self.internal_activity:
                fate = f"queda como «{INTERNAL_NAME}»" if n == 1 else f"quedan como «{INTERNAL_NAME}»"
            else:
                fate = "no se muestra" if n == 1 else "no se muestran"
            where = ", ".join(self.ref(t) for t in self.unmarked[:10]) + (" y más" if n > 10 else "")
            them = "la" if n == 1 else "las"
            out.append(Issue("CLI-031", message=f"{what} valor en '{self.field_name}' y su tarea superior se muestra"
                                                f"{sub}: {fate}. Si el cliente {them} debe ver, hay que poner Sí; si "
                                                f"no, No (así no se avisa): {where}."))
        if self.internal_activity and self.hidden:
            what = plural(len(self.anonymous), "tarea que no se muestra queda", "tareas que no se muestran quedan")
            code = "sin nombre ni código EDT" if self.codes_removed else "sin nombre"
            gone = (f"; {plural(len(self.dropped), 'tarea inactiva se quita', 'tareas inactivas se quitan')}"
                    if self.dropped else "")
            out.append(Issue("CLI-022", message=f"{what} como «{INTERNAL_NAME}» ({code}), con sus fechas, su avance y "
                                                f"sus vínculos{gone}."))
            if not self.codes_removed:
                out.append(Issue("CLI-027", message="El campo EDT configurado es el de MS Project (EDT o número de "
                                                    f"esquema): las «{INTERNAL_NAME}» conservan su código EDT."))
        if self.collapsed:
            out.append(Issue("CLI-005", message=f"{plural(len(self.collapsed), 'resumen pasa', 'resúmenes pasan')} a "
                                                "ser tareas: no se muestra ninguna de sus subtareas."))
        if self.new_links:
            bridges = sum(1 for x in self.new_links if x.origin == ORIGIN_BRIDGE)
            adjusted = sum(1 for x in self.new_links if x.adjusted or x.from_dates)
            out.append(Issue("CLI-006", message=f"Vínculos nuevos: {len(self.new_links)} ({bridges} puentes por tareas "
                                                f"ocultas y {len(self.new_links) - bridges} pasados a un resumen "
                                                f"convertido); {adjusted} con la posición de las fechas actuales."))
        if any(s.reason != SKIP_LINKED for s in self.skipped):
            reasons = Counter(s.reason for s in self.skipped if s.reason != SKIP_LINKED)
            out.append(Issue("CLI-007", message="Vínculos que no se crean: " + ", ".join(
                f"{n} porque {r}" if r == SKIP_CYCLE else f"{n} entre {r}" for r, n in reasons.items()) + "."))
        if self.progress_changes:
            what = plural(len(self.progress_changes), "resumen tendrá", "resúmenes tendrán")
            out.append(Issue("CLI-020", message=f"{what} otro % completado en el archivo del cliente: MS Project lo "
                                                "calcula solo con las tareas que se muestran, sin el avance de las "
                                                "tareas ocultas (ver el informe de control)."))
        if self.span_changes:
            out.append(Issue("CLI-008", message=f"{plural(len(self.span_changes), 'resumen cambia', 'resúmenes cambian')} "
                                                "de fechas: las fija una subtarea que no se muestra."))
        if self.comparison is not None:
            out += self.comparison.issues()
        return out


# ----------------------------------------------------------------------------
# Cálculo
# ----------------------------------------------------------------------------
def _mark(value: str) -> tuple[bool, str]:
    """(se muestra por su propio valor, categoría del valor)."""
    text = plain_text(value or "")
    if not text:
        return False, MARK_EMPTY
    if text in _YES:
        return True, MARK_YES
    return False, MARK_NO if text in _NO else MARK_OTHER


def _active(t: Task) -> bool:
    """La tarea y sus tareas superiores están activas (una tarea inactiva no participa en el cálculo de fechas)."""
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


def _is_ancestor(a: Task, b: Task) -> bool:
    n = b.parent
    while n is not None:
        if n is a:
            return True
        n = n.parent
    return False


def _working(link: Link) -> bool:
    return link.lag == 0 or link.lag_format % 32 in _WORKING_FORMATS


def working_minutes(calendar: WorkCalendar, a: datetime, b: datetime) -> float:
    """Minutos laborables de a a b (negativos si b es anterior)."""
    if b >= a:
        return sum(calendar.minutes_by_month(a, b).values())
    return -sum(calendar.minutes_by_month(b, a).values())


def _point(t: Task, point: str) -> datetime | None:
    return t.start if point == START else t.finish


def _from_start(link: NewLink, calendar: WorkCalendar) -> None:
    """Un vínculo con la posición de las fechas actuales que queda con adelanto (fin a comienzo con -40 días, por
    ejemplo, desde un resumen convertido que termina después de la subtarea que fijaba la sucesora) pasa a partir del
    comienzo de la predecesora, con la posición positiva que da la misma fecha. Los adelantos se consideran un defecto
    de la lógica del cronograma (evaluación DCMA)."""
    pred_point, succ_point = POINTS[link.kind]
    target = _point(link.succ, succ_point)
    if pred_point != FINISH or link.pred.start is None or target is None:
        return
    lag = working_minutes(calendar, link.pred.start, target)
    if lag >= 0:
        link.kind, link.lag = KIND_OF[(START, succ_point)], lag


def rollup_percent(leaves: list[Task]) -> float | None:
    """% completado de un resumen como lo calcula MS Project: la suma de la duración real (duración × % completado)
    de sus tareas de detalle dividida por la suma de su duración. Si ninguna tiene duración (solo hitos): 100 si todas
    están completas y 0 si no. None si no tiene tareas de detalle."""
    if not leaves:
        return None
    total = sum(t.duration or 0.0 for t in leaves)
    if total > 0:
        return sum((t.duration or 0.0) * t.percent_complete for t in leaves) / total
    return 100.0 if all(t.percent_complete >= 100 for t in leaves) else 0.0


def removable_codes(wbs_field: str) -> bool:
    """El código EDT se puede quitar de una tarea si está en un campo personalizado (no en los campos EDT o número de
    esquema de MS Project). Con COM, evm.client_writer lo confirma con el campo que encuentra en el archivo."""
    return builtin_kind(wbs_field) is None and not is_spanish_wbs_name(wbs_field)


def _internal_activity(plan: ClientPlan, cfg: Config, in_scope: set[int]) -> None:
    """actividad_interna: las tareas que no se muestran se conservan (como «Otras tareas»), salvo las inactivas, que no
    cuentan en el cálculo de fechas, del % completado ni del libro EVM (con [exportacion] incluir_inactivas, también
    se conservan). Se conservan todos los vínculos entre las tareas que quedan."""
    plan.internal_activity = True
    plan.codes_removed = removable_codes(cfg.wbs_field)
    if not cfg.include_inactive:
        plan.dropped = [t for t in plan.tasks if t.uid not in plan.shown and not _active(t)]
    gone = {t.uid for t in plan.dropped}
    by_uid = {t.uid: t for t in plan.project.tasks}
    kept = in_scope | ({plan.root.uid} if plan.root is not None else set())     # también los de la planificación
    for t in plan.project.tasks:
        for lk in t.predecessors:
            p = by_uid.get(lk.uid)
            if p is None:
                continue
            if t.uid in kept and p.uid in kept:
                if t.uid in gone or p.uid in gone:
                    plan.removed_links += 1
                else:
                    plan.kept_links += 1
                    plan.kept_pairs.add((p.uid, t.uid))
            elif t.uid in kept or p.uid in kept:
                plan.outside_links += 1


@dataclass
class _Candidate:
    pred: Task
    succ: Task
    kind: int
    lag: float
    origin: str
    hidden: int
    working: bool


def build_client_plan(project: ProjectData, cfg: Config) -> ClientPlan:
    """Tareas que se muestran al cliente, resúmenes convertidos y vínculos nuevos de la planificación del cliente."""
    roots = build_tree(project.tasks)
    scope = select_scope(roots, cfg.plan)
    if scope.plan is not None and scope.plan.parent is not None:
        raise EvmError("CLI-001", f"La planificación '{cfg.plan}' no es una tarea de primer nivel del esquema: la "
                                  "planificación del cliente se genera a partir de una tarea de primer nivel (por "
                                  "ejemplo INTERNO) y sus subtareas.")
    tasks = list(scope.tasks())
    in_scope = {t.uid for t in tasks}
    root = scope.plan
    plan = ClientPlan(project=project, scope=scope, field_name=cfg.client_field,
                      new_name=cfg.client_plan_name if root is not None else "", tasks=tasks, shown=set(),
                      marks=Counter())
    plan.outside = [t for t in project.tasks if t.uid not in in_scope and t is not root]

    # Tareas que se muestran: Sí en la tarea y en todas sus tareas superiores
    own: dict[int, bool] = {}
    empty: set[int] = set()
    for t in tasks:
        own[t.uid], category = _mark(t.client_mark)
        plan.marks[category] += 1
        if category == MARK_OTHER:
            plan.unknown.append(t)
        elif category == MARK_EMPTY:
            empty.add(t.uid)

    def walk(nodes: list[Task], parent_shown: bool) -> None:
        for t in nodes:
            shown = parent_shown and own[t.uid]
            if shown:
                plan.shown.add(t.uid)
            elif own[t.uid]:
                plan.orphan_count += 1
                if t.parent is not None and t.parent is not root and not own.get(t.parent.uid, False):
                    plan.orphans.append(t)
            walk(t.children, shown)

    walk(scope.roots, True)
    # Sin valor debajo de una tarea que se muestra: lo habitual es una tarea nueva sin marcar (CLI-031). Las de más
    # arriba de cada rama; las inactivas no cuentan
    plan.unmarked = [t for t in tasks if t.uid in empty and t.parent is not None and t.parent.uid in plan.shown
                     and _active(t)]
    if not plan.shown:
        raise EvmError("CLI-002", f"Ninguna tarea de la planificación tiene Sí en '{cfg.client_field}' (también en sus "
                                  "tareas superiores): no hay nada que mostrar al cliente.")
    if cfg.client_internal_activity:
        _internal_activity(plan, cfg, in_scope)
        return plan
    plan.collapsed = [t for t in tasks if t.uid in plan.shown and t.children
                      and not any(c.uid in plan.shown for c in t.children)]
    collapsed = {t.uid for t in plan.collapsed}

    def rep(t: Task) -> Task | None:
        """Tarea del cliente que representa a t en la lógica: ella misma, el resumen convertido que la contiene o
        None (tarea oculta entre tareas que se muestran, o fuera de la planificación)."""
        if t.uid not in in_scope:
            return None
        if t.uid in plan.shown:
            return t
        a = t.parent
        while a is not None and a.uid in in_scope:
            if a.uid in plan.shown:
                return a if a.uid in collapsed else None
            a = a.parent
        return None

    by_uid = {t.uid: t for t in project.tasks}
    succs: dict[int, list[tuple[Task, Link]]] = defaultdict(list)
    links: list[tuple[Task, Task, Link]] = []
    for t in project.tasks:
        for lk in t.predecessors:
            p = by_uid.get(lk.uid)
            if p is None:
                continue
            if t.uid in in_scope and p.uid in in_scope:
                links.append((p, t, lk))
                succs[p.uid].append((t, lk))
            elif t.uid in in_scope or p.uid in in_scope:
                plan.outside_links += 1

    reps = {t.uid: rep(t) for t in tasks}
    candidates: dict[tuple[int, int], list[_Candidate]] = defaultdict(list)
    existing: set[tuple[int, int]] = set()
    starts: list[tuple[Task, str, Task, str, float, bool]] = []
    for p, s, lk in links:
        rp, rs = reps[p.uid], reps[s.uid]
        if rp is p and rs is s:
            plan.kept_links += 1
            plan.kept_pairs.add((p.uid, s.uid))
            existing.add((p.uid, s.uid))
            continue
        plan.removed_links += 1
        if not (_active(p) and _active(s)) or rp is None:
            continue
        pp, ps = POINTS.get(lk.kind, POINTS[LINK_FS])
        if rs is not None:
            if rs is not rp:
                candidates[(rp.uid, rs.uid)].append(_Candidate(rp, rs, KIND_OF[(pp, ps)], lk.lag, ORIGIN_SUMMARY, 0,
                                                               _working(lk)))
        else:
            starts.append((rp, pp, s, ps, lk.lag, _working(lk)))

    # Puentes: caminos por tareas ocultas (sin representante) desde una tarea que se muestra hasta otra
    by_start: dict[tuple[int, str], list[tuple[Task, str, float, bool]]] = defaultdict(list)
    for a, pa, s, ps, lag, ok in starts:
        by_start[(a.uid, pa)].append((s, ps, lag, ok))
    for (a_uid, pa), firsts in by_start.items():
        a = by_uid[a_uid]
        best: dict[tuple[int, str], float] = {}
        stack = [(s, ps, lag, 1, ok) for s, ps, lag, ok in firsts]
        steps = 0
        while stack and steps < _MAX_STEPS:
            steps += 1
            h, entry, acc, n, ok = stack.pop()
            if best.get((h.uid, entry), float("-inf")) >= acc:
                continue
            best[(h.uid, entry)] = acc
            duration = 0.0 if h.milestone else (h.duration or 0.0)
            for y, lk in succs[h.uid]:
                if not _active(y):
                    continue
                exit_point, entry2 = POINTS.get(lk.kind, POINTS[LINK_FS])
                step = duration if (entry, exit_point) == (START, FINISH) else \
                    -duration if (entry, exit_point) == (FINISH, START) else 0.0
                acc2 = acc + step + lk.lag
                ok2 = ok and _working(lk)
                ry = reps.get(y.uid)
                if ry is not None:
                    if ry is not a:
                        candidates[(a.uid, ry.uid)].append(_Candidate(a, ry, KIND_OF[(pa, entry2)], acc2,
                                                                      ORIGIN_BRIDGE, n, ok2))
                elif y.uid in in_scope:
                    stack.append((y, entry2, acc2, n + 1, ok2))

    # Un vínculo por par de tareas: el que más restringe a la sucesora, sin retrasarla respecto de las fechas actuales
    calendar = calendar_for(project, cfg)
    graph: dict[int, set[int]] = defaultdict(set)
    for p_uid, s_uid in existing:
        graph[p_uid].add(s_uid)

    def reaches(src: int, dst: int) -> bool:
        seen, stack = set(), [src]
        while stack:
            x = stack.pop()
            if x == dst:
                return True
            if x in seen:
                continue
            seen.add(x)
            stack.extend(graph[x])
        return False

    order = {t.uid: k for k, t in enumerate(project.tasks)}
    for (p_uid, s_uid) in sorted(candidates, key=lambda k: (order[k[0]], order[k[1]])):
        cands = candidates[(p_uid, s_uid)]
        p, s = by_uid[p_uid], by_uid[s_uid]
        if (p_uid, s_uid) in existing or (s_uid, p_uid) in existing:
            plan.skipped.append(SkippedLink(p, s, SKIP_LINKED))
            continue
        if _is_ancestor(p, s) or _is_ancestor(s, p):
            plan.skipped.append(SkippedLink(p, s, SKIP_SUMMARY))
            continue
        chosen: NewLink | None = None
        for c in cands:
            pp, ps = POINTS[c.kind]
            a, b = _point(p, pp), _point(s, ps)
            gap = working_minutes(calendar, a, b) if a is not None and b is not None else None
            link = NewLink(p, s, c.kind, c.lag, c.origin, c.hidden)
            if not c.working:
                if gap is None:
                    continue
                link.lag, link.from_dates = gap, True
            if gap is not None:
                link.slack = gap - link.lag
            if chosen is None or (link.slack is not None and (chosen.slack is None or link.slack < chosen.slack)):
                chosen = link
        if chosen is None:
            continue
        if chosen.slack is not None and chosen.slack < 0:
            chosen.lag += chosen.slack          # = minutos laborables entre las dos fechas actuales
            chosen.slack, chosen.adjusted = 0.0, True
        if (chosen.adjusted or chosen.from_dates) and chosen.lag < 0:
            _from_start(chosen, calendar)
        if reaches(s_uid, p_uid):
            plan.skipped.append(SkippedLink(p, s, SKIP_CYCLE))
            continue
        graph[p_uid].add(s_uid)
        existing.add((p_uid, s_uid))
        plan.new_links.append(chosen)

    # Resúmenes que se muestran con fechas distintas: las fija una subtarea oculta
    def leaves(t: Task) -> list[Task]:
        out = []
        for c in t.children:
            if c.uid not in plan.shown or not _active(c):
                continue
            out.extend(leaves(c) if c.uid not in collapsed and c.children else [c])
        return out

    for t in tasks:
        if t.uid not in plan.shown or t.uid in collapsed or not t.children or not _active(t):
            continue
        found = leaves(t)
        first = [x.start for x in found if x.start is not None]
        last = [x.finish for x in found if x.finish is not None]
        if not first or not last:
            continue
        start, finish = min(first), max(last)
        if (start, finish) != (t.start, t.finish):
            plan.span_changes.append(SpanChange(t, start, finish))

    # % completado de los resúmenes que se muestran: MS Project lo calcula con sus tareas de detalle (en el archivo
    # del cliente, las que se muestran y los resúmenes convertidos), sin el avance de las tareas ocultas
    def inner_leaves(t: Task) -> list[Task]:
        out = []
        for c in t.children:
            if _active(c):
                out.extend(inner_leaves(c) if c.children else [c])
        return out

    for t in tasks:
        if t.uid not in plan.shown or t.uid in collapsed or not t.children or not _active(t):
            continue
        client = rollup_percent(leaves(t))
        if client is None or abs(client - t.percent_complete) < 0.5:
            continue
        hidden = sum(1 for x in inner_leaves(t) if x.uid not in plan.shown and (x.duration or 0) > 0
                     and reps.get(x.uid) is None)
        converted = sum(1 for x in leaves(t) if x.uid in collapsed
                        and abs((x.duration or 0.0) - sum(y.duration or 0.0 for y in inner_leaves(x))) >= 1)
        plan.progress_changes.append(ProgressChange(t, client, hidden, converted))
    return plan


# ----------------------------------------------------------------------------
# Informe de control (Markdown)
# ----------------------------------------------------------------------------
def _cell(value) -> str:
    return " ".join(str(value or "").split()).replace("|", "\\|")


def _d(value: datetime | None) -> str:
    return f"{value:%d/%m/%Y}" if value else "—"


def _days(minutes: float, per_day: float) -> str:
    return fmt_number(round(minutes / per_day, 2), trim=True)


def render_control(plan: ClientPlan, source: str, today: date, result: str,
                   issues: list[Issue] | None = None, outcome: WriteResult | None = None) -> str:
    """Informe de control interno de generar-cliente: lo que se quita, se convierte y se vincula, y (outcome) lo que
    hizo MS Project al escribir el archivo. issues: las observaciones que se muestran (sin las silenciadas); por
    defecto, las del cálculo y las de la escritura."""
    p = plan.project
    per_day = p.minutes_per_day or 480.0
    hidden = len(plan.hidden)
    marks = plan.marks
    hidden_detail = (f"No: {marks[MARK_NO]}, sin valor: {marks[MARK_EMPTY]}, otro valor: {marks[MARK_OTHER]}, con Sí "
                     f"debajo de una tarea que no se muestra: {plan.orphan_count}")
    plan_text = "todo el proyecto"
    if plan.root is not None:
        plan_text = plan.scope.prefix + (f" (en el archivo del cliente: {plan.new_name})"
                                         if plan.new_name and plan.new_name != plan.root.name else "")
    if plan.internal_activity:
        mode = (f"Quedan como «{INTERNAL_NAME}», con sus fechas, su avance y sus vínculos (actividad_interna = true): "
                "el % completado y el libro EVM son los de la planificación interna")
    else:
        mode = ("Se borran (actividad_interna = false): el % completado de los resúmenes y del proyecto, y el libro "
                "EVM, pueden cambiar")
    lines = [
        "# Planificación para el cliente: informe de control",
        "",
        "| | |",
        "|---|---|",
        f"| **Origen** | {_cell(Path(source).name)} |",
        f"| **Planificación** | {_cell(plan_text)} |",
        f"| **Campo** | {_cell(plan.field_name)}: solo se muestran las tareas con Sí (y con Sí en sus tareas superiores) |",
        f"| **Tareas que no se muestran** | {mode} |",
        f"| **Fecha de estado** | {_d(p.status_date)} |",
        f"| **Resultado** | {_cell(result)} |",
        f"| **Fecha de emisión** | {today:%d/%m/%Y} |",
        "",
        "> Documento interno: no se entrega al cliente. Las tareas se identifican por su código EDT y su UID (Id. único "
        "de MS Project).",
        "",
        "## Resumen",
        "",
        "| Dato | Cantidad |",
        "|---|---:|",
        f"| Tareas de la planificación | {len(plan.tasks)} |",
        f"| Se muestran | {len(plan.shown)} |",
        f"| No se muestran | {hidden} ({hidden_detail}; incluye sus subtareas) |" if hidden else "| No se muestran | 0 |",
        f"| Sin valor y con la tarea superior que se muestra (sin sus subtareas) | {len(plan.unmarked)} |",
        f"| Tareas de otras planificaciones (se quitan) | {len(plan.outside)} |",
    ]
    if plan.internal_activity:
        summaries = sum(1 for t in plan.anonymous if t.children)
        codes = "sin código EDT" if plan.codes_removed else "con su código EDT"
        lines += [
            f"| Quedan como «{INTERNAL_NAME}» | {len(plan.anonymous)} ({summaries} resúmenes; {codes}) |",
            f"| Tareas que no se muestran inactivas (se quitan) | {len(plan.dropped)} |",
            f"| Vínculos que se conservan | {plan.kept_links} |",
            f"| Vínculos con tareas inactivas que se quitan (desaparecen) | {plan.removed_links} |",
            f"| Vínculos con otras planificaciones (desaparecen) | {plan.outside_links} |",
            "",
        ]
    else:
        lines += _delete_rows(plan)
    return _report_body(plan, lines, per_day, issues, outcome)


def _delete_rows(plan: ClientPlan) -> list[str]:
    """Filas del resumen del informe con actividad_interna = false."""
    bridges = [x for x in plan.new_links if x.origin == ORIGIN_BRIDGE]
    by_kind = Counter(LINK_NAMES[x.kind] for x in plan.new_links)
    return [
        f"| Resúmenes que pasan a ser tareas | {len(plan.collapsed)} |",
        f"| Vínculos que se conservan | {plan.kept_links} |",
        f"| Vínculos que desaparecen con las tareas ocultas o dentro de un resumen convertido | {plan.removed_links} |",
        f"| Vínculos con otras planificaciones (desaparecen) | {plan.outside_links} |",
        f"| Vínculos nuevos | {len(plan.new_links)} ({len(bridges)} puentes y {len(plan.new_links) - len(bridges)} "
        f"pasados a un resumen convertido; {', '.join(f'{k} {n}' for k, n in sorted(by_kind.items())) or 'ninguno'}) |",
        f"| Vínculos nuevos con la posición ajustada a las fechas actuales | "
        f"{sum(1 for x in plan.new_links if x.adjusted or x.from_dates)} |",
        f"| Vínculos nuevos que no se crean porque las dos tareas ya están vinculadas | "
        f"{sum(1 for x in plan.skipped if x.reason == SKIP_LINKED)} |",
        f"| Vínculos nuevos que no se crean (entre un resumen y su subtarea, o cerrarían un ciclo) | "
        f"{sum(1 for x in plan.skipped if x.reason != SKIP_LINKED)} |",
        f"| Resúmenes que cambian de fechas | {len(plan.span_changes)} |",
        "",
    ]


def _report_body(plan: ClientPlan, lines: list[str], per_day: float, issues: list[Issue] | None,
                 outcome: WriteResult | None) -> str:
    from evm.client_check import comparison_lines
    if issues is None:
        issues = plan.issues() + (outcome.issues(plan) if outcome is not None else [])
    if issues:
        lines += ["## Observaciones", ""]
        lines += [f"- **{i.id}** ({i.severity}): {_cell(i.message)}" for i in issues]
        lines.append("")
    comparison = outcome.comparison if outcome is not None and outcome.comparison is not None else plan.comparison
    if comparison is not None:
        lines += comparison_lines(plan, comparison, per_day)
    if plan.internal_activity and plan.hidden:
        gone = {t.uid for t in plan.dropped}
        sizes = {}
        for t in plan.hidden_branches:
            branch = list(iter_subtree(t))
            sizes[t.uid] = (sum(1 for x in branch if x.uid not in gone), sum(1 for x in branch if x.uid in gone))
        lines += ["## Ramas que no se muestran", "",
                  f"La tarea de más arriba de cada rama que no se muestra: en el archivo del cliente, sus tareas activas "
                  f"quedan como «{INTERNAL_NAME}» y las inactivas se quitan.", "",
                  f"| Tarea | Tareas como «{INTERNAL_NAME}» | Tareas inactivas (se quitan) |", "|---|---:|---:|"]
        lines += [f"| {plan.ref(t)} | {sizes[t.uid][0]} | {sizes[t.uid][1]} |" for t in plan.hidden_branches]
        lines.append("")
    if plan.unmarked:
        fate = f"quedan como «{INTERNAL_NAME}»" if plan.internal_activity else "no se muestran"
        fate += ", con sus subtareas"
        lines += ["## Tareas sin valor debajo de una tarea que se muestra", "",
                  f"No tienen valor en '{_cell(plan.field_name)}' y su tarea superior se muestra: {fate}. Lo habitual "
                  "es una tarea nueva que no se marcó. Si el cliente la debe ver, hay que poner Sí; si no, No (así no "
                  "se avisa). Las inactivas no se incluyen.", "",
                  "| Tarea | Tarea superior | Tareas de su rama |", "|---|---|---:|"]
        lines += [f"| {plan.ref(t)} | {plan.ref(t.parent)} | {sum(1 for _ in iter_subtree(t))} |"
                  for t in plan.unmarked]
        lines.append("")
    if plan.orphans:
        lines += ["## Tareas con Sí que no se muestran", "",
                  "Tienen Sí, pero su tarea superior no: no se muestran, con sus subtareas. Si se deben mostrar, hay "
                  "que poner Sí también en la tarea superior.", "",
                  "| Tarea | Tarea superior sin Sí |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {plan.ref(t.parent)} |" for t in plan.orphans]
        lines.append("")
    if plan.unknown:
        lines += ["## Valores no reconocidos", "", "Se tratan como No.", "", "| Tarea | Valor |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {_cell(t.client_mark)} |" for t in plan.unknown]
        lines.append("")
    if plan.collapsed:
        lines += ["## Resúmenes que pasan a ser tareas", "",
                  "No se muestra ninguna de sus subtareas: el resumen pasa a ser una tarea con sus fechas, su duración, "
                  "su avance y sus líneas base, y recibe los vínculos de sus subtareas.", "",
                  "| Tarea | Comienzo | Fin | % completado |", "|---|---|---|---:|"]
        lines += [f"| {plan.ref(t)} | {_d(t.start)} | {_d(t.finish)} | {fmt_number(t.percent_complete, trim=True)} |"
                  for t in plan.collapsed]
        lines.append("")
    if plan.new_links:
        lines += ["## Vínculos nuevos", "",
                  "Posición en días laborables (negativa = adelanto). *Puente*: reemplaza un camino que pasa por "
                  "tareas ocultas. *Resumen*: vínculo de una subtarea que pasa a su resumen convertido. *Ajustada*: "
                  "con la posición calculada, el vínculo retrasaría a la sucesora; se usa la de las fechas actuales.",
                  "", "| Predecesora | Sucesora | Tipo | Posición | Origen | Observación |",
                  "|---|---|---|---:|---|---|"]
        for x in plan.new_links:
            remarks = []
            if x.origin == ORIGIN_BRIDGE:
                remarks.append(plural(x.hidden, "tarea oculta", "tareas ocultas"))
            if x.adjusted:
                remarks.append("ajustada")
            if x.from_dates:
                remarks.append("posición de las fechas (camino con posición transcurrida o en %)")
            lines.append(f"| {plan.ref(x.pred)} | {plan.ref(x.succ)} | {LINK_NAMES[x.kind]} | "
                         f"{_days(x.lag, per_day)} | {x.origin.capitalize()} | {', '.join(remarks)} |")
        lines.append("")
    if plan.skipped:
        lines += ["## Vínculos que no se crean", "", "| Predecesora | Sucesora | Motivo |", "|---|---|---|"]
        lines += [f"| {plan.ref(s.pred)} | {plan.ref(s.succ)} | {s.reason.capitalize()} |" for s in plan.skipped]
        lines.append("")
    if plan.span_changes:
        lines += ["## Resúmenes que cambian de fechas", "",
                  "Sus fechas las fija una subtarea que no se muestra: en el archivo del cliente, el resumen va de la "
                  "primera a la última de sus subtareas que se muestran.", "",
                  "| Resumen | Comienzo | Fin | Comienzo para el cliente | Fin para el cliente |",
                  "|---|---|---|---|---|"]
        lines += [f"| {plan.ref(c.task)} | {_d(c.task.start)} | {_d(c.task.finish)} | {_d(c.start)} | {_d(c.finish)} |"
                  for c in plan.span_changes]
        lines.append("")
    lines += _progress_lines(plan, outcome)
    if outcome is not None:
        lines += _outcome_lines(plan, outcome, per_day)
    return "\n".join(lines)


def _pct(value: float | None) -> str:
    return "—" if value is None else f"{fmt_number(round(value, 1), trim=True)} %"


def _progress_lines(plan: ClientPlan, out: WriteResult | None) -> list[str]:
    """Resúmenes con otro % completado en el archivo del cliente, y tareas que no conservaron el suyo."""
    lines: list[str] = []
    changes = {c.task.uid: c for c in plan.progress_changes}
    if out is not None:              # resúmenes que MS Project calculó distinto de lo previsto
        by_uid = {t.uid: t for t in plan.tasks}
        for uid, pct in out.summary_percent.items():
            if uid not in changes and abs(pct - by_uid[uid].percent_complete) >= 0.5:
                changes[uid] = ProgressChange(by_uid[uid], pct)
    if changes:
        order = {t.uid: k for k, t in enumerate(plan.tasks)}
        lines += ["## Avance de los resúmenes", "",
                  "MS Project calcula el % completado de un resumen con la duración de sus tareas de detalle. En el "
                  "archivo del cliente solo están las que se muestran: el avance de las tareas ocultas no cuenta. Un "
                  "resumen convertido conserva su %, pero en el resumen superior pesa por su duración (de su comienzo "
                  "a su fin), no por la suma de la de sus subtareas.", ""]
        if out is None:
            lines += ["| Resumen | % completado | En el archivo del cliente | Motivo |", "|---|---:|---:|---|"]
            lines += [f"| {plan.ref(c.task)} | {_pct(c.task.percent_complete)} | {_pct(c.client)} | {c.reason} |"
                      for c in sorted(changes.values(), key=lambda c: order[c.task.uid])]
        else:
            lines += ["| Resumen | % completado | Previsto para el cliente | Según MS Project | Motivo |",
                      "|---|---:|---:|---:|---|"]
            for c in sorted(changes.values(), key=lambda c: order[c.task.uid]):
                expected = next((x.client for x in plan.progress_changes if x.task is c.task), None)
                actual = out.summary_percent.get(c.task.uid)
                lines.append(f"| {plan.ref(c.task)} | {_pct(c.task.percent_complete)} | {_pct(expected)} | "
                             f"{_pct(actual)} | {c.reason} |")
        lines.append("")
    if out is not None and out.progress_tasks:
        lines += ["## Tareas que no conservan su % completado", "",
                  "Tareas de detalle o resúmenes convertidos cuyo % completado en el archivo del cliente no es el de "
                  "la planificación interna.", "", "| Tarea | % completado | En el archivo del cliente |",
                  "|---|---:|---:|"]
        lines += [f"| {plan.ref(c.task)} | {_pct(c.task.percent_complete)} | {_pct(c.client)} |"
                  for c in out.progress_tasks]
        lines.append("")
    return lines


def _dt_min(value: datetime | None) -> str:
    return f"{value:%d/%m/%Y %H:%M}" if value else "—"


def _outcome_lines(plan: ClientPlan, out: WriteResult, per_day: float) -> list[str]:
    """Secciones de la escritura con MS Project."""
    source_open = "Sí: se usó la versión guardada" if out.source_open else "No"
    field_name = "Quitado" if out.field_name_removed else "No se pudo quitar"
    lines = ["## Escritura con MS Project", "",
             "| Dato | Resultado |", "|---|---|",
             f"| Archivo del cliente | {_cell(Path(out.path).name)} (de solo lectura) |",
             f"| Origen abierto en MS Project | {source_open} |",
             f"| Cálculos de MS Project | {out.rounds} |",
             f"| Tareas con el comienzo fijado (No comenzar antes del) | {len(out.fixed)} |"]
    if plan.internal_activity:
        filtered = {None: "No (filtrar_vista = false)", True: "Sí: oculta las «Otras tareas»",
                    False: "No se pudo crear"}[out.filter_applied]
        lines += [f"| Tareas que quedan como «{INTERNAL_NAME}» | {out.anonymized} |",
                  f"| Código EDT de las «{INTERNAL_NAME}» | {'Quitado' if out.codes_removed else 'Se conserva'} |",
                  f"| Campos que MS Project no dejó quitar | {len(out.not_cleared)} |",
                  f"| Campos calculados con fórmula (no se vacían) | {len(out.calculated)} |",
                  f"| Filtro de la vista | {filtered} |"]
    else:
        lines += [f"| Tareas con fechas distintas de la planificación interna | {len(out.differences)} |",
                  f"| Vínculos nuevos que MS Project no aceptó | {len(out.rejected)} |"]
    lines += [f"| Vínculos que MS Project creó al borrar tareas (se quitaron) | {len(out.removed_links)} |",
              f"| Nombre del campo del cliente | {field_name} (valores vaciados) |"]
    c = out.content
    if c.checked:
        lines += [f"| Calendarios | {len(c.calendars)} se conservan, {len(c.removed_calendars)} sin uso quitados"
                  + (f", {len(c.failed_calendars)} sin uso que no se pudieron quitar" if c.failed_calendars else "")
                  + " |",
                  f"| Macros | {_cell(_macros_text(c))} |"]
    lines.append("")
    if c.checked:
        lines += _content_lines(c)
    if out.not_cleared:
        lines += [f"### Datos de las «{INTERNAL_NAME}» que no se pudieron quitar", "",
                  "MS Project no aceptó vaciar el campo (por ejemplo, un campo calculado con una fórmula): el cliente "
                  "puede ver su valor.", "", "| Campo | Tareas | Error de MS Project |", "|---|---:|---|"]
        lines += [f"| {_cell(name)} | {n} | {_cell(error)} |" for name, (n, error) in sorted(out.not_cleared.items())]
        lines.append("")
    if out.calculated:
        lines += [f"### Campos calculados de las «{INTERNAL_NAME}»", "",
                  "Tienen una fórmula: MS Project calcula su valor con los datos que se conservan (fechas, avance, "
                  "líneas base...), así que no se pueden vaciar. Si la fórmula usa un campo con información interna, "
                  "el cliente puede ver el resultado.", "", "| Campo | Tareas |", "|---|---:|"]
        lines += [f"| {_cell(name)} | {n} |" for name, n in sorted(out.calculated.items())]
        lines.append("")
    if out.fixed:
        lines += ["### Fechas fijadas", "",
                  "Sin las tareas ocultas, MS Project las adelantaba: se fijó su comienzo con la delimitación «No "
                  "comenzar antes del» en la fecha de la planificación interna.", "",
                  "| Tarea | Comienzo | Fin |", "|---|---|---|"]
        lines += [f"| {plan.ref(t)} | {_dt_min(t.start)} | {_dt_min(t.finish)} |" for t in out.fixed]
        lines.append("")
    if out.differences:
        lines += ["### Fechas distintas de la planificación interna", "",
                  "| Tarea | Comienzo | Fin | Comienzo para el cliente | Fin para el cliente | Motivo |",
                  "|---|---|---|---|---|---|"]
        lines += [f"| {plan.ref(c.task)} | {_dt_min(c.task.start)} | {_dt_min(c.task.finish)} | {_dt_min(c.start)} | "
                  f"{_dt_min(c.finish)} | {_cell(c.note)} |" for c in out.differences]
        lines.append("")
    if out.rejected:
        lines += ["### Vínculos que MS Project no aceptó", "",
                  "| Predecesora | Sucesora | Tipo | Posición | Error de MS Project |", "|---|---|---|---:|---|"]
        lines += [f"| {plan.ref(x.pred)} | {plan.ref(x.succ)} | {LINK_NAMES[x.kind]} | {_days(x.lag, per_day)} | "
                  f"{_cell(error)} |" for x, error in out.rejected]
        lines.append("")
    if out.removed_links:
        by_uid = {t.uid: t for t in plan.project.tasks}

        def ref(uid: int) -> str:
            return plan.ref(by_uid[uid]) if uid in by_uid else f"UID {uid}"

        lines += ["### Vínculos que MS Project creó al borrar tareas", "",
                  "Se quitaron: no son vínculos de la planificación interna.", "",
                  "| Predecesora | Sucesora |", "|---|---|"]
        lines += [f"| {ref(a)} | {ref(b)} |" for a, b in out.removed_links]
        lines.append("")
    if out.restore_problems:
        lines += ["### Resúmenes convertidos con datos sin restaurar", "", "| Tarea | Detalle |", "|---|---|"]
        lines += [f"| {plan.ref(t)} | {_cell(detail)} |" for t, detail in out.restore_problems]
        lines.append("")
    return lines


def _macros_text(c: FileContent) -> str:
    """Resultado de las macros en una frase."""
    left = c.remaining_macros
    if left:
        return "El archivo del cliente las conserva: " + ", ".join(left)
    if c.removed_macros:
        how = f" ({c.macro_method})" if c.macro_method else ""
        checked = "; comprobado en el archivo guardado" if c.macros_left is not None else ""
        return "Quitadas" + how + ": " + ", ".join(c.removed_macros) + checked
    if c.macros_left is not None or c.macros is not None:
        return "No tiene"
    return "No se pudo comprobar" + (f" ({c.macro_error})" if c.macro_error else "")


def _content_lines(c: FileContent) -> list[str]:
    """Calendarios, propiedades, vistas, tablas, filtros y grupos del archivo del cliente."""
    lines = ["### Contenido del archivo del cliente", "",
             "El archivo del cliente es una copia del .mpp: se quitan los calendarios que no usa ninguna tarea ni "
             "recurso (salvo el de MS Project para las líneas base de Project 98) y las macros (resultado en la tabla "
             "anterior). El resto se conserva y se indica aquí.", "",
             "| Calendario | Uso |", "|---|---|"]
    lines += [f"| {_cell(name)} | {_cell(use[:1].upper() + use[1:])} |" for name, use in c.calendars]
    lines += [f"| {_cell(name)} | Sin uso: quitado |" for name in c.removed_calendars]
    lines += [f"| {_cell(name)} | Sin uso: no se pudo quitar ({_cell(error)}) |" for name, error in c.failed_calendars]
    lines += ["", "| Elemento | En el archivo del cliente |", "|---|---|",
              f"| Propiedades del archivo con valor | {_cell(', '.join(c.properties) or 'Ninguna')} (se ven en "
              "Archivo > Información) |"]
    for label, names in c.lists.items():
        lines.append(f"| {label} | {len(names)}: {_cell(', '.join(names))} |" if names else f"| {label} | 0 |")
    lines.append("")
    return lines


def write_control(plan: ClientPlan, source: str, path: str | Path, result: str, issues: list[Issue] | None = None,
                  today: date | None = None, outcome: WriteResult | None = None) -> None:
    Path(path).write_text(render_control(plan, source, today or date.today(), result, issues, outcome),
                          encoding="utf-8")


def require_field(cfg: Config) -> None:
    """generar-cliente necesita el campo que indica qué tareas se muestran ([cliente] campo)."""
    if not cfg.client_field:
        raise ConfigError("CFG-054", "Falta el campo que indica qué tareas se muestran al cliente: [cliente] campo "
                                     "(por ejemplo campo = \"Etiqueta Cliente\").")


def require_client_field(project: ProjectData, cfg: Config) -> None:
    """El campo de [cliente] campo debe existir en el archivo (lo indica el lector en fields_found)."""
    if "cliente" not in project.fields_found:
        raise ConfigError("LEC-026", f"No se encontró el campo '{cfg.client_field}' ([cliente] campo) en el archivo de "
                                     "MS Project: no se puede saber qué tareas se muestran al cliente.")
