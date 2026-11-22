"""corregir-linea-base: mueve o borra la línea base de las tareas indicadas por su código EDT (con su rama), para
corregir erratas: tareas guardadas en una línea base sin estar aprobadas, o que faltan en la aprobada.

Guía: docs/lineas-base.md. Reglas:

- Mover (--de N --a M): copia los cinco campos de la línea base N de cada tarea (comienzo, fin, duración, trabajo y
  costo previstos) a la línea base M y deja la N vacía (NA; ND en MS Project en español). Borrar (--borrar N): la
  deja vacía. Una tarea sin línea base N no cambia (LBA-016).
- --tarea sin --rama: solo esa tarea; una tarea resumen sin --rama es un error (se movería sin sus subtareas).
- Las tareas resumen de encima de la rama no se modifican: si una conserva la línea base de origen y ninguna de sus
  subtareas la tiene, el PV la contaría como un bloque (LBA-015); se decide con --interactivo, o se incluye con
  --tarea y --rama. Si sus fechas ya no coinciden con las de sus subtareas, se informa (LBA-017).
- Avisos que piden confirmación (LBA-011 a LBA-015): sin --continuar ni --interactivo no se escribe nada.

plan_branch y check_case calculan los cambios sin MS Project (también con un XML, para --simular); ComWriter los
escribe en el proyecto abierto en MS Project, los vuelve a leer para comprobarlos y agrega el motivo a las Notas.
"""
from __future__ import annotations

import logging
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Callable

from evm.config import Config, ConfigError
from evm.fields import fmt_number
from evm.messages import EvmError, Issue, ProjectAppError
from evm.model import ProjectData, Task
from evm.wbs import Scope, code_key, iter_subtree, plural, strip_prefix

log = logging.getLogger("evm")

UNDO_NAME = "Corregir línea base"
FIELDS = ("Start", "Finish", "Duration", "Work", "Cost")
# Avisos que piden confirmación: sin --continuar (o una decisión en --interactivo) no se escribe nada
CONFIRM = ("LBA-011", "LBA-012", "LBA-013", "LBA-014", "LBA-015")
LIMIT = 8


def label(k: int) -> str:
    return f"línea base {k}"


def _codes_text(items: list[str], limit: int = LIMIT) -> str:
    more = f" y {len(items) - limit} más" if len(items) > limit else ""
    return ", ".join(items[:limit]) + more


def ref(t: Task) -> str:
    """Tarea en los mensajes: código EDT (o Id de MS Project, si no tiene) y nombre."""
    code = (t.wbs or "").strip()
    return f"{code} «{t.name}»" if code else f"Id {t.id} «{t.name}»"


def _code(t: Task) -> str:
    return (t.wbs or "").strip() or f"Id {t.id}"


# ----------------------------------------------------------------------------
# Valores de una línea base
# ----------------------------------------------------------------------------
@dataclass(frozen=True)
class Values:
    """Campos de una línea base de una tarea: fechas, duración y trabajo (en minutos) y costo."""

    start: datetime | None = None
    finish: datetime | None = None
    duration: float = 0.0
    work: float = 0.0
    cost: float = 0.0

    @property
    def empty(self) -> bool:
        """Sin fechas (NA): la tarea no tiene esa línea base."""
        return self.start is None and self.finish is None

    def dates(self) -> str:
        if self.empty:
            return "NA"
        return f"{_day(self.start)} – {_day(self.finish)}"

    def short(self, minutes_per_day: float) -> str:
        if self.empty:
            return "NA"
        return f"{self.dates()} ({_days(self.duration, minutes_per_day)})"

    def full(self, minutes_per_day: float) -> str:
        if self.empty:
            return "NA"
        return (f"{self.dates()}; {_days(self.duration, minutes_per_day)}; "
                f"{fmt_number(round(self.work / 60, 2), trim=True)} h; {fmt_number(round(self.cost, 2), trim=True)}")


EMPTY = Values()


def _day(d: datetime | None) -> str:
    return f"{d:%d/%m/%Y}" if d is not None else "NA"


def _days(minutes: float, per_day: float) -> str:
    days = round((minutes or 0) / (per_day or 480), 2)
    return f"{fmt_number(days, trim=True)} d"


def same(a: Values, b: Values, numbers: bool = True) -> bool:
    """Mismos valores (las fechas al minuto; la duración, el trabajo y el costo con dos decimales)."""
    def d(x, y) -> bool:
        return (x is None) == (y is None) and (x is None or abs((x - y).total_seconds()) < 60)

    def n(x, y) -> bool:
        return abs((x or 0.0) - (y or 0.0)) < 0.01 + 1e-6 * max(abs(x or 0.0), abs(y or 0.0))

    if not (d(a.start, b.start) and d(a.finish, b.finish)):
        return False
    if b.empty or not numbers:
        return True
    return n(a.duration, b.duration) and n(a.work, b.work) and n(a.cost, b.cost)


@dataclass(frozen=True)
class Action:
    src: int
    dst: int | None             # None = borrar

    def text(self) -> str:
        if self.dst is None:
            return f"borrar la {label(self.src)}"
        return f"mover de la {label(self.src)} a la {label(self.dst)}"

    def note(self, motivo: str, when: datetime) -> str:
        what = (f"Línea base {self.src} borrada" if self.dst is None else
                f"Línea base {self.src} movida a línea base {self.dst}")
        return f"{when:%d/%m/%Y} · {what}: {motivo.strip()}"


def parse_action(de: int | None, a: int | None, borrar: int | None) -> Action:
    if borrar is not None:
        if de is not None or a is not None:
            raise ConfigError("LBA-005", "Se indica --borrar N o bien --de N --a M, no los dos.")
        return Action(borrar, None)
    if de is None or a is None:
        raise ConfigError("LBA-005", "Se debe indicar la acción: --de N --a M para mover la línea base N a la M, o "
                                     "--borrar N para borrarla (N y M de 0 a 10).")
    if de == a:
        raise ConfigError("LBA-005", f"--de y --a indican la misma línea base ({de}).")
    return Action(de, a)


# ----------------------------------------------------------------------------
# Lectura de las líneas base
# ----------------------------------------------------------------------------
class XmlStore:
    """Líneas base y notas de las tareas de un XML de MS Project (solo lectura: --simular)."""

    def __init__(self, path: str | Path):
        from evm.reader_xml import duration_minutes
        root = ET.parse(path).getroot()
        ns = root.tag[:root.tag.index("}") + 1] if root.tag.startswith("{") else ""

        def text(el, name: str) -> str:
            x = el.find(ns + name)
            return (x.text or "").strip() if x is not None and x.text else ""

        def date(el, name: str) -> datetime | None:
            value = text(el, name)
            try:
                return datetime.fromisoformat(value) if value else None
            except ValueError:
                return None

        def number(value: str) -> float:
            try:
                return float(value)
            except ValueError:
                return 0.0

        self.data: dict[int, dict[int, Values]] = {}
        self._notes: dict[int, str] = {}
        tasks = root.find(ns + "Tasks")
        for el in tasks.findall(ns + "Task") if tasks is not None else []:
            try:
                uid = int(text(el, "UID"))
            except ValueError:
                continue
            self._notes[uid] = text(el, "Notes")
            slots: dict[int, Values] = {}
            for b in el.findall(ns + "Baseline"):
                k = text(b, "Number")
                if not k.isdigit():
                    continue
                slots[int(k)] = Values(date(b, "Start"), date(b, "Finish"),
                                       duration_minutes(text(b, "Duration")) or 0.0,
                                       duration_minutes(text(b, "Work")) or 0.0,
                                       number(text(b, "Cost")) / 100)       # el costo del MSPDI va en centésimas
            self.data[uid] = slots

    def values(self, uid: int, k: int) -> Values:
        return self.data.get(uid, {}).get(k, EMPTY)

    def notes(self, uid: int) -> str:
        return self._notes.get(uid, "")


def _prop(k: int, name: str) -> str:
    return f"Baseline{'' if k == 0 else k}{name}"


def read_values(com, k: int) -> Values:
    from evm.reader_com import _dt, _number

    def get(name: str):
        try:
            return getattr(com, _prop(k, name))
        except Exception:
            return None

    return Values(_dt(get("Start")), _dt(get("Finish")), _number(get("Duration")) or 0.0,
                  _number(get("Work")) or 0.0, _number(get("Cost")) or 0.0)


def _notes_of(com) -> str:
    try:
        value = com.Notes
    except Exception:
        return ""
    return "" if value is None else str(value)


class ComStore:
    """Líneas base y notas de las tareas del proyecto abierto en MS Project."""

    def __init__(self, session):
        from evm.edt_renumber import _com_tasks
        self.session = session
        self.tasks = _com_tasks(session.project)

    def com(self, task: Task):
        com = self.tasks.get(task.uid)
        if com is None:
            raise ProjectAppError("LBA-010", f"No se encontró en MS Project la tarea {ref(task)} (UID {task.uid}).")
        return com

    def values(self, uid: int, k: int) -> Values:
        com = self.tasks.get(uid)
        return read_values(com, k) if com is not None else EMPTY

    def notes(self, uid: int) -> str:
        com = self.tasks.get(uid)
        return _notes_of(com) if com is not None else ""


class State:
    """Líneas base de cada tarea: los valores de MS Project más los cambios ya escritos o simulados (overlay).
    Saber si una tarea tiene una línea base no necesita leer MS Project: el lector ya trae las fechas de cada
    línea base."""

    def __init__(self, project: ProjectData, store):
        self.tasks = {t.uid: t for t in project.tasks}
        self.store = store
        self.overlay: dict[tuple[int, int], Values] = {}
        self._cache: dict[tuple[int, int], Values] = {}

    def values(self, uid: int, k: int) -> Values:
        if (uid, k) in self.overlay:
            return self.overlay[(uid, k)]
        if (uid, k) not in self._cache:
            self._cache[(uid, k)] = self.store.values(uid, k)
        return self._cache[(uid, k)]

    def dates(self, uid: int, k: int) -> tuple[datetime | None, datetime | None] | None:
        if (uid, k) in self.overlay:
            v = self.overlay[(uid, k)]
            return None if v.empty else (v.start, v.finish)
        got = self.tasks[uid].baseline_dates.get(k)
        return got if got and (got[0] is not None or got[1] is not None) else None

    def has(self, uid: int, k: int) -> bool:
        return self.dates(uid, k) is not None

    def apply(self, changes: list["Change"]) -> None:
        for c in changes:
            self.overlay[(c.task.uid, c.slot)] = c.after

    def revert(self, changes: list["Change"]) -> None:
        for c in reversed(changes):
            self.overlay[(c.task.uid, c.slot)] = c.before

    def notes(self, uid: int) -> str:
        return self.store.notes(uid)


def baseline_counts(project: ProjectData) -> dict[int, int]:
    """Tareas con fechas en cada línea base (0 a 10)."""
    counts = {k: 0 for k in range(11)}
    for t in project.tasks:
        for k, (s, f) in t.baseline_dates.items():
            if 0 <= k <= 10 and (s is not None or f is not None):
                counts[k] += 1
    return counts


# ----------------------------------------------------------------------------
# Plan de los cambios
# ----------------------------------------------------------------------------
@dataclass
class Change:
    task: Task
    slot: int
    before: Values
    after: Values


@dataclass
class Case:
    """Un caso: la tarea de --tarea con su rama, o una tarea resumen de encima (--interactivo)."""

    kind: str                   # "rama" o "resumen"
    top: Task
    action: Action
    members: int = 0            # tareas revisadas
    changes: list[Change] = field(default_factory=list)
    skipped: list[Task] = field(default_factory=list)       # sin la línea base de origen: no cambian
    issues: list[Issue] = field(default_factory=list)
    pv_changed: list[str] = field(default_factory=list)     # filas del libro cuyo PV cambia
    notes_before: dict[int, str] = field(default_factory=dict)
    decision: str = ""          # opción elegida en --interactivo
    result: str = ""            # qué se hizo

    @property
    def confirm(self) -> list[Issue]:
        return [i for i in self.issues if i.id in CONFIRM]

    @property
    def tasks(self) -> list[Task]:
        """Tareas que cambian, sin repetir."""
        seen: dict[int, Task] = {}
        for c in self.changes:
            seen.setdefault(c.task.uid, c.task)
        return list(seen.values())

    def title(self) -> str:
        if self.kind == "resumen" and self.action.dst == self.action.src:
            return f"tarea resumen {ref(self.top)}: conservar la {label(self.action.src)}"
        if self.kind == "resumen":
            return f"tarea resumen {ref(self.top)}: {self.action.text()}"
        rest = f" y su rama ({plural(self.members, 'tarea', 'tareas')})" if self.members > 1 else ""
        return f"{ref(self.top)}{rest}: {self.action.text()}"


def find_task(scope: Scope, code: str) -> Task:
    """Tarea de --tarea: la única de la planificación con ese código EDT (sin el prefijo ni los ceros finales)."""
    raw = strip_prefix(code, scope.prefix)
    key = code_key(raw)
    hits = [t for t in scope.tasks() if scope.raw_code(t) and code_key(scope.raw_code(t)) == key] if key else []
    where = f" de la planificación {scope.prefix}" if scope.prefix else ""
    if not hits:
        raise ConfigError("LBA-001", f"--tarea {code}: ninguna tarea{where} tiene ese código EDT. Se deben revisar el "
                                     "código y la planificación elegida (--planificacion).")
    if len(hits) > 1:
        ids = ", ".join(f"Id {t.id}" for t in hits[:10])
        raise ConfigError("LBA-002", f"--tarea {code}: {len(hits)} tareas{where} tienen ese código EDT ({ids}). Se "
                                     "debe elegir la planificación con --planificacion o corregir el código "
                                     "duplicado.")
    return hits[0]


def _inside(t: Task, ancestor: Task) -> bool:
    n = t.parent
    while n is not None:
        if n is ancestor:
            return True
        n = n.parent
    return False


def find_targets(scope: Scope, codes: list[str], branch: bool) -> list[Task]:
    found: list[Task] = []
    for code in codes:
        t = find_task(scope, code)
        for other in found:
            if t is other or _inside(t, other) or _inside(other, t):
                outer = other if t is other or _inside(t, other) else t
                raise ConfigError("LBA-003", f"--tarea {code}: la tarea {ref(t)} está en la rama de {ref(outer)}, "
                                             "que ya se corrige. Se debe indicar solo la tarea superior.")
        if t.children and not branch:
            raise ConfigError("LBA-004", f"--tarea {code}: {ref(t)} es una tarea resumen. Para corregirla con todas "
                                         "sus subtareas se usa --rama; sin --rama se movería la tarea resumen sin "
                                         "ellas.")
        found.append(t)
    return found


def plan_branch(state: State, top: Task, branch: bool, action: Action, kind: str = "rama") -> Case:
    """Cambios de la tarea (y de su rama con --rama): la línea base de origen queda vacía y, al mover, el destino
    recibe una copia exacta. Las tareas sin la línea base de origen no cambian."""
    case = Case(kind, top, action)
    members = list(iter_subtree(top)) if branch else [top]
    case.members = len(members)
    for t in members:
        if not state.has(t.uid, action.src):
            case.skipped.append(t)
            continue
        before = state.values(t.uid, action.src)
        case.changes.append(Change(t, action.src, before, EMPTY))
        if action.dst is not None:
            case.changes.append(Change(t, action.dst, state.values(t.uid, action.dst), before))
    return case


@dataclass
class Context:
    """Datos del proyecto que necesitan los avisos: la línea base del PV y la aprobada."""

    project: ProjectData
    cfg: Config
    pv: int                     # línea base que usa el PV
    approved: int | None = None

    def pv_rows(self, state: State) -> dict[str, frozenset[int]] | None:
        """Tareas que forman el PV de cada fila del libro con los cambios de state (None si no se puede calcular:
        proyecto sin línea base del PV, --pv-fechas-actuales o error de los datos)."""
        if not self.project.baseline_found or self.cfg.pv_current_dates:
            return None
        from evm.wbs import build_part
        n = self.pv
        tasks = []
        for t in self.project.tasks:
            c = replace(t, parent=None, children=[])
            v = state.overlay.get((t.uid, n))
            if v is not None:
                c.baseline_start, c.baseline_finish = v.start, v.finish
                c.has_baseline = not v.empty
                c.baseline_duration = None if v.empty else v.duration
            tasks.append(c)
        try:
            part = build_part(replace(self.project, tasks=tasks, warnings=[]), self.cfg)
        except EvmError as exc:
            log.debug("Efecto en el PV no calculado: %s", exc)
            return None
        return {row.code: frozenset(t.uid for t in row.pv_leaves) for row in part.rows
                if row.is_input and row.pv_leaves is not None}


def summary_above(state: State, top: Task, src: int) -> Task | None:
    """La tarea resumen más cercana por encima de top que tiene la línea base src, si ninguna de sus subtareas la
    tiene: el PV la contaría como un bloque con sus fechas (wbs.pv_selection)."""
    a = top.parent
    while a is not None and not state.has(a.uid, src):
        a = a.parent
    if a is None:
        return None
    if any(state.has(d.uid, src) for d in iter_subtree(a) if d is not a):
        return None
    return a


def _span(state: State, task: Task, k: int) -> tuple[datetime | None, datetime | None] | None:
    """Comienzo y fin de la línea base k según las subtareas de task (None si ninguna la tiene): las de detalle, o la
    tarea resumen con línea base cuyas subtareas no la tienen (como el PV). Las tareas resumen intermedias no cuentan:
    también pueden haber quedado desfasadas."""
    starts, finishes = [], []

    def below(t: Task) -> bool:
        return any(state.has(c.uid, k) or below(c) for c in t.children)

    for d in iter_subtree(task):
        if d is task or below(d):
            continue
        got = state.dates(d.uid, k)
        if got is not None:
            starts += [got[0]] if got[0] is not None else []
            finishes += [got[1]] if got[1] is not None else []
    if not starts and not finishes:
        return None
    return (min(starts) if starts else None, max(finishes) if finishes else None)


def check_case(case: Case, state: State, ctx: Context, *, blocks: bool = True) -> None:
    """Avisos del caso (con state antes del cambio): los que piden confirmación y los informativos. blocks: avisa de
    la tarea resumen de encima que quedaría como un bloque (LBA-015); --interactivo la pregunta aparte."""
    a = case.action
    case.issues = []
    case.pv_changed = []

    def issue(id: str, message: str, task: Task | None = None) -> None:
        t = task or case.top
        case.issues.append(Issue(id, (t.wbs or "").strip(), t.name, message, uid=t.uid))

    if case.skipped and case.kind == "rama":
        issue("LBA-016", f"{plural(len(case.skipped), 'tarea de la rama no tiene', 'tareas de la rama no tienen')} la "
                         f"{label(a.src)}: no {'cambia' if len(case.skipped) == 1 else 'cambian'} "
                         f"({_codes_text([_code(t) for t in case.skipped])}).")
    if not case.changes:
        return
    if a.dst is not None:
        full = [c.task for c in case.changes if c.slot == a.dst and not c.before.empty]
        if full:
            issue("LBA-012", f"La {label(a.dst)} ya tiene valores en {plural(len(full), 'tarea', 'tareas')} "
                             f"({_codes_text([_code(t) for t in full])}): se reemplazan (los anteriores quedan en el "
                             "informe).")
        if a.dst == ctx.pv:
            issue("LBA-013", f"La {label(a.dst)} es la que usa el PV: las tareas entran en el PV, lo que equivale a "
                             "aprobar el cambio. Se confirma si el cliente lo aprobó.")
    if ctx.pv in (a.src, a.dst):
        before = ctx.pv_rows(state)
        state.apply(case.changes)
        try:
            after = ctx.pv_rows(state) if before is not None else None
        finally:
            state.revert(case.changes)
        if before is not None and after is not None:
            lost = [code for code, units in before.items() if units and not after.get(code)]
            case.pv_changed = [code for code in before if before[code] != after.get(code, frozenset())] + \
                [code for code in after if code not in before and after[code]]
            if lost:
                issue("LBA-014", f"{'La fila' if len(lost) == 1 else 'Las filas'} {_codes_text(lost)} se "
                                 f"{'queda' if len(lost) == 1 else 'quedan'} sin ninguna tarea con la línea base del "
                                 f"PV: no {'tendrá' if len(lost) == 1 else 'tendrán'} PV y su BAC debe ser 0 hasta la "
                                 "aprobación del trabajo.")
            if case.pv_changed:
                issue("LBA-018", f"Cambia el PV de {'la fila' if len(case.pv_changed) == 1 else 'las filas'} "
                                 f"{_codes_text(case.pv_changed)}.")
    state.apply(case.changes)
    try:
        block = summary_above(state, case.top, a.src) if blocks else None
        if block is not None:
            effect = (" El PV la contaría como un bloque con sus fechas." if a.src == ctx.pv else
                      " Si esa línea base pasa a usarse para el PV, la contaría como un bloque con sus fechas.")
            issue("LBA-015", f"La tarea resumen {ref(block)} conserva la {label(a.src)} y ninguna de sus subtareas la "
                             f"tendrá.{effect} Es lo correcto si su línea base está aprobada (la tarea se desglosó "
                             "después); si también es nueva, se debe incluir.", block)
        for k in (a.src, a.dst):
            if k is None:
                continue
            p = case.top.parent
            while p is not None:
                own = state.dates(p.uid, k)
                span = _span(state, p, k) if own is not None else None
                if own is not None and span is not None and (own[0] != span[0] or own[1] != span[1]):
                    issue("LBA-017", f"En la {label(k)}, la tarea resumen {ref(p)} va del {_day(own[0])} al "
                                     f"{_day(own[1])} y sus subtareas, del {_day(span[0])} al {_day(span[1])}: no se "
                                     "modifica. Si se quiere, se corrige en MS Project (Establecer línea base, "
                                     "Resumir líneas base).", p)
                p = p.parent
    finally:
        state.revert(case.changes)


def options_text(i: Issue) -> str:
    """Opciones de un aviso que pide confirmación (consola e informe, fuera de --interactivo)."""
    if i.id == "LBA-012":
        return "--continuar los reemplaza."
    if i.id == "LBA-013":
        return "--continuar si el cliente aprobó el cambio."
    if i.id == "LBA-014":
        return ("--continuar deja la fila sin PV; o ubicar en MS Project las tareas dentro de una fila existente; o "
                "no corregir si el trabajo estaba aprobado.")
    if i.id == "LBA-015":
        return (f"--continuar la conserva (línea base aprobada); si también es nueva, incluirla con --tarea {i.code} "
                "--rama; o decidir caso a caso con --interactivo.")
    if i.id == "LBA-011":
        return "corregir la configuración o MS Project; --continuar sigue igualmente."
    return ""


def approved_issue(ctx: Context) -> Issue | None:
    """LBA-011: la línea base aprobada no es la que usa el PV."""
    if ctx.approved is None or ctx.approved == ctx.pv:
        return None
    where = ("linea_base de la configuración" if ctx.cfg.baseline is not None else
             "la línea base del valor acumulado de MS Project")
    return Issue("LBA-011", message=f"La línea base aprobada es la {ctx.approved}, pero el PV usa la {ctx.pv} "
                                    f"({where}): el PV no sería el aprobado. Se debe indicar linea_base = "
                                    f"{ctx.approved} en la configuración o elegir la línea base {ctx.approved} para el "
                                    "valor acumulado en MS Project (Archivo > Opciones > Avanzado).")


# ----------------------------------------------------------------------------
# Escritura en MS Project
# ----------------------------------------------------------------------------
def append_note(old: str, line: str) -> str:
    """Notas con una línea más al final (MS Project separa las líneas con \\r)."""
    old = (old or "").rstrip("\r\n ")
    return f"{old}\r{line}" if old else line


def _norm_notes(text: str) -> str:
    return "\n".join(x.rstrip() for x in (text or "").replace("\r\n", "\n").replace("\r", "\n").strip().split("\n"))


class ComWriter:
    """Escribe líneas base y notas en el proyecto abierto en MS Project y las vuelve a leer para comprobarlas."""

    def __init__(self, session, store: ComStore):
        self.session = session
        self.store = store
        self.na: str | None = None          # texto de NA que aceptó MS Project («NA» o «ND»)
        self.warnings: list[Issue] = []

    def _put(self, com, task: Task, k: int, v: Values) -> None:
        from evm.reader_com import _com_date, set_na

        def date(name: str, value: datetime | None) -> None:
            prop = _prop(k, name)
            if value is not None:
                setattr(com, prop, _com_date(value))
                return
            got = set_na(com, prop, self.na)
            if got is None:
                raise ProjectAppError("LBA-010", f"MS Project no aceptó dejar vacía (NA o ND) la {label(k)} de "
                                                 f"{ref(task)}.")
            self.na = got

        if v.empty:
            date("Start", None)
            date("Finish", None)
            for name in ("Duration", "Work", "Cost"):
                try:
                    setattr(com, _prop(k, name), 0)
                except Exception as exc:        # p. ej. trabajo de una tarea con recursos: no cuenta en el EVM
                    log.debug("%s de UID %s no se puso a 0: %s", _prop(k, name), task.uid, exc)
            return
        setattr(com, _prop(k, "Duration"), v.duration)
        date("Start", v.start)
        date("Finish", v.finish)
        for name, value in (("Work", v.work), ("Cost", v.cost)):
            try:
                setattr(com, _prop(k, name), value)
            except Exception as exc:
                log.debug("%s de UID %s no se escribió: %s", _prop(k, name), task.uid, exc)

    def _wrong(self, items: list[tuple[Task, int, Values]], numbers: bool) -> list[tuple[Task, int, Values]]:
        return [(t, k, v) for t, k, v in items if not same(read_values(self.store.com(t), k), v, numbers)]

    def write_values(self, items: list[tuple[Task, int, Values]]) -> None:
        """Escribe y comprueba. Si un valor no coincide al volver a leerlo (MS Project recalcula algún campo al
        escribir otro), se escribe otra vez. Las fechas y la duración deben coincidir (LBA-010); el trabajo y el costo
        distintos se informan (LBA-020)."""
        from evm.client_writer import com_error_text
        current = None
        try:
            for item in items:
                current = item
                t, k, v = item
                self._put(self.store.com(t), t, k, v)
            wrong = self._wrong(items, numbers=False)
            for t, k, v in wrong:
                current = (t, k, v)
                self._put(self.store.com(t), t, k, v)
        except ProjectAppError:
            raise
        except Exception as exc:
            where = f" de {ref(current[0])} ({label(current[1])})" if current else ""
            raise ProjectAppError("LBA-010", f"MS Project devolvió un error al escribir la línea base{where}: "
                                             f"{com_error_text(exc)}.") from None
        wrong = self._wrong(items, numbers=False)
        if wrong:
            shown = [f"{_code(t)} ({label(k)}: se esperaba {v.dates()}, se leyó "
                     f"{read_values(self.store.com(t), k).dates()})" for t, k, v in wrong]
            raise ProjectAppError("LBA-010", f"{len(wrong)} valores no coinciden con los previstos al volver a "
                                             f"leerlos: {_codes_text(shown, 5)}.")
        diff = [f"{_code(t)} ({label(k)})" for t, k, v in self._wrong(items, numbers=True) if not v.empty]
        if diff:
            self.warnings.append(Issue("LBA-020", message=f"Duración, trabajo o costo previstos distintos de los "
                                                          f"copiados en {plural(len(diff), 'tarea', 'tareas')} "
                                                          f"({_codes_text(diff)}): MS Project los recalculó. Las "
                                                          "fechas son las copiadas."))

    def write_notes(self, notes: dict[int, tuple[Task, str]]) -> None:
        from evm.client_writer import com_error_text
        for uid, (task, text) in notes.items():
            com = self.store.com(task)
            try:
                com.Notes = text
            except Exception as exc:
                raise ProjectAppError("LBA-010", f"MS Project devolvió un error al escribir las Notas de {ref(task)}: "
                                                 f"{com_error_text(exc)}.") from None
            if _norm_notes(_notes_of(com)) != _norm_notes(text):
                raise ProjectAppError("LBA-010", f"Las Notas de {ref(task)} no coinciden con las previstas al volver "
                                                 "a leerlas.")

    @property
    def app(self):
        return self.session.app

    def begin(self, name: str) -> bool:
        """Abre una acción de deshacer de MS Project (False si no se puede: LBA-019) y detiene el redibujado."""
        from evm.client_writer import com_error_text
        self.session._activate()
        try:
            self.app.OpenUndoTransaction(name)
            return True
        except Exception as exc:
            log.debug("OpenUndoTransaction no disponible: %s", com_error_text(exc))
            if not any(i.id == "LBA-019" for i in self.warnings):
                self.warnings.append(Issue("LBA-019", message="MS Project no permitió agrupar los cambios en una sola "
                                                              "acción de deshacer: cada valor escrito es una acción."))
            return False

    def end(self, grouped: bool) -> None:
        if grouped:
            try:
                self.app.CloseUndoTransaction()
            except Exception as exc:
                log.debug("CloseUndoTransaction: %s", exc)


def case_notes(case: Case, state: State, motivo: str, when: datetime) -> dict[int, tuple[Task, str]]:
    """Notas nuevas de las tareas del caso (guarda las anteriores en case.notes_before)."""
    out = {}
    for t in case.tasks:
        old = state.notes(t.uid)
        case.notes_before[t.uid] = old
        out[t.uid] = (t, append_note(old, case.action.note(motivo, when)))
    return out


def write_cases(writer: ComWriter, cases: list[Case], state: State, motivo: str, when: datetime,
                name: str = UNDO_NAME) -> bool:
    """Escribe los cambios y las notas de los casos en una acción de deshacer. Si MS Project da un error, restaura
    los valores anteriores (y lo indica en el mensaje). Devuelve si los cambios quedaron en una sola acción."""
    changes = [c for case in cases for c in case.changes]
    notes: dict[int, tuple[Task, str]] = {}
    for case in cases:
        notes.update(case_notes(case, state, motivo, when))
    grouped = writer.begin(name)
    try:
        writer.write_values([(c.task, c.slot, c.after) for c in changes])
        writer.write_notes(notes)
    except ProjectAppError as exc:
        restored = restore(writer, cases)
        writer.end(grouped)
        extra = (" Se restauraron los valores anteriores." if restored else
                 " No se pudieron restaurar todos los valores anteriores: se deben revisar las tareas (o cerrar el "
                 "archivo sin guardar).")
        raise ProjectAppError("LBA-010", str(exc) + extra) from None
    writer.end(grouped)
    state.apply(changes)
    return grouped


def restore(writer: ComWriter, cases: list[Case]) -> bool:
    """Vuelve a escribir los valores y las notas anteriores de los casos (deshacer de --interactivo, o tras un error).
    Devuelve si todo quedó como antes."""
    changes = [c for case in cases for c in case.changes]
    ok = True
    try:
        writer.write_values([(c.task, c.slot, c.before) for c in reversed(changes)])
    except ProjectAppError as exc:
        log.debug("Restauración de valores: %s", exc)
        ok = False
    for case in cases:
        try:
            writer.write_notes({t.uid: (t, case.notes_before[t.uid]) for t in case.tasks
                                if t.uid in case.notes_before})
        except ProjectAppError as exc:
            log.debug("Restauración de notas: %s", exc)
            ok = False
    return ok


# ----------------------------------------------------------------------------
# Consola de --interactivo
# ----------------------------------------------------------------------------
class Stop(Exception):
    """El usuario eligió salir (x) en --interactivo."""


def console_available() -> bool:
    """Hay una consola en la que responder (--interactivo)."""
    import sys
    try:
        return sys.stdin is not None and sys.stdin.isatty()
    except (AttributeError, ValueError):
        return False


class Console:
    """Preguntas de --interactivo. ask: función que lee una respuesta (input); out: función que escribe una línea."""

    def __init__(self, ask: Callable[[str], str] | None = None, out: Callable[[str], None] = print):
        self._ask = ask or (lambda question: input(question))
        self.out = out

    def choose(self, question: str, options: list[tuple[str, str]]) -> str:
        for key, text in options:
            self.out(f"    {key}: {text}")
        keys = [k for k, _t in options]
        while True:
            answer = self._ask(f"  {question} [{'/'.join(keys)}]: ").strip().lower()
            if answer in keys:
                return answer
            self.out(f"  Respuesta no válida: se debe escribir {', '.join(keys[:-1])} o {keys[-1]}.")

    def number(self, question: str, valid: list[int], default: int | None = None) -> int:
        hint = f" (Intro = {default})" if default is not None else ""
        while True:
            answer = self._ask(f"  {question}{hint}: ").strip()
            if not answer and default is not None:
                return default
            if answer.isdigit() and int(answer) in valid:
                return int(answer)
            numbers = _codes_text([str(v) for v in valid], 11)
            self.out(f"  Respuesta no válida: se debe escribir un número de {numbers}.")

    def text(self, question: str) -> str:
        while True:
            answer = self._ask(f"  {question}: ").strip()
            if answer:
                return answer


def baseline_lines(project: ProjectData, ctx: Context) -> list[str]:
    """Líneas base del proyecto: tareas con fechas, fecha de guardado y cuál usa el PV y cuál MS Project."""
    counts = baseline_counts(project)
    lines = []
    for k in range(11):
        saved = project.baseline_saved.get(k)
        if not counts[k] and saved is None and k not in (ctx.pv, project.ev_baseline):
            continue
        marks = []
        if k == ctx.pv:
            marks.append("la usa el PV" + (" (linea_base de la configuración)" if ctx.cfg.baseline is not None else ""))
        if k == project.ev_baseline:
            marks.append("configurada en MS Project para el valor acumulado")
        when = f"guardada el {saved:%d/%m/%Y}" if saved is not None else ""
        lines.append(f"  {k:>2}  {plural(counts[k], 'tarea', 'tareas'):<14} {when:<22}"
                     f"{'  <- ' + '; '.join(marks) if marks else ''}".rstrip())
    return lines


def show_case(console: Console, case: Case, number: str, minutes_per_day: float, limit: int = 30) -> None:
    out = console.out
    out("")
    out(f"Caso {number}: {case.title()}")
    a = case.action
    rows = {}
    for c in case.changes:
        rows.setdefault(c.task.uid, (c.task, []))[1].append(c)
    for i, (t, changes) in enumerate(rows.values()):
        if i == limit:
            out(f"    ... y {len(rows) - limit} tareas más (en el informe).")
            break
        parts = [f"LB {c.slot}: {c.before.short(minutes_per_day)} -> {c.after.short(minutes_per_day)}"
                 for c in changes]
        out(f"    {_code(t):<14} {t.name[:40]:<40}  {' | '.join(parts)}")
    if not case.changes:
        out(f"    Ninguna tarea tiene la {label(a.src)}: no hay nada que cambiar.")
    for i in case.issues:
        out(f"  {'[!]' if i.id in CONFIRM or i.id in ('LBA-019', 'LBA-020') else '[i]'} {i.id}: {i.message}")


# ----------------------------------------------------------------------------
# Informe
# ----------------------------------------------------------------------------
def _cell(text: str) -> str:
    return (text or "").replace("|", "\\|").replace("\r\n", " ").replace("\r", " ").replace("\n", " ").strip()


def write_report(path: Path, project: ProjectData, ctx: Context, cases: list[Case], *, origen: str, version: str,
                 action: Action, motivo: str, mode: str, result: str, notices: list[Issue],
                 backup: Path | None = None, scope_prefix: str = "") -> None:
    """Informe local de corregir-linea-base (con nombres de tareas: uso interno)."""
    per_day = project.minutes_per_day
    lines = [f"# Corrección de línea base: {Path(origen).name if origen.lower() != 'activo' else 'proyecto activo'}",
             "",
             f"- Fecha: {datetime.now():%Y-%m-%d %H:%M}",
             f"- Versión del programa: {version}",
             f"- Origen: {origen}"]
    if scope_prefix:
        lines.append(f"- Planificación: {scope_prefix}")
    lines += [f"- Acción: {action.text()}",
              f"- Motivo: {motivo or '(sin motivo: simulación)'}",
              f"- Línea base del PV: {ctx.pv}" + (" (linea_base de la configuración)" if ctx.cfg.baseline is not None
                                                   else " (la del valor acumulado de MS Project)"),
              f"- Línea base aprobada: {ctx.approved if ctx.approved is not None else 'no indicada'}",
              f"- Modo: {mode}",
              f"- Resultado: {result}"]
    if backup is not None:
        lines.append(f"- Copia de respaldo: {backup}")
    lines += ["", "## Líneas base del proyecto", "",
              "| Línea base | Tareas | Guardada el | Nota |", "|---:|---:|---|---|"]
    counts = baseline_counts(project)
    for k in range(11):
        saved = project.baseline_saved.get(k)
        marks = []
        if k == ctx.pv:
            marks.append("la usa el PV")
        if k == project.ev_baseline:
            marks.append("MS Project: valor acumulado")
        if k == ctx.approved:
            marks.append("aprobada")
        if not counts[k] and saved is None and not marks:
            continue
        lines.append(f"| {k} | {counts[k]} | {f'{saved:%d/%m/%Y}' if saved else ''} | {', '.join(marks)} |")
    lines.append("")
    if notices:
        lines += ["## Observaciones", "", "| Código | Severidad | Tarea | Mensaje |", "|---|---|---|---|"]
        for i in notices:
            task = f"{i.code} {i.name}".strip()
            opts = options_text(i) if i.id in CONFIRM else ""
            lines.append(f"| {i.id} | {i.severity} | {_cell(task)} | {_cell(i.message)}"
                         f"{' Opciones: ' + _cell(opts) if opts else ''} |")
        lines.append("")
    lines += ["## Casos", ""]
    if not cases:
        lines.append("Ninguno.")
    for n, case in enumerate(cases, start=1):
        lines += [f"### {n}. {_cell(case.title())}", ""]
        if case.decision:
            lines.append(f"- Decisión: {case.decision}")
        lines.append(f"- Resultado: {case.result or 'sin cambios'}")
        if case.pv_changed:
            lines.append(f"- Filas cuyo PV cambia: {', '.join(case.pv_changed)}")
        if case.skipped:
            lines.append(f"- Sin la {label(case.action.src)} (no cambian): "
                         f"{_cell(', '.join(_code(t) for t in case.skipped))}")
        lines.append("")
        if case.changes:
            lines += ["| Id | Código | Tarea | Línea base | Antes | Después |", "|---:|---|---|---:|---|---|"]
            for c in case.changes:
                lines.append(f"| {c.task.id} | {_cell(_code(c.task))} | {_cell(c.task.name)} | {c.slot} | "
                             f"{c.before.full(per_day)} | {c.after.full(per_day)} |")
            lines.append("")
        old = [(t, case.notes_before[t.uid]) for t in case.tasks if case.notes_before.get(t.uid, "").strip()]
        if old:
            lines += ["Notas anteriores:", "", "| Código | Notas |", "|---|---|"]
            lines += [f"| {_cell(_code(t))} | {_cell(text)} |" for t, text in old]
            lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


# ----------------------------------------------------------------------------
# --interactivo: un cambio cada vez
# ----------------------------------------------------------------------------
@dataclass
class Session:
    """Lo que necesita --interactivo para cada caso."""

    console: Console
    ctx: Context
    state: State
    writer: ComWriter
    motivo: str
    when: datetime
    keep: Callable[[list[Issue]], list[Issue]] = lambda issues: issues      # sin las observaciones silenciadas
    cases: list[Case] = field(default_factory=list)


def _written(s: Session, case: Case) -> bool:
    """Escribe el caso y pregunta cómo quedó. 'd' lo deshace (devuelve False: se vuelve a preguntar el caso)."""
    code = (case.top.wbs or "").strip() or f"Id {case.top.id}"
    write_cases(s.writer, [case], s.state, s.motivo, s.when, name=f"{UNDO_NAME} {code}")
    case.result = "escrito en MS Project, sin guardar"
    s.console.out(f"  Escrito en MS Project ({plural(len(case.tasks), 'tarea', 'tareas')}, sin guardar), con el motivo "
                  "en las Notas. Se puede revisar en MS Project antes de seguir.")
    for i in s.writer.warnings:
        if i.id == "LBA-020":
            case.issues.append(i)
    s.writer.warnings = [i for i in s.writer.warnings if i.id != "LBA-020"]
    key = s.console.choose("¿Cómo quedó?", [("s", "correcto, seguir"), ("d", "deshacer este cambio"),
                                            ("x", "salir (se conservan los cambios hechos, sin guardar)")])
    if key == "d":
        if not restore(s.writer, [case]):
            case.result = "escrito; al deshacerlo, no se restauraron todos los valores"
            s.cases.append(case)
            raise ProjectAppError("LBA-010", f"No se pudieron restaurar todos los valores anteriores de "
                                             f"{ref(case.top)}: se deben revisar las tareas en MS Project (o cerrar "
                                             "el archivo sin guardar).")
        s.state.revert(case.changes)
        case.result = "escrito y deshecho"
        s.cases.append(case)
        s.console.out("  Deshecho: las tareas tienen otra vez sus valores y sus Notas anteriores.")
        return False
    s.cases.append(case)
    if key == "x":
        raise Stop
    return True


def branch_case(s: Session, top: Task, branch: bool, action: Action, number: str) -> bool:
    """Caso de una tarea de --tarea (con su rama). Devuelve si quedó escrito."""
    while True:
        case = plan_branch(s.state, top, branch, action)
        check_case(case, s.state, s.ctx, blocks=False)
        case.issues = s.keep(case.issues)
        show_case(s.console, case, number, s.ctx.project.minutes_per_day)
        if not case.changes:
            case.result = "sin cambios"
            s.cases.append(case)
            return False
        key = s.console.choose("¿Aplicar este cambio?", [("s", "aplicar"), ("o", "omitir este caso"), ("x", "salir")])
        case.decision = {"s": "aplicar", "o": "omitir", "x": "salir"}[key]
        if key != "s":
            case.result = "no se aplicó"
            s.cases.append(case)
            if key == "x":
                raise Stop
            return False
        if _written(s, case):
            return True


def summary_cases(s: Session, top: Task, src: int) -> None:
    """Tareas resumen de encima que quedarían como un bloque (LBA-015): se pregunta qué hacer con cada una, de la más
    cercana hacia arriba (si se conserva una, las de encima tienen una subtarea con la línea base)."""
    current = top
    while True:
        block = summary_above(s.state, current, src)
        if block is None:
            return
        while True:
            effect = ("el PV la cuenta como un bloque con sus fechas" if src == s.ctx.pv else
                      "si esa línea base pasa a usarse para el PV, la contaría como un bloque")
            s.console.out("")
            s.console.out(f"Tarea resumen de encima: {ref(block)} conserva la {label(src)} y ninguna de sus "
                          f"subtareas la tiene: {effect}.")
            v = s.state.values(block.uid, src)
            s.console.out(f"    {label(src).capitalize()}: {v.short(s.ctx.project.minutes_per_day)}")
            options = []
            approved = s.ctx.approved
            if approved is not None and approved != src:
                options.append(("a", f"mover a la línea base aprobada ({approved}): no se incluyó antes por error"))
            options += [("b", f"borrar de la {label(src)} (queda en NA/ND): se incluyó por error"),
                        ("m", "mover a otra línea base (se indica el número)"),
                        ("c", "omitir: dejarla como está (su línea base está aprobada)"),
                        ("x", "salir")]
            key = s.console.choose("¿Qué hacer con la tarea resumen?", options)
            if key == "x":
                raise Stop
            if key == "c":
                case = Case("resumen", block, Action(src, src), decision="omitir: se conserva",
                            result="sin cambios: se conserva su línea base")
                s.cases.append(case)
                return
            if key == "m":
                dst = s.console.number("¿A qué línea base se mueve? [0-10]", [k for k in range(11) if k != src])
            else:
                dst = approved if key == "a" else None
            action = Action(src, dst)
            case = plan_branch(s.state, block, False, action, kind="resumen")
            check_case(case, s.state, s.ctx, blocks=False)
            case.issues = s.keep(case.issues)
            case.decision = dict(options)[key]
            show_case(s.console, case, "de una tarea resumen", s.ctx.project.minutes_per_day)
            if case.confirm:
                ok = s.console.choose("¿Aplicar este cambio?", [("s", "aplicar"), ("o", "volver a las opciones"),
                                                                ("x", "salir")])
                if ok == "x":
                    case.result = "no se aplicó"
                    s.cases.append(case)
                    raise Stop
                if ok == "o":
                    continue
            if _written(s, case):
                break
        current = block


def run_interactive(s: Session, targets: list[Task], branch: bool, action: Action) -> None:
    """Recorre los casos: cada tarea de --tarea con su rama y, si queda escrita, las tareas resumen de encima que
    quedarían como un bloque. Termina al acabar los casos o al elegir salir."""
    total = len(targets)
    try:
        for n, top in enumerate(targets, start=1):
            if branch_case(s, top, branch, action, f"{n} de {total}"):
                summary_cases(s, top, action.src)
    except Stop:
        s.console.out("")
        s.console.out("Salida a petición del usuario: los cambios ya escritos se conservan, sin guardar.")
