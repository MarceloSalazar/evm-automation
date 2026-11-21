"""actualizar-edt: numera los códigos EDT de las subtareas de una tarea según la jerarquía de MS Project.

Reglas (docs/comandos.md, sección actualizar-edt):

- La tarea de --desde conserva su código. Sus subtareas se numeran en el orden del esquema de MS Project: los hijos
  <código>.1, <código>.2...; los nietos <código>.1.1... Una tarea sin código recibe uno: su tarea superior lo tiene.
- Tareas protegidas, cuyo código nunca cambia: las de código de nivel 1 a nivel_maximo, las de Visible = Sí y las de
  la lista EDT (son filas del libro: el cliente recibe siempre el mismo código), y las que contienen alguna de ellas.
- Dentro de cada tarea, las subtareas protegidas van primero, con códigos correlativos desde 1 (el cliente no debe
  ver saltos); las demás se numeran a continuación. Si no es así, es un error y no se escribe ningún código.
- Nunca se crea un código de nivel 1 a nivel_maximo (sería una fila nueva del libro): bajo una tarea de nivel menor,
  las subtareas se numeran con ceros hasta nivel_maximo (2.7.0.1 bajo 2.7.0, una actividad sin paquete).
- Un código nuevo que ya tiene una tarea de fuera de las ramas se informa (NUM-014), y la rama se actualiza igual.

plan_renumber calcula los cambios sin MS Project (también con un XML, para --simular); write_codes los escribe en el
proyecto abierto en MS Project y los vuelve a leer para comprobarlos.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from evm import progress
from evm.config import Config, ConfigError
from evm.fields import TASK_TEXT_IDS, parse_visible
from evm.messages import SEV_ERROR, SEV_WARN, Issue, ProjectAppError
from evm.model import ProjectData, Task
from evm.wbs import Scope, build_tree, code_key, code_level, iter_subtree, pad_code, select_scope, strip_prefix

log = logging.getLogger("evm")

UNDO_NAME = "Actualizar códigos EDT"
NATIVE_FIELDS = {"WBS", "OutlineNumber", "Notes"}
PROTECTED_INSIDE = "contiene tareas protegidas"


@dataclass
class Change:
    task: Task
    old: str                    # código EDT antes (tal como está en MS Project; "" = sin código)
    new: str


@dataclass
class Kept:
    task: Task
    reason: str                 # por qué no cambia: "nivel 3", "Visible = Sí", "lista EDT", PROTECTED_INSIDE


@dataclass
class RenumberPlan:
    scope: Scope
    starts: list[Task]
    changes: list[Change] = field(default_factory=list)
    kept: list[Kept] = field(default_factory=list)
    skipped: list[Task] = field(default_factory=list)       # sin numerar (NUM-013), con su rama
    tasks: int = 0              # subtareas de las ramas (sin las tareas de --desde)
    issues: list[Issue] = field(default_factory=list)

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == SEV_ERROR]

    @property
    def added(self) -> int:
        """Tareas que no tenían código EDT."""
        return sum(1 for c in self.changes if not c.old)


def _numeric(code: str) -> bool:
    parts = (code or "").strip().split(".")
    return bool(code and code.strip()) and all(p.strip().isdigit() for p in parts)


def ref(t: Task) -> str:
    """Tarea en los mensajes: código EDT (o Id de MS Project, si no tiene) y nombre."""
    code = (t.wbs or "").strip()
    return f"{code} «{t.name}»" if code else f"Id {t.id} «{t.name}»"


def _codes_text(codes: list[str], limit: int = 8) -> str:
    more = f" y {len(codes) - limit} más" if len(codes) > limit else ""
    return ", ".join(codes[:limit]) + more


def find_start(scope: Scope, code: str) -> Task:
    """Tarea de --desde: la única de la planificación con ese código (sin el prefijo ni los ceros finales)."""
    raw = strip_prefix(code, scope.prefix)
    if not _numeric(raw):
        raise ConfigError("NUM-003", f"--desde {code}: no es un código EDT numérico (p. ej. 2.7.1). Solo se actualizan "
                                     "ramas cuya tarea superior tiene un código numérico.")
    key = code_key(raw)
    hits = [t for t in scope.tasks() if _numeric(scope.raw_code(t)) and code_key(scope.raw_code(t)) == key]
    where = f" de la planificación {scope.prefix}" if scope.prefix else ""
    if not hits:
        raise ConfigError("NUM-001", f"--desde {code}: ninguna tarea{where} tiene ese código EDT. Se deben revisar el "
                                     "código y la planificación elegida (--planificacion).")
    if len(hits) > 1:
        ids = ", ".join(f"Id {t.id}" for t in hits[:10])
        raise ConfigError("NUM-002", f"--desde {code}: {len(hits)} tareas{where} tienen ese código EDT ({ids}). Se "
                                     "debe elegir la planificación con --planificacion o corregir el código "
                                     "duplicado.")
    return hits[0]


def plan_renumber(project: ProjectData, cfg: Config, starts: list[str]) -> RenumberPlan:
    """Calcula los códigos nuevos de las subtareas de cada tarea de --desde (sin modificar el proyecto)."""
    roots = build_tree(project.tasks)
    scope = select_scope(roots, cfg.plan)
    found: list[Task] = []
    for code in starts:
        t = find_start(scope, code)
        for other in found:
            if t is other or _inside(t, other) or _inside(other, t):
                outer = other if t is other or _inside(t, other) else t
                raise ConfigError("NUM-004", f"--desde {code}: la tarea {ref(t)} está en la rama de {ref(outer)}, "
                                             "que ya se numera. Se debe indicar solo el código de la tarea superior.")
        found.append(t)
    plan = RenumberPlan(scope, found)
    listed = {code_key(strip_prefix(e, scope.prefix)) for e in cfg.wbs_list or []}
    raw_of = scope.raw_code

    def visible_yes(t: Task) -> bool:
        return bool(cfg.visible_field) and parse_visible(t.visible)[0] is True

    def own(t: Task) -> str:
        raw = raw_of(t)
        if not _numeric(raw):
            return ""
        level = code_level(raw, cfg.ignore_interior_zeros)
        if level <= cfg.max_level:
            return f"nivel {level}"
        if visible_yes(t):
            return "Visible = Sí"
        if code_key(raw) in listed:
            return "lista EDT"
        return ""

    protected: dict[int, str] = {}

    def protect(t: Task) -> str:
        inner = [protect(c) for c in t.children]
        reason = own(t) or (PROTECTED_INSIDE if any(inner) else "")
        protected[t.uid] = reason
        return reason

    for start in found:
        for c in start.children:
            protect(c)
        plan.tasks += sum(1 for _ in iter_subtree(start)) - 1

    def issue(id: str, t: Task, message: str) -> None:
        plan.issues.append(Issue(id, (t.wbs or "").strip(), t.name, message, uid=t.uid))

    def number(node: Task, literal: str) -> None:
        """Numera las subtareas de node, cuyo código (conservado o nuevo) es literal."""
        literal = literal.strip()
        raw = strip_prefix(literal, scope.prefix)
        prefix = literal[:len(literal) - len(raw)]
        key = code_key(raw)
        base = pad_code(key, cfg.max_level) if code_level(raw, cfg.ignore_interior_zeros) < cfg.max_level else key
        kids = node.children
        free = next((c for c in kids if not protected[c.uid]), None)
        if free is not None:
            for c in kids[kids.index(free) + 1:]:
                if protected[c.uid]:
                    issue("NUM-008", c, f"Tarea protegida ({protected[c.uid]}) después de {ref(free)}, que se numera: "
                                        f"las subtareas protegidas de {ref(node)} deben ir al principio. Se debe "
                                        "mover la tarea en MS Project.")
        groups: dict[str, list[tuple[Task, str]]] = {}
        valid: list[Task] = []
        for c in kids:
            if not protected[c.uid]:
                continue
            ck = code_key(raw_of(c))
            if not _numeric(raw_of(c)):
                issue("NUM-011", c, f"Tarea sin código EDT numérico que contiene tareas protegidas "
                                    f"({_first_protected(c, protected, own)}): su código no se puede calcular. Se "
                                    "debe poner a mano.")
                continue
            parent_key = ck.rpartition(".")[0]
            if not ck.startswith(key + ".") or parent_key not in (key, base):
                expected = f"{key}.N" if base == key else f"{key}.N o {base}.N"
                issue("NUM-010", c, f"Código protegido ({protected[c.uid]}) que no cuelga directamente de su tarea "
                                    f"superior {ref(node)}: se esperaba {expected}. Se debe corregir el código o la "
                                    "posición de la tarea.")
                continue
            groups.setdefault(parent_key, []).append((c, ck))
            valid.append(c)
        for parent_key, items in groups.items():
            numbers = [int(ck.rpartition(".")[2]) for _c, ck in items]
            if numbers != list(range(1, len(numbers) + 1)):
                c = items[next(i for i, n in enumerate(numbers) if n != i + 1)][0]
                issue("NUM-009", c, f"Las subtareas protegidas de {ref(node)} tienen los códigos "
                                    f"{_codes_text([raw_of(x) for x, _ck in items])}: deben ser correlativos desde "
                                    f"{parent_key}.1, sin saltos y en el orden del esquema. Se deben corregir a mano.")
        n = len(groups.get(base, []))
        for c in kids:
            if protected[c.uid]:
                if c in valid:
                    plan.kept.append(Kept(c, protected[c.uid]))
                    number(c, c.wbs)
                continue
            if not raw_of(c) and visible_yes(c):
                issue("NUM-012", c, "Visible = Sí en una tarea sin código EDT: tendrá fila propia en el libro, así que "
                                    "su código se debe poner a mano.")
                continue
            n += 1
            new_raw = f"{base}.{n}"
            level = code_level(new_raw, cfg.ignore_interior_zeros)
            if level <= cfg.max_level:
                issue("NUM-013", c, f"Su código sería {new_raw}, de nivel {level} (nivel_maximo = {cfg.max_level}, con "
                                    "ignorar_ceros_intermedios = true): sería una fila nueva del libro. La tarea y su "
                                    "rama no se numeran; se debe poner su código a mano.")
                plan.skipped.append(c)
                n -= 1
                continue
            new = prefix + new_raw
            old = (c.wbs or "").strip()
            if old != new:
                plan.changes.append(Change(c, old, new))
            number(c, new)

    for start in found:
        number(start, start.wbs)
    _duplicates(plan)
    return plan


def _inside(t: Task, ancestor: Task) -> bool:
    n = t.parent
    while n is not None:
        if n is ancestor:
            return True
        n = n.parent
    return False


def _first_protected(t: Task, protected: dict[int, str], own) -> str:
    for n in iter_subtree(t):
        if n is not t and own(n):
            return ref(n)
    return ""


def _duplicates(plan: RenumberPlan) -> None:
    """NUM-014: código nuevo que ya tiene una tarea de fuera de las ramas (la rama se actualiza igualmente)."""
    scope = plan.scope
    if not plan.changes:
        return
    in_branches = {t.uid for s in plan.starts for t in iter_subtree(s)} - {s.uid for s in plan.starts}
    outside: dict[str, list[Task]] = {}
    for t in scope.tasks():
        raw = scope.raw_code(t)
        if t.uid not in in_branches and _numeric(raw):
            outside.setdefault(code_key(raw), []).append(t)
    for c in plan.changes:
        others = outside.get(code_key(strip_prefix(c.new, scope.prefix)), [])
        if others:
            plan.issues.append(Issue("NUM-014", c.new, c.task.name,
                                     f"El código nuevo {c.new} ya lo tiene {ref(others[0])}, fuera de las ramas que se "
                                     "actualizan: el código se actualiza igualmente. Se debe revisar el de la otra "
                                     "tarea.", uid=c.task.uid, severity=SEV_WARN))


# ----------------------------------------------------------------------------
# Escritura en MS Project
# ----------------------------------------------------------------------------
@dataclass
class WriteOutcome:
    written: int = 0
    undo_grouped: bool = False
    saved: bool = False
    was_open: bool = False      # el archivo ya estaba abierto en MS Project (o 'activo'): no se guarda
    backup: Path | None = None
    issues: list[Issue] = field(default_factory=list)


def check_field(project: ProjectData, cfg: Config) -> None:
    """El código EDT debe estar en un campo personalizado (no en el campo EDT ni en el número de esquema)."""
    name = project.fields_found.get("edt")
    if name is None:
        raise ConfigError("NUM-005", f"No se encontró el campo del código EDT '{cfg.wbs_field}' ([campos] edt): no se "
                                     "puede escribir. Se debe indicar el campo personalizado de texto con los códigos, "
                                     "por ejemplo \"Item EDT\".")
    if name in NATIVE_FIELDS:
        raise ConfigError("NUM-005", f"El código EDT se lee del campo {name} de MS Project ([campos] edt = "
                                     f"'{cfg.wbs_field}'): actualizar-edt solo escribe en un campo personalizado de "
                                     "texto, como \"Item EDT\". Para el campo EDT de MS Project se usa su función "
                                     "Renumerar.")


def _field_ref(session, cfg: Config) -> tuple[str, object]:
    ref_ = session._resolve_field(cfg.wbs_field)
    if ref_ is None or (ref_[0] == "attr" and str(ref_[1]) in NATIVE_FIELDS):
        raise ConfigError("NUM-005", f"El campo del código EDT '{cfg.wbs_field}' no es un campo personalizado de texto "
                                     "de MS Project: actualizar-edt no puede escribir en él.")
    m = re.fullmatch(r"Text(\d+)", str(ref_[1])) if ref_[0] == "attr" else None
    fid = TASK_TEXT_IDS.get(int(m.group(1))) if m else (int(ref_[1]) if ref_[0] == "fid" else None)
    if fid is not None:
        try:
            formula = str(session.app.CustomFieldGetFormula(fid) or "").strip()
        except Exception:
            formula = ""
        if formula:
            raise ConfigError("NUM-005", f"El campo del código EDT '{cfg.wbs_field}' tiene una fórmula: sus valores los "
                                         "calcula MS Project y no se pueden escribir.")
    return ref_


def _com_tasks(project) -> dict[int, object]:
    out = {}
    for t in project.Tasks:
        if t is None:
            continue
        try:
            if bool(getattr(t, "ExternalTask", False)):
                continue
            out[int(t.UniqueID)] = t
        except Exception:
            continue
    return out


def write_codes(session, plan: RenumberPlan, cfg: Config) -> WriteOutcome:
    """Escribe los códigos nuevos en el proyecto abierto (sin guardarlo) y los vuelve a leer. Los cambios se agrupan
    en una sola acción de deshacer de MS Project si es posible."""
    from evm.client_writer import com_error_text
    out = WriteOutcome()
    kind, key = _field_ref(session, cfg)
    app = session.app
    session._activate()
    tasks = _com_tasks(session.project)

    def put(com, value: str) -> None:
        if kind == "attr":
            setattr(com, key, value)
        else:
            com.SetField(key, value)

    def get(com) -> str:
        value = getattr(com, key) if kind == "attr" else com.GetField(key)
        return "" if value is None else str(value).strip()

    try:
        app.OpenUndoTransaction(UNDO_NAME)
        out.undo_grouped = True
    except Exception as exc:
        log.debug("OpenUndoTransaction no disponible: %s", com_error_text(exc))
        out.issues.append(Issue("NUM-017", message="MS Project no permitió agrupar los cambios en una sola acción de "
                                                   "deshacer: cada código escrito es una acción."))
    screen = None
    try:
        screen = app.ScreenUpdating
        app.ScreenUpdating = False
    except Exception:
        screen = None
    progress.stage("escribir", "Escritura de los códigos EDT")
    current = None
    try:
        for i, change in enumerate(plan.changes, start=1):
            progress.step(i, len(plan.changes), "tareas")
            current = change
            com = tasks.get(change.task.uid)
            if com is None:
                raise ProjectAppError("NUM-015", f"No se encontró en MS Project la tarea {ref(change.task)} (UID "
                                                 f"{change.task.uid}).")
            put(com, change.new)
            out.written += 1
    except ProjectAppError:
        raise
    except Exception as exc:
        log.debug("Error al escribir el código de UID %s: %s", current.task.uid if current else "?",
                  com_error_text(exc))
        raise ProjectAppError("NUM-015", f"MS Project devolvió un error al escribir el código "
                                         f"{current.new if current else ''} ({com_error_text(exc)}). Códigos "
                                         f"escritos: {out.written}.") from None
    finally:
        if out.undo_grouped:
            try:
                app.CloseUndoTransaction()
            except Exception as exc:
                log.debug("CloseUndoTransaction: %s", com_error_text(exc))
        if screen is not None:
            try:
                app.ScreenUpdating = screen
            except Exception:
                pass
    progress.stage("comprobar", "Comprobación de los códigos escritos")
    wrong = []
    for change in plan.changes:
        try:
            got = get(tasks[change.task.uid])
        except Exception as exc:
            got = f"(error: {com_error_text(exc)})"
        if got != change.new:
            wrong.append(f"{change.new} (se leyó '{got}')")
    if wrong:
        raise ProjectAppError("NUM-016", f"{len(wrong)} códigos no coinciden con los previstos al volver a leerlos: "
                                         f"{_codes_text(wrong, 5)}.")
    return out


def backup_path(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = path.with_name(f"{path.stem}.respaldo-{stamp}{path.suffix}")
    n = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}.respaldo-{stamp}-{n}{path.suffix}")
        n += 1
    return candidate


# ----------------------------------------------------------------------------
# Informe
# ----------------------------------------------------------------------------
def _cell(text: str) -> str:
    return (text or "").replace("|", "\\|").replace("\n", " ").strip()


def write_report(path: Path, plan: RenumberPlan, cfg: Config, *, origen: str, version: str, result: str,
                 notices: list[Issue], outcome: WriteOutcome | None) -> None:
    """Informe local de actualizar-edt (con nombres de tareas: uso interno)."""
    scope = plan.scope
    lines = [f"# Actualización de códigos EDT: {Path(origen).name if origen.lower() != 'activo' else 'proyecto activo'}",
             "",
             f"- Fecha: {datetime.now():%Y-%m-%d %H:%M}",
             f"- Versión del programa: {version}",
             f"- Origen: {origen}"]
    if scope.prefix:
        lines.append(f"- Planificación: {scope.prefix}")
    lines += [f"- Desde: {', '.join(ref(s) for s in plan.starts)}",
              f"- Protegidas: nivel 1 a {cfg.max_level} (nivel_maximo)"
              + (f", Visible = Sí (campo «{cfg.visible_field}»)" if cfg.visible_field else "")
              + (", códigos de la lista EDT" if cfg.wbs_list is not None else "")
              + " y las tareas que las contienen",
              f"- Subtareas en las ramas: {plan.tasks}; códigos que cambian: {len(plan.changes)} (de ellos, "
              f"{plan.added} nuevos); protegidas: {len(plan.kept)}",
              f"- Resultado: {result}"]
    if outcome is not None and outcome.backup is not None:
        lines.append(f"- Copia de respaldo: {outcome.backup}")
    lines.append("")
    if notices:
        lines += ["## Observaciones", "", "| Código | Severidad | Tarea | Mensaje |", "|---|---|---|---|"]
        for i in notices:
            task = f"{i.code} {i.name}".strip()
            lines.append(f"| {i.id} | {i.severity} | {_cell(task)} | {_cell(i.message)} |")
        lines.append("")
    lines += ["## Cambios", ""]
    if plan.changes:
        lines += ["| Id | Código actual | Código nuevo | Tarea | Nota |", "|---:|---|---|---|---|"]
        for c in plan.changes:
            note = "inactiva" if not c.task.active else ""
            lines.append(f"| {c.task.id} | {_cell(c.old) or '(sin código)'} | {c.new} | {_cell(c.task.name)} | "
                         f"{note} |")
    else:
        lines.append("Ningún código cambia.")
    lines += ["", "## Tareas protegidas", ""]
    if plan.kept:
        lines += ["| Id | Código | Tarea | Motivo |", "|---:|---|---|---|"]
        for k in plan.kept:
            lines.append(f"| {k.task.id} | {_cell(k.task.wbs)} | {_cell(k.task.name)} | {k.reason} |")
    else:
        lines.append("Ninguna.")
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")
