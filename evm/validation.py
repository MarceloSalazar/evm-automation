"""Validaciones del archivo de MS Project.

- Datos que forman el libro: se informan en la hoja Validación de 'generar' y en el comando 'validar'. Son los
  códigos EDT (dependencia entre tareas superiores y subtareas, formato, orden y duplicados), las fechas y la
  duración, los pesos de los hitos, el método de medición y el reparto entre empresas.
- Cronograma: solo en 'validar'. Son el avance a la fecha de estado y la lógica del cronograma (vínculos y
  delimitaciones, según los criterios de la evaluación DCMA de 14 puntos).
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import replace
from datetime import date, datetime

from evm.config import Config, ConfigError
from evm.fields import fmt_number, same_name
from evm.methods import FIXED, MILESTONES, UNKNOWN, is_complete, is_started, parse_method
from evm.model import (CONSTRAINT_FNLT, CONSTRAINT_MFO, CONSTRAINT_MSO, CONSTRAINT_SNLT, LINK_FS, LINK_SF,
                       ProjectData, Task)
from evm.wbs import (KIND_RESIDUAL, SEV_ERROR, SEV_INFO, SEV_WARN, TOPIC_BOOK, TOPIC_COMPANY, TOPIC_DATES,
                     TOPIC_LOGIC, TOPIC_METHOD, TOPIC_PROGRESS, TOPIC_READ, TOPIC_WBS, TOPIC_WEIGHTS, Issue,
                     PlanNotFound, Report, ReportPart, ReportRow, Scope, assemble_report, build_part, iter_subtree,
                     pad_code, plural, tagged_issues, task_name)

CodeOf = Callable[[Task], str]

LIMIT = 200                     # observaciones por tarea de cada control (como en el resto de la hoja Validación)
MONTH_DAYS = 365.25 / 12

# Criterios de la evaluación DCMA de 14 puntos
DCMA_LOGIC = 5.0                # % máximo de tareas sin predecesoras o sin sucesoras
DCMA_LAGS = 5.0                 # % máximo de vínculos con posición positiva
DCMA_FS = 90.0                  # % mínimo de vínculos fin a comienzo
DCMA_CONSTRAINTS = 5.0          # % máximo de tareas con delimitaciones fijas
DCMA_BEI = 0.95                 # índice de ejecución de la línea base mínimo

HARD_CONSTRAINTS = {
    CONSTRAINT_MSO: "Debe comenzar el",
    CONSTRAINT_MFO: "Debe finalizar el",
    CONSTRAINT_SNLT: "No comenzar después del",
    CONSTRAINT_FNLT: "No finalizar después del",
}


def _num(v: float, decimals: int = 1) -> str:
    """Número con coma decimal y sin ceros finales: 2,5; 3."""
    return fmt_number(v, decimals, trim=True)


def _pct(part: int, total: int) -> str:
    return _num(100.0 * part / total) if total else "0"


def _day(d: datetime | date | None) -> str:
    return f"{d:%d/%m/%Y}" if d else "-"


def _as_date(d: datetime | date | None) -> date | None:
    if isinstance(d, datetime):
        return d.date()
    return d


def _codes(tasks: list[Task], code_of: CodeOf, limit: int = 10) -> str:
    """Lista breve de tareas: '1.2.1, 1.2.3 y 4 más' (las tareas sin código, por su Id)."""
    names = [code_of(t) or f"Id {t.id}" for t in tasks[:limit]]
    more = f" y {len(tasks) - limit} más" if len(tasks) > limit else ""
    return ", ".join(names) + more


def _label(t: Task, code_of: CodeOf) -> str:
    return f"{code_of(t) or f'Id {t.id}'} «{t.name}»"


def _issue(severity: str, t: Task, code_of: CodeOf, message: str, topic: str) -> Issue:
    return Issue(severity, code_of(t), task_name(t, code_of), message, topic)


# ----------------------------------------------------------------------------
# Códigos EDT
# ----------------------------------------------------------------------------
def _segments(code: str) -> list[str] | None:
    """Segmentos de un código EDT numérico ('1.2.0' -> ['1', '2', '0']); None si tiene otros caracteres."""
    parts = [p.strip() for p in code.split(".")]
    return parts if all(p.isdigit() for p in parts) else None


def _significant(parts: list[str]) -> tuple[int, ...]:
    """Segmentos sin los ceros de relleno del final: '1.2.0' -> (1, 2); '1.0.0' -> (1,)."""
    nums = [int(p) for p in parts]
    while len(nums) > 1 and nums[-1] == 0:
        nums.pop()
    return tuple(nums)


def _pattern(prefix: tuple[int, ...]) -> str:
    return ".".join(str(n) for n in prefix) + ".x"


def _depth(sig: tuple[int, ...]) -> int:
    """Nivel de un código: sus segmentos significativos sin los ceros intermedios ((1, 2, 0, 1) -> 3)."""
    return 1 + sum(1 for n in sig[1:] if n)


def _upper(sig: tuple[int, ...]) -> tuple[int, ...]:
    """Código del elemento del que depende un código: (1, 2, 3) -> (1, 2); (1, 2, 0, 1) -> (1, 2)."""
    up = list(sig[:-1])
    while len(up) > 1 and up[-1] == 0:
        up.pop()
    return tuple(up)


def wbs_checks(scope: Scope, code_of: CodeOf, cfg: Config) -> list[Issue]:
    """Códigos EDT de las tareas exportadas: formato, dependencia de su tarea superior, orden y duplicados.

    El código de cada tarea debe empezar por el de su tarea superior con código EDT y tener un nivel más: bajo
    1.2.0 van 1.2.1, 1.2.2... (1.14.2 es un error). Un 0 intermedio indica un hijo directo del elemento anterior:
    1.2.0.1 depende de 1.2, como 1.2.1.0 (sirve para insertar un elemento sin renumerar a sus hermanos). Los
    códigos no numéricos sin puntos (p. ej. INTERNO) y el código 0 (el proyecto) se admiten como contenedores: sus
    subtareas no se comparan con ellos.
    """
    issues: list[Issue] = []
    tasks = list(scope.tasks())
    in_scope = {t.uid for t in tasks}
    sig: dict[int, tuple[int, ...] | None] = {}        # UniqueID -> segmentos significativos (None = no numérico)
    for t in tasks:
        raw = scope.raw_code(t)
        if raw:
            parts = _segments(raw)
            sig[t.uid] = _significant(parts) if parts else None

    def coded_parent(t: Task) -> Task | None:
        p = t.parent
        while p is not None and p.uid in in_scope:
            if p.uid in sig:
                return p
            p = p.parent
        return None

    bad_format: list[Issue] = []
    groups: dict[int | None, list[tuple[Task, tuple[int, ...]]]] = {}
    for t in tasks:
        if t.uid not in sig:
            continue
        raw = scope.raw_code(t)
        parent = coded_parent(t)
        mine = sig[t.uid]
        if mine is None:
            if any(not p.strip() for p in raw.split(".")):
                bad_format.append(_issue(SEV_ERROR, t, code_of,
                                         f"Código EDT '{raw}' con formato no válido: deben ser números separados "
                                         "por un punto, sin segmentos vacíos (p. ej. 1.2.3).", TOPIC_WBS))
            elif "." in raw or (parent is not None and sig.get(parent.uid) is not None):
                bad_format.append(_issue(SEV_WARN, t, code_of,
                                         f"Código EDT '{raw}' no numérico: se esperaban números separados por un "
                                         "punto (p. ej. 1.2.3); no se puede comprobar con su tarea superior.",
                                         TOPIC_WBS))
            continue
        top = sig.get(parent.uid) if parent is not None else None
        if top == (0,):                                 # 0 (o 0.0.0) = el proyecto: contenedor, como INTERNO
            top = None
        if top is None:                                 # sin tarea superior con código numérico
            if _depth(mine) > 1:
                upper = ".".join(str(n) for n in _upper(mine))
                upper = pad_code(upper, cfg.max_level) if cfg.pad_wbs else upper
                issues.append(_issue(SEV_WARN, t, code_of,
                                     f"Código de nivel {_depth(mine)} sin una tarea superior con código EDT: debería "
                                     f"depender de la tarea {upper}.", TOPIC_WBS))
            groups.setdefault(parent.uid if parent is not None else None, []).append((t, mine))
            continue
        where = f"su tarea superior {_label(parent, code_of)}"
        if mine == top:
            issues.append(_issue(SEV_ERROR, t, code_of, f"Tiene el mismo código EDT que {where}.", TOPIC_WBS))
        elif mine[:len(top)] != top or len(mine) <= len(top):
            issues.append(_issue(SEV_ERROR, t, code_of,
                                 f"El código EDT no depende del de {where}: se esperaba {_pattern(top)}.",
                                 TOPIC_WBS))
        elif _depth(mine) > _depth(top) + 1:
            issues.append(_issue(SEV_WARN, t, code_of,
                                 f"El código EDT salta un nivel: depende de {where}, de nivel {_depth(top)}, y su "
                                 f"código es de nivel {_depth(mine)} (se esperaba {_pattern(top)}). Si hay una "
                                 "tarea intermedia sin código EDT, se le debe asignar uno.", TOPIC_WBS))
        else:
            groups.setdefault(parent.uid, []).append((t, mine))

    # Orden y numeración de las subtareas de cada tarea (en el orden de MS Project)
    for members in groups.values():
        jumps: list[str] = []
        for (prev, a), (t, b) in zip(members, members[1:]):
            if b < a:
                issues.append(_issue(SEV_WARN, t, code_of,
                                     f"Código EDT fuera de orden: está después de {_label(prev, code_of)}.",
                                     TOPIC_WBS))
            elif len(a) == len(b) and b[:-1] == a[:-1] and b[-1] > a[-1] + 1:
                jumps.append(f"{code_of(prev)} → {code_of(t)}")
        if jumps:
            first = members[0][0]
            parent = coded_parent(first)
            under = f"Bajo {_label(parent, code_of)}" if parent is not None else "En el nivel superior"
            shown = ", ".join(jumps[:5]) + (f" y {len(jumps) - 5} más" if len(jumps) > 5 else "")
            issues.append(Issue(SEV_INFO, code_of(parent) if parent is not None else "",
                                parent.name if parent is not None else "",
                                f"{under}, la numeración tiene saltos ({shown}): si no es intencional (p. ej. un "
                                "paquete eliminado), se debe revisar el código.", TOPIC_WBS))

    # Códigos duplicados (como quedan en el libro: con el relleno de ceros)
    by_code: dict[str, list[Task]] = {}
    for t in tasks:
        if t.uid in sig:
            by_code.setdefault(code_of(t), []).append(t)
    hint = "" if scope.plan is not None else (" Si el proyecto tiene varias planificaciones, se debe elegir una con "
                                              "'planificacion'.")
    for code, same in by_code.items():
        if len(same) > 1:
            names = ", ".join(f"«{t.name}» (Id {t.id})" for t in same[:5])
            more = f" y {len(same) - 5} más" if len(same) > 5 else ""
            issues.append(Issue(SEV_ERROR, code, same[0].name,
                                f"Código EDT duplicado: lo tienen {len(same)} tareas ({names}{more}).{hint}",
                                TOPIC_WBS))
    return bad_format[:LIMIT] + issues


# ----------------------------------------------------------------------------
# Datos de las tareas: fechas y duración, pesos, método y empresas
# ----------------------------------------------------------------------------
def _zero_duration(t: Task) -> bool:
    """Duración 0 según MS Project (sin el dato de la duración, no se puede saber)."""
    return t.duration is not None and abs(t.duration) < 1e-9


def _positive_duration(t: Task) -> bool:
    return t.duration is not None and t.duration > 1e-9


def _row_months(row: ReportRow) -> float | None:
    if row.start is None or row.finish is None:
        return None
    return max(0.0, (row.finish - row.start).total_seconds() / 86400 / MONTH_DAYS)


def data_checks(project: ProjectData, scope: Scope, rows: list[ReportRow], cfg: Config,
                code_of: CodeOf) -> list[Issue]:
    """Controles de los datos de las tareas que forman el libro (scope sin las tareas inactivas ni las de
    Visible = No)."""
    issues: list[Issue] = []
    tasks = list(scope.tasks())
    leaves = [t for t in tasks if not t.children]

    # --- Fechas y duración ----------------------------------------------------
    dates: list[Issue] = []
    for t in leaves:
        if (project.baseline_found and t.baseline_start and t.baseline_finish
                and t.baseline_finish < t.baseline_start):
            dates.append(_issue(SEV_ERROR, t, code_of,
                                f"Fin previsto ({_day(t.baseline_finish)}) anterior al Comienzo previsto "
                                f"({_day(t.baseline_start)}) en la línea base.", TOPIC_DATES))
        if t.start and t.finish and t.finish < t.start:
            dates.append(_issue(SEV_ERROR, t, code_of,
                                f"Fin ({_day(t.finish)}) anterior al Comienzo ({_day(t.start)}).", TOPIC_DATES))
        elif (t.start is None or t.finish is None) and (t.baseline_start or t.baseline_finish):
            dates.append(_issue(SEV_WARN, t, code_of,
                                "Sin Comienzo o sin Fin en el programa actual (p. ej. una tarea programada "
                                "manualmente sin fechas).", TOPIC_DATES))
        if t.milestone and _positive_duration(t):
            dates.append(_issue(SEV_WARN, t, code_of,
                                "Está marcada como hito y tiene duración: su % planificado va entero al mes de su "
                                "fecha de fin. Si no es un hito, se debe quitar la marca de hito en MS Project "
                                "(Información de la tarea > Avanzado).", TOPIC_DATES))
        elif not t.milestone and _zero_duration(t):
            dates.append(_issue(SEV_WARN, t, code_of,
                                "Duración 0 y no está marcada como hito: no pesa en el % planificado de su fila "
                                "(pesan las tareas con duración y los hitos con peso). Si es un hito, se debe "
                                "marcar como hito.", TOPIC_DATES))
    issues += dates[:LIMIT]

    estimated = [t for t in leaves if t.estimated]
    if estimated:
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(estimated), 'tarea tiene', 'tareas tienen')} duración estimada (con «?», "
                            "p. ej. «5 días?»): la duración define el % planificado. Se debe confirmar o corregir.",
                            TOPIC_DATES))
        issues += [_issue(SEV_WARN, t, code_of, "Duración estimada («?»).", TOPIC_DATES) for t in estimated[:LIMIT]]
    manual = [t for t in leaves if t.manual]
    if manual:
        issues.append(Issue(SEV_INFO, "", "",
                            f"{plural(len(manual), 'tarea programada', 'tareas programadas')} manualmente "
                            f"({_codes(manual, code_of)}): MS Project no recalcula sus fechas cuando cambian sus "
                            "predecesoras o el avance.", TOPIC_DATES))

    # --- Hitos y pesos --------------------------------------------------------
    if cfg.milestone_weight_field:
        unused = [t for t in tasks if t.milestone_weight and (not t.milestone or t.children)]
        for t in unused[:LIMIT]:
            what = "es una tarea resumen" if t.children else "no es un hito"
            issues.append(_issue(SEV_WARN, t, code_of,
                                 f"Tiene '{cfg.milestone_weight_field}' = {_num(t.milestone_weight, 2)}, pero {what}: "
                                 "el peso no se usa (solo cuentan los hitos de detalle; las tareas pesan por su "
                                 "duración).", TOPIC_WEIGHTS))
        milestones = [t for t in leaves if t.milestone]
        if ("peso_hito" in project.fields_found and milestones
                and not any((t.milestone_weight or 0) > 0 for t in milestones)):
            issues.append(Issue(SEV_WARN, "", "",
                                f"El campo '{project.fields_found['peso_hito']}' existe, pero ningún hito tiene peso: "
                                "los hitos no tienen % planificado propio (con «Hitos ponderados», pesan igual). Se "
                                "debe revisar el campo indicado en [campos] peso_hito.", TOPIC_READ))
    for row in rows:
        if not row.is_input or parse_method(row.method).kind != MILESTONES:
            continue
        milestones = [t for t in row.leaves if t.milestone]
        weighted = [t for t in milestones if (t.milestone_weight or 0) > 0]
        if not milestones:
            issues.append(Issue(SEV_WARN, row.code, row.name,
                                "Método «Hitos ponderados» y la fila no tiene hitos: «actualizar» calculará el % "
                                "ganado con el % físico completado (o el % completado).", TOPIC_WEIGHTS))
        elif not weighted:
            issues.append(Issue(SEV_INFO, row.code, row.name,
                                "Método «Hitos ponderados» con hitos sin peso: los hitos se ponderan por igual.",
                                TOPIC_WEIGHTS))
        elif len(weighted) < len(milestones):
            loose = [t for t in milestones if t not in weighted]
            issues.append(Issue(SEV_INFO, row.code, row.name,
                                f"Método «Hitos ponderados»: {plural(len(loose), 'hito', 'hitos')} sin peso no "
                                f"{'suma' if len(loose) == 1 else 'suman'} % ganado al cumplirse "
                                f"({', '.join('«' + t.name + '»' for t in loose[:5])}).", TOPIC_WEIGHTS))

    # --- Método de medición ---------------------------------------------------
    method_names = "0/100, 50/50 (X/Y que sumen 100), Hitos ponderados, % Completado, % Físico completado, " \
                   "Unidades completadas o Avance lineal (LOE)"
    if "metodo" in project.fields_found and tasks and not any(t.method for t in tasks):
        issues.append(Issue(SEV_WARN, "", "",
                            f"El campo '{project.fields_found['metodo']}' existe, pero está vacío en todas las "
                            f"tareas: todas las filas tienen el método por defecto ('{cfg.default_method}').",
                            TOPIC_READ))
    for row in rows:
        if not row.is_input:
            continue
        method = parse_method(row.method)
        if method.kind == UNKNOWN:
            issues.append(Issue(SEV_WARN, row.code, row.name,
                                f"Método '{row.method}' no reconocido: «actualizar» calculará el % ganado con el % "
                                f"completado. Se admite {method_names}.", TOPIC_METHOD))
        if cfg.method_field and row.kind != KIND_RESIDUAL:
            own = (row.task.method or "").strip()
            other = [n for n in iter_subtree(row.task) if n is not row.task and (n.method or "").strip()
                     and (n.method or "").strip() != own]
            if other:
                values = ", ".join(dict.fromkeys(f"'{n.method.strip()}'" for n in other))
                subtasks = plural(len(other), "subtarea tiene", "subtareas tienen")
                if own:
                    issues.append(Issue(SEV_INFO, row.code, row.name,
                                        f"{subtasks} otro método en MS Project ({values}): se usa el de la fila, "
                                        f"'{own}'.", TOPIC_METHOD))
                else:
                    issues.append(Issue(SEV_WARN, row.code, row.name,
                                        f"La fila no tiene método y {subtasks} método en MS Project ({values}): se "
                                        f"usa el método por defecto ('{row.method}'). El método se indica en la tarea "
                                        "de la fila.", TOPIC_METHOD))
        if method.kind == FIXED:
            months = _row_months(row)
            zero = method.start_fraction <= 0
            limit = cfg.max_months_0_100 if zero else cfg.max_months_fixed
            if months is not None and limit > 0 and months > limit + 1e-9:
                key = "meses_0_100" if zero else "meses_formula_fija"
                effect = ("el EV queda en 0 hasta que termina" if zero else
                          f"el EV queda en el {_num(method.start_fraction * 100)} % desde que comienza hasta que "
                          "termina")
                issues.append(Issue(SEV_WARN, row.code, row.name,
                                    f"Método {row.method} en un paquete de {_num(months)} meses (línea base), más "
                                    f"de {_num(limit)}: {effect}. Se recomienda dividir el paquete o usar un método "
                                    f"de avance (% completado, hitos ponderados). Umbral: [validacion] {key}.",
                                    TOPIC_METHOD))

    # --- Empresas -------------------------------------------------------------
    if cfg.multi_enabled:
        if "campo_empresa" in project.fields_found and tasks and not any(t.company for t in tasks):
            name = project.fields_found["campo_empresa"]
            issues.append(Issue(SEV_WARN, "", "",
                                f"El campo '{name}' existe, pero está vacío en todas las tareas: se debe revisar "
                                "[multiempresa] campo_empresa.", TOPIC_READ))
        ignored = [t for t in tasks if t.split_a and not same_name(t.company or "", cfg.company_both)]
        for t in ignored[:LIMIT]:
            who = f"su empresa es '{t.company}'" if t.company else "no tiene empresa propia (la hereda)"
            issues.append(_issue(SEV_INFO, t, code_of,
                                 f"Tiene '{cfg.split_field}' = {_num(t.split_a, 2)}, pero {who}: el reparto solo se "
                                 f"usa en las tareas '{cfg.company_both}'.", TOPIC_COMPANY))
    return issues


# ----------------------------------------------------------------------------
# Avance a la fecha de estado (solo 'validar')
# ----------------------------------------------------------------------------
def progress_checks(project: ProjectData, scope: Scope, code_of: CodeOf, status: date | None) -> list[Issue]:
    """Fechas reales coherentes entre sí y con el % completado, y avance actualizado a la fecha de estado."""
    issues: list[Issue] = []
    leaves = [t for t in scope.tasks() if not t.children]
    found: list[Issue] = []
    for t in leaves:
        if t.actual_finish and not t.actual_start:
            found.append(_issue(SEV_ERROR, t, code_of, "Tiene Fin real y no tiene Comienzo real.", TOPIC_PROGRESS))
        elif t.actual_finish and t.actual_start and t.actual_finish < t.actual_start:
            found.append(_issue(SEV_ERROR, t, code_of,
                                f"Fin real ({_day(t.actual_finish)}) anterior al Comienzo real "
                                f"({_day(t.actual_start)}).", TOPIC_PROGRESS))
        if t.percent_complete > 0 and not t.actual_start:
            found.append(_issue(SEV_WARN, t, code_of,
                                f"% completado = {_num(t.percent_complete)} % y no tiene Comienzo real.",
                                TOPIC_PROGRESS))
        if t.percent_complete >= 100 and not t.actual_finish:
            found.append(_issue(SEV_WARN, t, code_of, "100 % completado y no tiene Fin real.", TOPIC_PROGRESS))
        elif t.actual_finish and t.percent_complete < 99.5:
            found.append(_issue(SEV_WARN, t, code_of,
                                f"Tiene Fin real y su % completado es {_num(t.percent_complete)} %: para el programa "
                                "no está terminada (se considera terminada con el 100 %).", TOPIC_PROGRESS))
    has_progress = any(is_started(t) for t in leaves)

    if status is None:
        issues += found[:LIMIT]
        if has_progress:
            issues.append(Issue(SEV_WARN, "", "",
                                "El proyecto tiene avance cargado y no tiene fecha de estado: no se revisó el avance "
                                "a la fecha de estado. «actualizar» necesita la fecha de estado de MS Project "
                                "(Proyecto > Información del proyecto) o --fecha-corte.", TOPIC_PROGRESS))
        else:
            issues.append(Issue(SEV_INFO, "", "",
                                "El proyecto no tiene fecha de estado ni avance cargado: no se revisó el avance (la "
                                "fecha se puede indicar con --fecha-corte).", TOPIC_PROGRESS))
        return issues

    for t in leaves:
        for label, when in (("Comienzo real", t.actual_start), ("Fin real", t.actual_finish)):
            if when and when.date() > status:
                found.append(_issue(SEV_ERROR, t, code_of,
                                     f"{label} ({_day(when)}) posterior a la fecha de estado ({_day(status)}): avance "
                                     "registrado en el futuro.", TOPIC_PROGRESS))
    pending: list[Task] = []
    for t in leaves:
        if is_complete(t):
            continue
        if t.finish and t.finish.date() < status:
            pending.append(t)
            found.append(_issue(SEV_WARN, t, code_of,
                                f"Fin ({_day(t.finish)}) anterior a la fecha de estado y no está terminada "
                                f"({_num(t.percent_complete)} %).", TOPIC_PROGRESS))
        elif not is_started(t) and t.start and t.start.date() < status:
            pending.append(t)
            found.append(_issue(SEV_WARN, t, code_of,
                                f"Comienzo ({_day(t.start)}) anterior a la fecha de estado y no tiene avance ni "
                                "Comienzo real.", TOPIC_PROGRESS))
    if pending:
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(pending), 'tarea tiene', 'tareas tienen')} trabajo pendiente programado "
                            f"antes de la fecha de estado ({_day(status)}): el cronograma no está actualizado. Se "
                            "debe registrar el avance o reprogramar el trabajo pendiente después de la fecha de "
                            "estado (Proyecto > Actualizar proyecto).", TOPIC_PROGRESS))
    issues += found[:LIMIT]

    if project.baseline_found:
        late = [t for t in leaves if t.baseline_start and t.baseline_start.date() <= status and not is_started(t)]
        if late:
            issues.append(Issue(SEV_INFO, "", "",
                                f"{plural(len(late), 'tarea debía', 'tareas debían')} comenzar a la fecha de estado "
                                f"según la línea base y no {'comenzó' if len(late) == 1 else 'comenzaron'} "
                                f"({_codes(late, code_of)}).", TOPIC_PROGRESS))
        due = [t for t in leaves if t.baseline_finish and t.baseline_finish.date() <= status]
        if due:
            done = sum(1 for t in leaves if is_complete(t))
            missed = [t for t in due if not is_complete(t)]
            bei = done / len(due)
            detail = f" Sin terminar: {_codes(missed, code_of)}." if missed else ""
            issues.append(Issue(SEV_WARN if bei < DCMA_BEI - 1e-9 else SEV_INFO, "", "",
                                f"Índice de ejecución de la línea base (BEI) = {_num(bei, 2)}: {done} tareas "
                                f"terminadas y {len(due)} que debían terminar a la fecha de estado (DCMA: al menos "
                                f"{_num(DCMA_BEI, 2)}).{detail}", TOPIC_PROGRESS))
    return issues


# ----------------------------------------------------------------------------
# Lógica del cronograma (solo 'validar')
# ----------------------------------------------------------------------------
def logic_checks(project: ProjectData, scope: Scope, code_of: CodeOf) -> list[Issue]:
    """Vínculos y delimitaciones de las tareas exportadas, según los criterios de la evaluación DCMA."""
    issues: list[Issue] = []
    tasks = list(scope.tasks())
    detail = [t for t in tasks if not t.children]
    open_tasks = [t for t in detail if not is_complete(t)]
    by_uid = {t.uid: t for t in project.tasks}
    successors: dict[int, list[Task]] = {}
    links = 0
    for t in project.tasks:
        for link in t.predecessors:
            successors.setdefault(link.uid, []).append(t)
            links += 1

    def pred_label(uid: int) -> str:
        p = by_uid.get(uid)
        return _label(p, code_of) if p is not None else f"una tarea de otro proyecto (UID {uid})"

    if not links:
        if detail:
            issues.append(Issue(SEV_WARN, "", "",
                                "El cronograma no tiene vínculos entre tareas: las fechas no se recalculan cuando "
                                "cambian las duraciones o el avance. Se recomienda vincular las tareas (cada tarea "
                                "con al menos una predecesora y una sucesora).", TOPIC_LOGIC))
    else:
        def linked(t: Task, side: str) -> bool:
            n: Task | None = t
            while n is not None:                    # un vínculo de la tarea resumen alcanza a sus subtareas
                if (n.predecessors if side == "pred" else successors.get(n.uid)):
                    return True
                n = n.parent
            return False

        starts = [t.start for t in detail if t.start]
        finishes = [t.finish for t in detail if t.finish]
        first = min(starts) if starts else None     # comienzo y fin del proyecto: no necesitan predecesora o sucesora
        last = max(finishes) if finishes else None
        no_pred = [t for t in open_tasks if not linked(t, "pred") and t.start != first]
        no_succ = [t for t in open_tasks if not linked(t, "succ") and t.finish != last]
        if no_pred or no_succ:
            share = _pct(len({t.uid for t in no_pred + no_succ}), len(open_tasks))
            issues.append(Issue(SEV_WARN, "", "",
                                f"{plural(len(no_pred), 'tarea', 'tareas')} sin predecesoras y "
                                f"{plural(len(no_succ), 'tarea', 'tareas')} sin sucesoras: {share} % de las tareas sin "
                                f"terminar (DCMA: hasta {_num(DCMA_LOGIC)} %). Sus fechas no dependen de la lógica "
                                "del cronograma.", TOPIC_LOGIC))
            missing_pred = {t.uid for t in no_pred}
            missing_succ = {t.uid for t in no_succ}
            found = []
            for t in open_tasks:
                if t.uid in missing_pred and t.uid in missing_succ:
                    found.append(_issue(SEV_WARN, t, code_of, "Sin predecesoras ni sucesoras.", TOPIC_LOGIC))
                elif t.uid in missing_pred:
                    found.append(_issue(SEV_WARN, t, code_of, "Sin predecesoras.", TOPIC_LOGIC))
                elif t.uid in missing_succ:
                    found.append(_issue(SEV_WARN, t, code_of, "Sin sucesoras.", TOPIC_LOGIC))
            issues += found[:LIMIT]

        own = [(t, link) for t in tasks for link in t.predecessors]
        found = []
        for t, link in own:
            if link.lag < 0:
                found.append(_issue(SEV_WARN, t, code_of,
                                    f"Vínculo con adelanto (posición negativa) desde {pred_label(link.uid)}: se "
                                    "recomienda no usar adelantos (DCMA).", TOPIC_LOGIC))
            if link.kind == LINK_SF:
                found.append(_issue(SEV_WARN, t, code_of,
                                    f"Vínculo comienzo a fin (CF) desde {pred_label(link.uid)}: poco habitual; se "
                                    "recomienda un vínculo fin a comienzo (FC).", TOPIC_LOGIC))
        issues += found[:LIMIT]
        if own:
            lags = [t for t, link in own if link.lag > 0]
            if lags:
                share = 100.0 * len(lags) / len(own)
                issues.append(Issue(SEV_WARN if share > DCMA_LAGS else SEV_INFO, "", "",
                                    f"{plural(len(lags), 'vínculo tiene', 'vínculos tienen')} posición positiva "
                                    f"({_num(share)} % de los vínculos; DCMA: hasta {_num(DCMA_LAGS)} %): "
                                    f"{_codes(lags, code_of)}. Una espera fija oculta trabajo que conviene "
                                    "representar como tarea.", TOPIC_LOGIC))
            fs = sum(1 for _, link in own if link.kind == LINK_FS)
            share = 100.0 * fs / len(own)
            if share < DCMA_FS - 1e-9:
                issues.append(Issue(SEV_WARN, "", "",
                                    f"El {_num(share)} % de los vínculos son fin a comienzo (FC) (DCMA: al menos "
                                    f"{_num(DCMA_FS)} %).", TOPIC_LOGIC))
        for t in tasks:
            if t.children and (t.predecessors or successors.get(t.uid)):
                issues.append(_issue(SEV_WARN, t, code_of,
                                     "Tarea resumen con vínculos: se recomienda vincular sus tareas de detalle.",
                                     TOPIC_LOGIC))

    hard = [t for t in tasks if t.constraint_type in HARD_CONSTRAINTS and not is_complete(t)]
    if hard:
        share = _pct(len(hard), len(open_tasks) or len(tasks))
        issues.append(Issue(SEV_WARN, "", "",
                            f"{plural(len(hard), 'tarea tiene', 'tareas tienen')} delimitaciones que fijan sus "
                            f"fechas ({share} % de las tareas sin terminar; DCMA: hasta {_num(DCMA_CONSTRAINTS)} %): "
                            "la lógica del cronograma no las puede mover. Se recomienda «Lo antes posible» y, si "
                            "hace falta, una fecha límite.", TOPIC_LOGIC))
        for t in hard[:LIMIT]:
            when = f" {_day(t.constraint_date)}" if t.constraint_date else ""
            issues.append(_issue(SEV_WARN, t, code_of,
                                 f"Delimitación «{HARD_CONSTRAINTS[t.constraint_type]}{when}».", TOPIC_LOGIC))
    return issues


# ----------------------------------------------------------------------------
# Comando validar
# ----------------------------------------------------------------------------
def validate_part(project: ProjectData, cfg: Config, name: str = "", status: date | None = None) -> ReportPart:
    """Filas y observaciones de un archivo, con los controles del cronograma (avance y lógica).

    Si la planificación indicada no existe, se informa como error y se revisa el proyecto completo.
    """
    first: list[Issue] = []
    try:
        part = build_part(project, cfg, name)
    except PlanNotFound as exc:
        first.append(Issue(SEV_ERROR, "", "", f"{exc} Se revisa el proyecto completo.", TOPIC_WBS))
        part = build_part(project, replace(cfg, plan=""), name)
    when = status or _as_date(project.status_date)
    part.issues[:0] = first
    part.checks += progress_checks(project, part.scope, part.code_of, when)
    part.checks += logic_checks(project, part.scope, part.code_of)
    return part


def combine(parts: list[ReportPart], cfg: Config) -> tuple[Report | None, list[Issue]]:
    """Informe combinado y sus observaciones (sin los recordatorios del libro). Si los archivos no se pueden
    combinar (códigos EDT repetidos entre archivos), el error y las observaciones de cada archivo."""
    try:
        report = assemble_report(parts, cfg)
        issues = report.issues
    except ConfigError as exc:
        report = None
        issues = [Issue(SEV_ERROR, "", "", str(exc), TOPIC_WBS)]
        for part in parts:
            issues += tagged_issues(part, part.issues + part.checks)
    return report, [i for i in issues if not (i.topic == TOPIC_BOOK and i.severity == SEV_INFO)]
