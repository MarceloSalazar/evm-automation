"""Escritura de la planificación del cliente con MS Project (comando generar-cliente, sin --simular).

El archivo de origen nunca se abre: se copia a un archivo temporal en la carpeta de destino, y MS Project trabaja
sobre esa copia:

1. Abre la copia (sin macros) y la lee como el resto del programa.
2. Calcula los cambios (evm.client_plan.build_client_plan).
3. Con el cálculo manual y sin la vinculación automática de tareas, con [cliente] actividad_interna = true (por
   defecto):
   - borra las tareas de otras planificaciones y las tareas inactivas que no se muestran, y comprueba que quedan
     exactamente las demás;
   - quita los vínculos que MS Project haya creado al borrar tareas;
   - deja las tareas que no se muestran como «Otras tareas»: sin nombre, código EDT, notas, hipervínculo, contacto ni
     campos de texto o códigos de esquema (salvo los que usa el libro EVM: método, Visible, empresa y reparto);
   con actividad_interna = false:
   - guarda los datos de los resúmenes que pasan a ser tareas (duración, avance, fechas reales, líneas base 0 a 10);
   - borra las tareas de otras planificaciones y las ramas que no se muestran, y comprueba que quedan exactamente las
     tareas que se muestran;
   - restaura los datos de los resúmenes convertidos;
   - quita los vínculos que MS Project haya creado al borrar tareas y crea los vínculos nuevos;
   en los dos casos, vacía el campo del cliente en todas las tareas, quita su nombre (alias) y renombra la
   planificación.
4. Recalcula y compara las fechas de cada tarea con las de la planificación interna. Con [cliente] fijar_fechas, fija
   con «No comenzar antes del» las tareas sin comenzar que se adelantan, y quita los vínculos nuevos de las que se
   atrasan; repite el cálculo hasta que no hay más cambios.
5. Con actividad_interna y [cliente] filtrar_vista, aplica a la vista un filtro que oculta las «Otras tareas».
6. Quita los calendarios que no usa ninguna tarea ni recurso, y las macros; anota las propiedades del archivo, las
   vistas, las tablas, los filtros y los grupos para el informe de control.
7. Guarda la copia, la lee de nuevo y la compara con la planificación interna (evm.client_check): % completado del
   proyecto, libro EVM fila a fila y, con actividad_interna, los datos de cada tarea.
8. Cierra la copia, la mueve al destino y la marca como de solo lectura: el archivo del cliente no se edita, se vuelve
   a generar en cada entrega. Comprueba en el archivo, sin MS Project, que no le queda ningún módulo de macros
   (evm.mpp_macros).

El registro de diagnóstico (-v o --registro) indica cada paso con los UID y los códigos EDT de las tareas, nunca sus
nombres.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import stat
import time
import traceback
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from evm import progress
from evm.client_check import BookComparison, Comparison, compare
from evm.client_plan import (INTERNAL_NAME, ClientPlan, DateChange, ProgressChange, WriteResult, build_client_plan,
                              require_client_field)
from evm.config import Config, ConfigError
from evm.fields import TASK_TEXT_IDS, field_id, parse_custom_name
from evm.messages import EvmError, ProjectAppError
from evm.model import CONSTRAINT_ASAP, Task
from evm.mpp_macros import macro_modules
from evm.reader_com import ComSession, _com_date, _dt, _get, _number
from evm.wbs import plural

log = logging.getLogger(__name__)

PJ_MANUAL = 0                   # PjCalculation: cálculo manual
CONSTRAINT_SNET = 4             # pjSNET: No comenzar antes del
MAX_ROUNDS = 4                  # cálculos de MS Project para fijar las fechas
_BASELINE_FIELDS = ("Start", "Finish", "Duration")
# actividad_interna: datos que se quitan de las «Otras tareas» (salvo los campos que usa el libro EVM)
_CLEARED_FIELDS = ([f"Text{i}" for i in range(1, 31)] + [f"OutlineCode{i}" for i in range(1, 11)]
                   + ["Contact", "Hyperlink", "HyperlinkAddress", "HyperlinkSubAddress", "Notes"])
_BUILTIN_CODES = {"WBS", "OutlineNumber"}
FILTER_NAME = "Tareas del cliente"
# Filtro «Nombre no es igual a Otras tareas», con los nombres en inglés y en español (MS Project en español puede
# pedir el nombre del campo y la comparación en su idioma)
_FILTER_TESTS = (("Name", "does not equal"), ("Nombre", "no es igual a"))
_QUOTED = re.compile(r"[\"'«“‘][^\"'»”’]*[\"'»”’]")
# Contenido del archivo: tipos del Organizador (PjOrganizer) y de los módulos de VBA (vbext_ComponentType)
PJ_MODULES, PJ_CALENDARS = 2, 5
_VBA_MODULES = {1, 2, 3}        # módulo estándar, módulo de clase y formulario
_VBA_DOCUMENT = 100             # ThisProject: no se puede quitar; se le quita el código
# Calendario que crea MS Project al abrir archivos de Project 98 (para sus líneas base): se conserva
_PROJECT98 = re.compile(r"Project\s*98", re.IGNORECASE)
_PROPERTIES = (("Title", "Título"), ("Subject", "Asunto"), ("Author", "Autor"), ("Manager", "Administrador"),
               ("Company", "Organización"), ("Category", "Categoría"), ("Keywords", "Palabras clave"),
               ("Comments", "Comentarios"))
_LISTS = (("Vistas", "ViewList"), ("Tablas de tareas", "TaskTableList"), ("Filtros de tareas", "TaskFilterList"),
          ("Grupos de tareas", "TaskGroupList"))


def com_error_text(exc: BaseException) -> str:
    """Error de COM para los mensajes y el registro: código HRESULT y descripción, sin los textos entre comillas (MS
    Project cita ahí los nombres de las tareas)."""
    args = getattr(exc, "args", ()) or ()
    code = args[0] if args and isinstance(args[0], int) else None
    detail = ""
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        detail = str(args[2][2])
    elif len(args) > 1 and isinstance(args[1], str):
        detail = args[1]
    else:
        detail = str(exc)
    detail = _QUOTED.sub("«…»", " ".join(detail.split()))
    return f"{code & 0xFFFFFFFF:#010x} {detail}".strip() if code is not None else detail


def _same(a: datetime | None, b: datetime | None) -> bool:
    """Misma fecha y hora, al minuto."""
    if a is None or b is None:
        return a is b
    return a.replace(second=0, microsecond=0) == b.replace(second=0, microsecond=0)


def _active(t: Task) -> bool:
    n: Task | None = t
    while n is not None:
        if not n.active:
            return False
        n = n.parent
    return True


def _deps(com_task) -> list:
    deps = _get(com_task, "TaskDependencies")
    if deps is None:
        return []
    try:
        return list(deps)
    except Exception:
        return []


def _ends(dep) -> tuple[int, int] | None:
    try:
        return int(dep.From.UniqueID), int(dep.To.UniqueID)
    except Exception:           # vínculo con otro proyecto
        return None


# Posición de un vínculo en el texto de las predecesoras: signo, número (con coma o punto decimal) y unidad o %
# ('+2 días', '-1,5 sem.', '+3 edías?', '+50%')
_LAG_TEXT = re.compile(r"[+-]\s*\d+(?:[.,]\d+)?\s*(?:[^\W\d_]+\.?\??|%)")
_PRED_TEXT = re.compile(r"(\d+)(?:[^\W\d_]{2})?")


def _pred_uids(text) -> set[int] | None:
    """UID de las predecesoras en el texto de Id. único de predecesoras de MS Project ('12;15FC+2,5 días;7CC'),
    separadas por coma o punto y coma según la configuración regional. None si no se reconoce (p. ej. un vínculo con
    otro proyecto): entonces se recorren los vínculos de la tarea."""
    if not isinstance(text, str):
        return None
    out: set[int] = set()
    for part in re.split(r"[,;]", _LAG_TEXT.sub("", text)):
        part = part.strip()
        if not part:
            continue
        m = _PRED_TEXT.fullmatch(part)
        if m is None:
            return None
        out.add(int(m.group(1)))
    return out


def _items(collection) -> list:
    """Elementos de una colección o lista de MS Project (Count e Item(1..Count))."""
    if collection is None:
        return []
    try:
        return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]
    except Exception:
        try:
            return list(collection)
        except Exception:
            return []


class _Writer:
    """Cambios en la copia abierta en MS Project."""

    def __init__(self, session: ComSession, plan: ClientPlan, cfg: Config, result: WriteResult):
        self.s = session
        self.app = session.app
        self.project = session.project
        self.plan = plan
        self.cfg = cfg
        self.r = result
        self.step = ""
        self.com: dict[int, object] = {}            # UID -> tarea de MS Project
        self.added: dict[int, list] = {}            # UID de la sucesora -> vínculos nuevos creados (objetos COM)
        self.by_uid = {t.uid: t for t in plan.project.tasks}

    # -- utilidades -------------------------------------------------------------------------------------------------
    def _index(self) -> None:
        self.com = {}
        for t in self.project.Tasks:
            if t is None or bool(_get(t, "ExternalTask", False)):
                continue
            try:
                self.com[int(t.UniqueID)] = t
            except Exception:
                continue

    def _calculate(self) -> None:
        for method in ("CalculateProject", "CalculateAll"):
            try:
                getattr(self.app, method)()
                return
            except Exception as exc:
                log.debug("%s no disponible: %s", method, com_error_text(exc))
        raise ProjectAppError("CLI-018", "MS Project no recalculó el proyecto (CalculateProject).")

    def _ref(self, uid: int) -> str:
        t = self.by_uid.get(uid)
        return self.plan.ref(t) if t is not None else f"UID {uid}"

    # -- pasos ------------------------------------------------------------------------------------------------------
    def _steps(self) -> list[tuple[str, object]]:
        if self.plan.internal_activity:
            steps = [("índice de tareas", self._index),
                     ("borrado de tareas", self._delete),
                     ("comprobación de las tareas", self._check_tasks),
                     ("vínculos", self._links),
                     ("tareas internas", self._anonymize)]
        else:
            steps = [("índice de tareas", self._index),
                     ("datos de los resúmenes convertidos", self._snapshot),
                     ("borrado de tareas", self._delete),
                     ("comprobación de las tareas", self._check_tasks),
                     ("restauración de los resúmenes convertidos", self._restore),
                     ("vínculos", self._links)]
        steps += [("campo del cliente", self._clear_field),
                  ("nombre de la planificación", self._rename_root),
                  ("fechas", self._dates)]
        if not self.plan.internal_activity:
            steps.append(("avance", self._progress))
        elif self.cfg.client_filter:
            steps.append(("filtro de la vista", self._filter))
        steps.append(("contenido del archivo", self._content))
        return steps

    def run(self) -> None:
        saved = self._options_off()
        try:
            for step, action in self._steps():
                self.step = step
                log.info("Paso: %s", step)
                action()
        except EvmError:
            raise
        except Exception as exc:
            # El traceback se registra sin el mensaje original de COM, que puede citar nombres de tareas
            log.debug("Error en el paso %s: %s %s\n%s", self.step, type(exc).__name__, com_error_text(exc),
                      "".join(traceback.format_tb(exc.__traceback__)))
            raise ProjectAppError("CLI-018", f"MS Project devolvió un error en el paso «{self.step}»: "
                                             f"{com_error_text(exc)}. El detalle está en el registro de diagnóstico "
                                             "(-v o --registro).") from None
        finally:
            self._options_restore(saved)

    def _options_off(self) -> list[tuple[object, str, object]]:
        """Cálculo manual (cada cambio no recalcula el proyecto), sin vinculación automática de las tareas (al borrar
        una tarea, MS Project podría vincular su predecesora con su sucesora), sin el Asistente para planeamiento
        (sus avisos detendrían la automatización) y sin actualizar la pantalla. Las opciones que esta versión de MS
        Project no tiene se omiten (queda en el registro)."""
        saved = []
        options = [(self.app, "Calculation", PJ_MANUAL), (self.project, "AutoLinkTasks", False),
                   (self.app, "ScreenUpdating", False)]
        options += [(self.app, name, False) for name in ("DisplayPlanningWizard", "DisplayWizardErrors",
                                                         "DisplayWizardScheduling", "DisplayWizardUsage")]
        for obj, name, value in options:
            old = _get(obj, name)
            try:
                setattr(obj, name, value)
                saved.append((obj, name, old))
                log.debug("Opción %s: %r -> %r", name, old, value)
            except Exception as exc:
                log.debug("No se pudo cambiar la opción %s: %s", name, com_error_text(exc))
        return saved

    def _options_restore(self, saved) -> None:
        for obj, name, old in reversed(saved):
            if old is None:
                continue
            try:
                setattr(obj, name, old)
                log.debug("Opción %s restaurada: %r", name, old)
            except Exception as exc:
                log.warning("No se pudo restaurar la opción %s de MS Project: %s", name, com_error_text(exc))

    def _snapshot(self) -> None:
        """Datos de los resúmenes que pasan a ser tareas, tal como los tiene MS Project antes de borrar sus
        subtareas."""
        self.snap: dict[int, dict[str, object]] = {}
        for t in self.plan.collapsed:
            com = self.com[t.uid]
            data = {name: _get(com, name) for name in ("Duration", "PercentComplete", "ActualStart", "ActualFinish",
                                                       "Start", "Finish")}
            for k in range(11):
                sfx = "" if k == 0 else str(k)
                for name in _BASELINE_FIELDS:
                    data[f"Baseline{sfx}{name}"] = _get(com, f"Baseline{sfx}{name}")
            self.snap[t.uid] = data
            log.debug("Resumen que pasa a ser tarea %s: duración %s, %% completado %s, comienzo real %s, fin real %s",
                      self._ref(t.uid), data["Duration"], data["PercentComplete"], _dt(data["ActualStart"]),
                      _dt(data["ActualFinish"]))
        log.info("Resúmenes que pasan a ser tareas: %d", len(self.snap))

    def _delete(self) -> None:
        roots = self.plan.delete_roots
        progress.stage("borrar", "Borrado de las tareas que no se muestran")
        for i, t in enumerate(roots, start=1):
            progress.step(i, len(roots), "ramas")
            com = self.com.get(t.uid)
            if com is None:
                log.warning("No se encontró en MS Project la tarea %s que se debía borrar", self._ref(t.uid))
                continue
            log.debug("Se borra %s y su rama", self._ref(t.uid))
            com.Delete()
        what = "inactivas que no se muestran" if self.plan.internal_activity else "que no se muestran"
        gone = self.plan.dropped if self.plan.internal_activity else self.plan.hidden
        log.info("Ramas borradas: %d (%d tareas de otras planificaciones y %d de la planificación %s)", len(roots),
                 len(self.plan.outside), len(gone), what)
        self._index()

    def _check_tasks(self) -> None:
        """Deben quedar exactamente las tareas del archivo del cliente (las que se muestran o, con actividad_interna,
        todas salvo las inactivas que no se muestran, y la tarea de la planificación): si no, el archivo no se
        guarda."""
        expected = self.plan.remaining
        found = set(self.com)
        extra, missing = sorted(found - expected), sorted(expected - found)
        if extra or missing:
            log.error("Tareas que sobran: %s; tareas que faltan: %s", extra[:50], missing[:50])
            raise ProjectAppError("CLI-018", f"Después de borrar las tareas quedan {len(found)} tareas y se "
                                             f"esperaban {len(expected)} ({len(extra)} que sobran y {len(missing)} que "
                                             "faltan): no se generó el archivo del cliente. Los UID están en el "
                                             "registro de diagnóstico.")
        log.info("Tareas en el archivo del cliente: %d", len(found))

    def _restore(self) -> None:
        """Un resumen sin subtareas pasa a ser una tarea: se le devuelven la duración, el avance y las líneas base que
        tenía, por si MS Project los cambió al borrar sus subtareas."""
        for t in self.plan.collapsed:
            com, data = self.com[t.uid], self.snap[t.uid]
            problems = []
            if bool(_get(com, "Summary", False)):
                problems.append("sigue siendo un resumen")
            actual_start, actual_finish = _dt(data["ActualStart"]), _dt(data["ActualFinish"])
            duration, pct = _number(data["Duration"]), _number(data["PercentComplete"])
            changes = []
            try:
                if actual_start is not None and not _same(_dt(_get(com, "ActualStart")), actual_start):
                    com.ActualStart = _com_date(actual_start)
                    changes.append("comienzo real")
                if actual_finish is not None:
                    if not _same(_dt(_get(com, "ActualFinish")), actual_finish):
                        com.ActualFinish = _com_date(actual_finish)
                        changes.append("fin real")
                else:
                    if duration is not None and _number(_get(com, "Duration")) != duration:
                        com.Duration = duration
                        changes.append("duración")
                    if actual_start is not None and pct is not None and _number(_get(com, "PercentComplete")) != pct:
                        com.PercentComplete = pct
                        changes.append("% completado")
            except Exception as exc:
                problems.append(f"avance: {com_error_text(exc)}")
            for k in range(11):
                sfx = "" if k == 0 else str(k)
                for name in _BASELINE_FIELDS:
                    key = f"Baseline{sfx}{name}"
                    before, now = data[key], _get(com, key)
                    if name == "Duration":
                        if _dt(data[f"Baseline{sfx}Start"]) is None or _number(before) == _number(now):
                            continue
                        value = _number(before)
                    else:
                        if _dt(before) is None or _same(_dt(before), _dt(now)):
                            continue
                        value = _com_date(_dt(before))
                    try:
                        setattr(com, key, value)
                        changes.append(key)
                    except Exception as exc:
                        problems.append(f"{key}: {com_error_text(exc)}")
            if changes:
                log.debug("Resumen convertido %s: se restauró %s", self._ref(t.uid), ", ".join(changes))
            if problems:
                self.r.restore_problems.append((t, "; ".join(problems)))
                log.warning("Resumen convertido %s: %s", self._ref(t.uid), "; ".join(problems))

    def _links(self) -> None:
        """Quita los vínculos que no son de la planificación (los que MS Project creó al borrar tareas) y crea los
        vínculos nuevos. Las predecesoras de cada tarea se leen de una vez (Id. único de predecesoras): solo se
        recorren, vínculo a vínculo, los de las tareas cuyas predecesoras no son las esperadas."""
        expected = self.plan.kept_pairs
        preds_of: dict[int, set[int]] = defaultdict(set)
        for a, b in expected:
            preds_of[b].add(a)
        found: dict[tuple[int, int], object] = {}
        missing: list[tuple[int, int]] = []
        walked = 0
        for i, (uid, com) in enumerate(self.com.items(), start=1):
            progress.step(i, len(self.com), "tareas (vínculos)")
            want = preds_of.get(uid, set())
            if _pred_uids(_get(com, "UniqueIDPredecessors")) == want:
                continue
            walked += 1
            have: set[int] = set()
            for dep in _deps(com):
                ends = _ends(dep)
                if ends is not None and ends[1] == uid:
                    found[ends] = dep
                    have.add(ends[0])
            missing += [(a, uid) for a in sorted(want - have)]
        for pair in sorted(set(found) - expected):
            log.debug("Vínculo creado por MS Project: %s -> %s; se quita", self._ref(pair[0]), self._ref(pair[1]))
            found[pair].Delete()
            self.r.removed_links.append(pair)
        if missing:
            log.warning("Faltan %d vínculos de la planificación entre tareas que se muestran: %s", len(missing),
                        [(self._ref(a), self._ref(b)) for a, b in missing[:20]])
        log.info("Vínculos que se conservan: %d; creados por MS Project y quitados: %d; tareas revisadas vínculo a "
                 "vínculo: %d de %d", len(expected) - len(missing), len(self.r.removed_links), walked, len(self.com))
        progress.stage("vincular", "Vínculos nuevos")
        for i, link in enumerate(self.plan.new_links, start=1):
            progress.step(i, len(self.plan.new_links), "vínculos")
            pred, succ = self.com[link.pred.uid], self.com[link.succ.uid]
            lag = int(round(link.lag))
            try:
                dep = succ.TaskDependencies.Add(pred, link.kind, lag)
            except Exception as exc:
                error = com_error_text(exc)
                self.r.rejected.append((link, error))
                log.warning("MS Project no aceptó el vínculo %s -> %s (%s, %d min): %s", self._ref(link.pred.uid),
                            self._ref(link.succ.uid), link.kind, lag, error)
                continue
            self.added.setdefault(link.succ.uid, []).append(dep)
            kind, got = _number(_get(dep, "Type")), _number(_get(dep, "Lag"))
            if kind != link.kind or got is None or abs(got - lag) > 1:
                log.warning("Vínculo %s -> %s: se pidió tipo %d y posición %d min; MS Project tiene tipo %s y "
                            "posición %s", self._ref(link.pred.uid), self._ref(link.succ.uid), link.kind, lag, kind,
                            got)
            else:
                log.debug("Vínculo nuevo %s -> %s: tipo %d, posición %d min", self._ref(link.pred.uid),
                          self._ref(link.succ.uid), link.kind, lag)
        log.info("Vínculos nuevos: %d creados, %d no aceptados", len(self.plan.new_links) - len(self.r.rejected),
                 len(self.r.rejected))

    def _attr(self, name: str) -> str | None:
        """Propiedad de la tarea de MS Project para un campo configurado (por ejemplo 'Text10'), o None."""
        ref = self.s._resolve_field(name) if name else None
        return str(ref[1]) if ref is not None and ref[0] == "attr" else None

    def _anonymize(self) -> None:
        """Las tareas que no se muestran quedan como «Otras tareas»: sin nombre, sin código EDT (si está en un campo
        personalizado), sin notas, hipervínculo ni contacto, y sin los campos de texto y códigos de esquema que no usa
        el libro EVM. Conservan las fechas, la duración, el avance, las líneas base, los vínculos y los recursos."""
        cfg = self.cfg
        keep = {self._attr(name) for name in (cfg.method_field, cfg.visible_field, cfg.milestone_weight_field)}
        if cfg.multi_enabled:
            keep |= {self._attr(cfg.company_field), self._attr(cfg.split_field)}
        keep.discard(None)
        wbs = self.s._resolve_field(cfg.wbs_field) or ("attr", "WBS")
        code_fid = None
        if wbs[0] == "attr":
            self.r.codes_removed = str(wbs[1]) not in _BUILTIN_CODES
            keep.discard(str(wbs[1]))
        else:
            code_fid = int(wbs[1])
        self.plan.codes_removed = self.r.codes_removed
        fields = [name for name in _CLEARED_FIELDS if name not in keep]
        log.info("«%s»: campo EDT %s (%s); se conservan %s", INTERNAL_NAME, wbs[1],
                 "se quita" if self.r.codes_removed else "no se puede quitar", sorted(keep) or "ningún campo")
        cleared: dict[str, int] = {}
        problems: dict[str, tuple[int, str]] = {}

        def clear(com, name: str, action) -> None:
            try:
                action()
                cleared[name] = cleared.get(name, 0) + 1
            except Exception as exc:
                n, first = problems.get(name, (0, com_error_text(exc)))
                problems[name] = (n + 1, first)

        targets = self.plan.anonymous
        progress.stage("anonimizar", f"Tareas que no se muestran como «{INTERNAL_NAME}»")
        for i, t in enumerate(targets, start=1):
            progress.step(i, len(targets), "tareas")
            com = self.com[t.uid]
            clear(com, "Name", lambda: setattr(com, "Name", INTERNAL_NAME))
            for name in fields:
                value = _get(com, name)
                if value is None or str(value) == "":
                    continue
                clear(com, name, lambda name=name: setattr(com, name, ""))
            if code_fid is not None:
                clear(com, "EDT", lambda: com.SetField(code_fid, ""))
        # Un campo con fórmula no se puede vaciar: su valor lo calcula MS Project con los datos que se conservan
        calculated = {name: problems.pop(name)[0] for name in sorted(problems) if self._has_formula(name)}
        self.r.anonymized = len(targets)
        self.r.not_cleared = problems
        self.r.calculated = calculated
        log.info("Tareas como «%s»: %d; datos quitados por campo: %s", INTERNAL_NAME, len(targets),
                 ", ".join(f"{k} {v}" for k, v in sorted(cleared.items())) or "ninguno")
        if calculated:
            log.info("Campos calculados con fórmula (no se vacían): %s",
                     ", ".join(f"{k} en {v} tareas" for k, v in calculated.items()))
        for name, (n, error) in sorted(problems.items()):
            log.warning("MS Project no dejó vaciar %s en %d tareas: %s", name, n, error)

    def _has_formula(self, name: str) -> bool:
        """El campo personalizado de texto (Text1..30) tiene una fórmula."""
        m = re.fullmatch(r"Text(\d+)", name)
        fid = TASK_TEXT_IDS.get(int(m.group(1))) if m else None
        if fid is None:
            return False
        try:
            return bool(str(self.app.CustomFieldGetFormula(fid) or "").strip())
        except Exception as exc:
            log.debug("No se pudo leer la fórmula de %s: %s", name, com_error_text(exc))
            return False

    def _filter(self) -> None:
        """Filtro de la vista que oculta las «Otras tareas» (con sus resúmenes, que también lo son): el archivo se
        abre con él. Si MS Project no lo acepta, el archivo se genera igual (CLI-028)."""
        for field_name, test in _FILTER_TESTS:
            try:
                self.app.FilterEdit(Name=FILTER_NAME, TaskFilter=True, Create=True, OverwriteExisting=True,
                                    FieldName=field_name, Test=test, Value=INTERNAL_NAME, ShowInMenu=True,
                                    ShowSummaryTasks=True)
                self.app.FilterApply(Name=FILTER_NAME)
            except Exception as exc:
                log.debug("Filtro con %s / %s: %s", field_name, test, com_error_text(exc))
                continue
            self.r.filter_applied = True
            log.info("Filtro de la vista «%s» aplicado (%s %s «%s»)", FILTER_NAME, field_name, test, INTERNAL_NAME)
            return
        self.r.filter_applied = False
        log.warning("No se pudo crear el filtro de la vista «%s»", FILTER_NAME)

    def _content(self) -> None:
        """Lo que no son tareas: quita los calendarios sin uso y las macros, y anota el resto (propiedades del
        archivo, vistas, tablas, filtros y grupos) para el informe de control."""
        progress.stage("contenido", "Calendarios y macros del archivo")
        c = self.r.content
        self._calendars()
        self._macros()
        for prop, label in _PROPERTIES:
            if str(_get(self.project, prop, "") or "").strip():
                c.properties.append(label)
        for label, prop in _LISTS:
            c.lists[label] = [str(x) for x in _items(_get(self.project, prop)) if str(x or "").strip()]
        c.checked = True
        log.info("Contenido del archivo: %d calendarios se conservan, %d sin uso quitados, %d sin uso que no se "
                 "pudieron quitar; macros quitadas: %s; propiedades con valor: %s", len(c.calendars),
                 len(c.removed_calendars), len(c.failed_calendars), ", ".join(c.removed_macros) or "ninguna",
                 ", ".join(c.properties) or "ninguna")

    def _calendars(self) -> None:
        """Calendarios base que no usa ninguna tarea ni recurso, ni el proyecto: se quitan (salvo el de las líneas base
        de Project 98, que crea MS Project). Si no se puede leer el calendario del proyecto, no se quita ninguno."""
        c = self.r.content
        calendars = _items(_get(self.project, "BaseCalendars"))
        project_calendar = str(_get(_get(self.project, "Calendar"), "Name", "") or "")
        tasks: Counter = Counter()
        for com in self.com.values():
            own = _get(com, "CalendarObject")              # Nothing si la tarea no tiene calendario propio
            name = str(_get(own, "Name", "") or "") if own is not None else ""
            if name:
                tasks[name] += 1
        resources: Counter = Counter()
        for res in _items(_get(self.project, "Resources")):
            name = str(_get(res, "BaseCalendar", "") or "") if res is not None else ""
            if name:
                resources[name] += 1
        removed: list[tuple[str, object]] = []
        for cal in calendars:
            name = str(_get(cal, "Name", "") or "")
            uses = []
            if name == project_calendar:
                uses.append("calendario del proyecto")
            if tasks[name]:
                uses.append(plural(tasks[name], "tarea", "tareas"))
            if resources[name]:
                uses.append(plural(resources[name], "recurso", "recursos"))
            if uses:
                c.calendars.append((name, ", ".join(uses)))
            elif not project_calendar:
                c.calendars.append((name, "no se comprobó su uso (no se pudo leer el calendario del proyecto)"))
            elif _PROJECT98.search(name):
                c.calendars.append((name, "de MS Project (líneas base de Project 98)"))
            else:
                removed.append((name, cal))
        for name, cal in removed:
            try:
                cal.Delete()
            except Exception as exc:
                first = com_error_text(exc)
                try:
                    self.app.OrganizerDeleteItem(Type=PJ_CALENDARS, FileName=_get(self.project, "Name", ""),
                                                 Name=name)
                except Exception:
                    c.failed_calendars.append((name, first))
                    log.warning("No se pudo quitar el calendario sin uso %s: %s", name, first)
                    continue
            c.removed_calendars.append(name)
            log.debug("Calendario sin uso quitado: %s", name)
        # Comprobación: los quitados ya no están
        if c.removed_calendars:
            still = {str(_get(cal, "Name", "") or "") for cal in _items(_get(self.project, "BaseCalendars"))}
            for name in [n for n in c.removed_calendars if n in still]:
                c.removed_calendars.remove(name)
                c.failed_calendars.append((name, "sigue en el archivo"))
                log.warning("El calendario %s sigue en el archivo después de quitarlo", name)

    def _macros(self) -> None:
        """Quita los módulos de macros con el proyecto VBA (también el código de ThisProject) o, si MS Project no da
        acceso a él (Centro de confianza), con el Organizador, por el nombre de los módulos leídos del archivo. Al
        final, generate_client comprueba en el archivo guardado que no queda ninguno."""
        c = self.r.content
        project = None
        components = None
        try:
            project = self.project.VBProject
            components = _items(project.VBComponents)
        except Exception as exc:
            c.macro_error = com_error_text(exc)
            log.debug("Sin acceso al proyecto VBA: %s", c.macro_error)
        pending = list(c.macros or [])
        if components is not None:
            c.macro_method = "proyecto VBA"
            for comp in components:
                kind = int(_number(_get(comp, "Type")) or 0)
                name = str(_get(comp, "Name", "") or "")
                try:
                    if kind in _VBA_MODULES:
                        project.VBComponents.Remove(comp)
                        c.removed_macros.append(name)
                        pending = [m for m in pending if m != name]
                    elif kind == _VBA_DOCUMENT:
                        code = _get(comp, "CodeModule")
                        lines = int(_number(_get(code, "CountOfLines")) or 0)
                        if lines > int(_number(_get(code, "CountOfDeclarationLines")) or 0):
                            code.DeleteLines(1, lines)
                            c.removed_macros.append(f"código de {name}")
                except Exception as exc:
                    log.warning("No se pudo quitar el módulo de macros %s con el proyecto VBA: %s", name,
                                com_error_text(exc))
        if pending:
            c.macro_method = c.macro_method or "Organizador"
            for name in pending:
                try:
                    self.app.OrganizerDeleteItem(Type=PJ_MODULES, FileName=_get(self.project, "Name", ""), Name=name)
                    c.removed_macros.append(name)
                except Exception as exc:
                    log.warning("No se pudo quitar el módulo de macros %s con el Organizador: %s", name,
                                com_error_text(exc))

    def _field_id(self) -> int | None:
        ref = self.s._resolve_field(self.cfg.client_field)
        if ref is None:
            return None
        kind, key = ref
        if kind == "fid":
            return int(key)
        parsed = parse_custom_name(str(key))
        return field_id(*parsed) if parsed else None

    def _clear_field(self) -> None:
        """Vacía el campo del cliente en todas las tareas y quita su nombre (el cliente no debe verlo)."""
        ref = self.s._resolve_field(self.cfg.client_field)
        if ref is None:
            raise ProjectAppError("CLI-018", f"No se encontró el campo '{self.cfg.client_field}' para vaciarlo.")
        kind, key = ref
        cleared = 0
        for uid, com in self.com.items():
            t = self.by_uid.get(uid)
            if t is not None and not t.client_mark:
                continue
            if kind == "attr":
                setattr(com, key, "")
            else:
                com.SetField(key, "")
            cleared += 1
        log.info("Campo '%s' vaciado en %d tareas", self.cfg.client_field, cleared)
        fid = self._field_id()
        removed = False
        if fid is not None:
            try:
                self.app.CustomFieldRename(fid, "")
                try:
                    removed = not str(self.app.CustomFieldGetName(fid) or "").strip()
                except Exception:           # sin forma de comprobarlo: el cambio de nombre no dio error
                    removed = True
            except Exception as exc:
                log.warning("No se pudo quitar el nombre del campo %s (FieldID %d): %s", self.cfg.client_field, fid,
                            com_error_text(exc))
        else:
            log.warning("No se encontró el FieldID del campo '%s'", self.cfg.client_field)
        self.r.field_name_removed = removed
        log.info("Nombre del campo '%s': %s", self.cfg.client_field, "quitado" if removed else "no se pudo quitar")

    def _rename_root(self) -> None:
        root, name = self.plan.root, self.plan.new_name
        if root is None or not name or name == root.name:
            return
        self.com[root.uid].Name = name
        log.info("Tarea de la planificación (UID %d) renombrada a %s", root.uid, name)

    def _dates(self) -> None:
        """Recalcula y compara las fechas de cada tarea que se muestra (sin los resúmenes) con las de la planificación
        interna. Fija las que se adelantan (sin comenzar, sin otra delimitación) con «No comenzar antes del» y quita
        los vínculos nuevos de las que se atrasan; repite hasta que no hay cambios. También las tareas inactivas:
        MS Project las sigue programando con sus predecesoras (aunque no muevan a sus sucesoras), así que se adelantan
        si se borra una predecesora oculta."""
        progress.stage("fechas", "Cálculo y comparación de las fechas")
        collapsed = {t.uid for t in self.plan.collapsed}
        if self.plan.internal_activity:      # todas las tareas de detalle que quedan (las demás se comparan al final)
            keep = self.plan.remaining
            targets = [t for t in self.plan.tasks if t.uid in keep and not t.children]
        else:
            targets = [t for t in self.plan.tasks if t.uid in self.plan.shown
                       and (t.uid in collapsed or not t.children)]
        fix = self.cfg.client_fix_dates
        fixed: dict[int, Task] = {}
        for rounds in range(1, MAX_ROUNDS + 1):
            self.r.rounds = rounds
            self._calculate()
            changes = 0
            moved = early = late = 0
            for i, t in enumerate(targets, start=1):
                progress.step(i, len(targets), f"tareas (cálculo {rounds})")
                com = self.com[t.uid]
                start, finish = _dt(_get(com, "Start")), _dt(_get(com, "Finish"))
                if _same(start, t.start) and _same(finish, t.finish):
                    continue
                moved += 1
                if t.actual_start is not None or start is None or t.start is None:
                    continue
                if start < t.start:
                    early += 1
                    if fix and t.constraint_type == CONSTRAINT_ASAP and t.uid not in fixed:
                        com.ConstraintType = CONSTRAINT_SNET
                        com.ConstraintDate = _com_date(t.start)
                        fixed[t.uid] = t
                        changes += 1
                        log.debug("Fecha fijada %s: %s (MS Project la adelantaba a %s)", self._ref(t.uid), t.start,
                                  start)
                elif start > t.start:
                    late += 1
                    for dep in self.added.pop(t.uid, []):
                        log.debug("Se quita un vínculo nuevo de %s, que la atrasaba a %s (interna %s)",
                                  self._ref(t.uid), start, t.start)
                        dep.Delete()
                        changes += 1
            log.info("Cálculo %d: %d tareas con fechas distintas (%d adelantadas, %d atrasadas); %d cambios", rounds,
                     moved, early, late, changes)
            if not changes:
                break
        if changes:
            self.r.rounds += 1
            self._calculate()
        self.r.fixed = list(fixed.values())
        if self.plan.internal_activity:      # la comparación final (evm.client_check) revisa cada tarea
            log.info("Fechas: %d fijadas", len(self.r.fixed))
            return
        for t in targets:
            com = self.com[t.uid]
            start, finish = _dt(_get(com, "Start")), _dt(_get(com, "Finish"))
            if _same(start, t.start) and _same(finish, t.finish):
                continue
            if t.actual_finish is not None:
                note = "terminada"
            elif t.actual_start is not None:
                note = "comenzada: el fin depende del trabajo restante"
            elif start is not None and t.start is not None and start < t.start and \
                    t.constraint_type != CONSTRAINT_ASAP:
                note = "tiene otra delimitación"
            elif start is not None and t.start is not None and start < t.start and not fix:
                note = "fijar_fechas = false"
            elif _same(start, t.start):
                note = "misma fecha de comienzo, otro fin"
            else:
                note = "se atrasa respecto de la planificación interna"
            self.r.differences.append(DateChange(t, start, finish, note))
            log.debug("Fecha distinta %s: interna %s - %s; cliente %s - %s (%s)", self._ref(t.uid), t.start, t.finish,
                      start, finish, note)
        log.info("Fechas: %d fijadas, %d distintas de la planificación interna", len(self.r.fixed),
                 len(self.r.differences))

    def _progress(self) -> None:
        """Compara el % completado de cada tarea que se muestra con el de la planificación interna: las tareas de
        detalle y los resúmenes convertidos deben conservarlo; el de los resúmenes lo calcula MS Project con las
        tareas que quedan (se guarda para el informe y se compara con el previsto)."""
        collapsed = {t.uid for t in self.plan.collapsed}
        expected = {c.task.uid: c.client for c in self.plan.progress_changes}
        summaries = changed = 0
        for t in self.plan.tasks:
            if t.uid not in self.plan.shown or not _active(t):
                continue
            pct = _number(_get(self.com[t.uid], "PercentComplete"))
            if pct is None:
                continue
            if t.children and t.uid not in collapsed:
                summaries += 1
                self.r.summary_percent[t.uid] = pct
                predicted = expected.get(t.uid, t.percent_complete)
                if abs(pct - predicted) >= 1:
                    log.warning("Resumen %s: MS Project calcula %s %% y se preveía %.1f %% (interno %s %%)",
                                self._ref(t.uid), pct, predicted, t.percent_complete)
                elif abs(pct - t.percent_complete) >= 0.5:
                    changed += 1
                    log.debug("Resumen %s: %% completado %s (interno %s), como se preveía", self._ref(t.uid), pct,
                              t.percent_complete)
            elif abs(pct - t.percent_complete) >= 0.5:
                self.r.progress_tasks.append(ProgressChange(t, pct))
                log.warning("Tarea %s: %% completado %s en el archivo del cliente e interno %s", self._ref(t.uid), pct,
                            t.percent_complete)
        log.info("Avance: %d resúmenes, %d con otro %% completado; %d tareas que no conservan el suyo", summaries,
                 changed, len(self.r.progress_tasks))


def _log_comparison(plan: ClientPlan, result: WriteResult) -> None:
    c = result.comparison
    b = c.book
    log.info("Comparación: %% completado del proyecto %s (interno) y %s (cliente); libro EVM: %d filas iguales, %d con "
             "otro código, %d con otras tareas o valores, %d solo en el interno, %d solo en el del cliente%s; tareas "
             "comparadas %d, con datos distintos %d", c.percent_internal, c.percent_client, b.same, len(b.moved),
             len(b.different), len(b.only_internal), len(b.only_client), f" ({b.error})" if b.error else "",
             c.checked, len(c.tasks))
    for d in c.tasks[:200]:
        log.debug("Datos distintos %s: %s", plan.ref(d.task),
                  "; ".join(f"{label} {a} -> {v}" for label, a, v in d.fields))


def _writable(path: Path) -> None:
    if path.exists():
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)


def generate_client(origen: str | Path, destino: str | Path, cfg: Config, *, visible: bool = False,
                    overwrite: bool = False, session_cls=None) -> tuple[ClientPlan, WriteResult]:
    """Genera el .mpp del cliente a partir del .mpp de origen, que no se modifica. Devuelve los cambios calculados y
    el resultado de la escritura. session_cls: la sesión de MS Project (por defecto ComSession; las pruebas usan una
    simulada)."""
    session_cls = session_cls or ComSession
    origen, destino = Path(origen).resolve(), Path(destino).resolve()
    if destino.exists() and not overwrite:
        raise ConfigError("CFG-042", f"'{destino}' ya existe. Para reemplazarlo se debe usar --sobrescribir.")
    temp = destino.with_name(f"~{destino.stem}.generando{destino.suffix}")
    _writable(temp)
    temp.unlink(missing_ok=True)
    progress.stage("copiar", "Copia del archivo de origen")
    shutil.copyfile(origen, temp)
    log.info("Copia temporal: %s (%d bytes)", temp, temp.stat().st_size)
    result = WriteResult(path=str(destino))
    result.content.macros = macro_modules(temp)
    log.info("Módulos de macros del archivo de origen: %s", "no se pudieron leer" if result.content.macros is None
             else ", ".join(result.content.macros) or "ninguno")
    done = False
    try:
        progress.stage("abrir:", "Apertura de la copia en MS Project")
        with session_cls(str(temp), visible=visible, read_only=False) as s:
            log.info("MS Project %s (compilación %s)", _get(s.app, "Version", "?"), _get(s.app, "Build", "?"))
            names = s._open_project_names() or set()
            result.source_open = str(origen).lower() in names
            if result.source_open:
                log.warning("El archivo de origen está abierto en MS Project: se usa la versión guardada")
            progress.stage("leer:", "Lectura de la planificación")
            project = s.read_project(cfg, client=True)
            require_client_field(project, cfg)
            progress.stage("calcular:", "Cálculo de la planificación del cliente")
            plan = build_client_plan(project, cfg)
            log.info("Tareas: %d de la planificación, %d se muestran, %d de otras planificaciones; vínculos nuevos: "
                     "%d", len(plan.tasks), len(plan.shown), len(plan.outside), len(plan.new_links))
            _Writer(s, plan, cfg, result).run()
            progress.stage("guardar", "Guardado del archivo del cliente")
            s.save()
            log.info("Copia guardada")
            progress.stage("comparar:", "Comparación con la planificación interna")
            try:
                client = s.read_project(cfg, links=False)
                result.comparison = compare(plan, client, cfg, generated=True)
            except Exception as exc:        # el archivo ya está guardado: se entrega con el aviso CLI-029
                log.debug("Error en la comparación: %s %s\n%s", type(exc).__name__, com_error_text(exc),
                          "".join(traceback.format_tb(exc.__traceback__)))
                result.comparison = Comparison(True, None, None, BookComparison(
                    error=f"error inesperado ({type(exc).__name__}); el detalle está en el registro de diagnóstico"))
            _log_comparison(plan, result)
        done = True
    finally:
        if not done:
            try:
                _writable(temp)
                temp.unlink(missing_ok=True)
            except OSError as exc:
                log.warning("No se pudo borrar la copia temporal %s: %s", temp, exc)
    _writable(destino)
    for attempt in range(20):           # MS Project puede tardar un momento en liberar la copia al cerrarla
        try:
            os.replace(temp, destino)
            break
        except PermissionError:
            if attempt == 19:
                raise
            log.debug("La copia sigue bloqueada; nuevo intento en 0,5 s")
            time.sleep(0.5)
    os.chmod(destino, stat.S_IREAD)
    log.info("Archivo del cliente: %s (de solo lectura)", destino)
    # Comprobación en el archivo guardado, sin MS Project: no debe quedar ningún módulo de macros
    result.content.macros_left = macro_modules(destino)
    log.info("Módulos de macros del archivo del cliente: %s", "no se pudieron leer"
             if result.content.macros_left is None else ", ".join(result.content.macros_left) or "ninguno")
    return plan, result
