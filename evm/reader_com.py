"""Lector por automatización COM de MS Project de escritorio (Windows + pywin32).

Abre el archivo .mpp en modo solo lectura y lee las tareas: fechas de la línea
base, fechas actuales, avance y campos personalizados. Los costos no se leen: el
BAC se ingresa en el libro.

Precauciones:
- Si MS Project ya estaba abierto, se usa esa instancia y no se cierra.
- El archivo se abre solo lectura, sin macros, y se cierra sin guardar.
- Las fechas se envían a COM con zona horaria UTC explícita (pywin32
  convierte las fechas sin zona como hora local y las desplazaría).
"""
from __future__ import annotations

import gc
import logging
import time
from datetime import date, datetime, timezone
from pathlib import Path

from evm.config import Config
from evm.fields import (TASK_NUMBER_IDS, TASK_TEXT_IDS, builtin_kind, is_spanish_wbs_name, parse_custom_name,
                        parse_number, same_name)
from evm.model import CONSTRAINT_ASAP, LINK_FS, Link, ProjectData, Task, use_current_schedule

log = logging.getLogger(__name__)

# Constantes de MS Project (referencia VBA)
PJ_DO_NOT_SAVE = 0                            # pjDoNotSave
PJ_POOL_READ_ONLY = 1                         # pjPoolReadOnly (no pregunta; abre subproyectos)
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3     # sin macros

PROGID = "MSProject.Application"
# Espera máxima (segundos) a que responda MS Project: una instancia que se está cerrando (p. ej. la que
# cerró una ejecución anterior) o que muestra un diálogo rechaza las llamadas por un momento.
CONNECT_WAIT = 30.0


def _dt(v) -> datetime | None:
    """Fecha COM -> datetime sin zona. 'NA' u otros valores -> None."""
    if isinstance(v, datetime):
        if not 1984 <= v.year <= 2149:
            return None
        return datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
    return None


def _com_date(d: date | datetime) -> datetime:
    """datetime con zona UTC para que pywin32 no lo desplace."""
    if isinstance(d, datetime):
        return datetime(d.year, d.month, d.day, d.hour, d.minute, tzinfo=timezone.utc)
    return datetime(d.year, d.month, d.day, tzinfo=timezone.utc)


def _f(v, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        n = parse_number(v)
        return default if n is None else n


def _hresult(exc) -> int:
    """Código HRESULT de un error COM, sin signo."""
    code = getattr(exc, "hresult", None)
    if code is None and getattr(exc, "args", None):
        code = exc.args[0]
    return int(code or 0) & 0xFFFFFFFF


def _get(obj, name: str, default=None):
    try:
        return getattr(obj, name)
    except Exception:  # propiedad inexistente en esta versión de Project, etc.
        return default


def _number(v) -> float | None:
    """Número de COM; None si no es numérico (p. ej. la duración en texto de una tarea manual)."""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)


def _predecessors(t, uid: int) -> list[Link]:
    """Vínculos con las predecesoras de la tarea (TaskDependencies tiene también los de sus sucesoras)."""
    links: list[Link] = []
    deps = _get(t, "TaskDependencies")
    if deps is None:
        return links
    try:
        items = list(deps)
    except Exception:
        return links
    for dep in items:
        try:
            if int(dep.To.UniqueID) != uid:
                continue
            pred = int(dep.From.UniqueID)
        except Exception:               # vínculo con otro proyecto que no se puede leer
            continue
        kind = _number(_get(dep, "Type", LINK_FS))
        links.append(Link(pred, LINK_FS if kind is None else int(kind), _number(_get(dep, "Lag", 0)) or 0.0))
    return links


class ComSession:
    """Sesión de MS Project. Uso: with ComSession(ruta) as s: s.read_project(cfg)."""

    def __init__(self, origen: str, visible: bool = False):
        self.origen = origen
        self.visible = visible
        self.app = None
        self.project = None
        self.started_by_us = False
        self.opened_by_us = False
        self._orig_status = None
        self._status_changed = False
        self._pythoncom = None
        self._prev_active = None            # proyecto activo antes de empezar (si ya había uno abierto)
        self._saved_app_props: dict[str, object] = {}
        self._opened_names: set[str] = set()    # proyectos que abrió FileOpenEx (el archivo, el fondo de recursos...)

    # -- conexión -------------------------------------------------------------
    def __enter__(self) -> "ComSession":
        try:
            import pythoncom
        except ImportError as exc:
            raise RuntimeError(
                "Para leer archivos .mpp se necesita pywin32 (pip install pywin32) y MS Project instalado. "
                "Alternativa: guardar el proyecto como XML desde MS Project (Archivo > Guardar como > XML) y "
                "procesar ese archivo.") from exc
        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        try:
            self._connect()
            self._open()
        except BaseException:
            self._cleanup()
            raise
        return self

    def __exit__(self, *exc) -> None:
        self._cleanup()

    def _connect(self) -> None:
        import pywintypes
        import win32com.client
        deadline = time.monotonic() + CONNECT_WAIT
        raw, started = None, False
        while True:
            if raw is None:
                try:
                    raw = win32com.client.GetActiveObject(PROGID)
                    started = False
                except pywintypes.com_error:
                    try:
                        raw = win32com.client.DispatchEx(PROGID)
                    except pywintypes.com_error as exc:
                        raise RuntimeError(
                            "No se pudo iniciar MS Project por automatización COM. Se debe verificar que MS "
                            "Project de escritorio esté instalado, o bien exportar el proyecto a XML y usar ese "
                            "archivo.") from exc
                    started = True
            try:
                n_open = int(raw.Projects.Count)        # la instancia responde
                break
            except pywintypes.com_error as exc:
                # Instancia que se está cerrando (servidor RPC no disponible, 0x800706BA) o que muestra un
                # diálogo (llamada rechazada): se espera y se vuelve a intentar. Una instancia ajena se vuelve
                # a buscar (al terminar de cerrarse se inicia otra); la propia se espera (no se inician más).
                if not started:
                    raw = None
                if time.monotonic() >= deadline:
                    raise RuntimeError(
                        f"MS Project no responde por automatización COM (código {_hresult(exc):#010x}): puede estar "
                        "cerrándose o mostrando un diálogo (inicio de sesión, recuperación de archivos...). Se debe "
                        "abrir MS Project, cerrar el diálogo o esperar a que termine de cerrarse (o cerrarlo desde "
                        "el Administrador de tareas) y volver a intentar.") from exc
                time.sleep(1)
        # Si en realidad se conectó a un MS Project abierto por el usuario, no cerrarlo después.
        self.started_by_us = started and not bool(_get(raw, "UserControl", True)) and n_open == 0
        # Enlace temprano (más rápido) si está disponible; si falla, enlace dinámico.
        try:
            from win32com.client import gencache
            self.app = gencache.EnsureDispatch(raw._oleobj_)
        except Exception as exc:  # caché gen_py dañada, sin biblioteca de tipos, etc.
            log.debug("Sin enlace temprano (%s); se usa enlace dinámico", exc)
            self.app = win32com.client.dynamic.Dispatch(raw._oleobj_)
        if self.started_by_us:
            try:
                self.app.Visible = bool(self.visible)
            except Exception:
                pass

    def _open(self) -> None:
        app = self.app
        if self.origen.lower() == "activo":
            if int(app.Projects.Count) == 0:
                raise RuntimeError("No hay proyectos abiertos en MS Project.")
            self.project = app.ActiveProject
            return
        full = str(Path(self.origen).resolve())
        if int(app.Projects.Count) > 0:
            self._prev_active = _get(app, "ActiveProject")
        for i in range(1, int(app.Projects.Count) + 1):
            p = app.Projects.Item(i)
            if str(_get(p, "FullName", "")).lower() == full.lower():
                self.project = p          # ya estaba abierto: se usa sin cerrarlo
                self._activate()
                return
        if not self.started_by_us:
            self._saved_app_props = {prop: _get(app, prop) for prop in ("DisplayAlerts", "AutomationSecurity")}
        for prop, value in (("DisplayAlerts", False), ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)):
            try:
                setattr(app, prop, value)
            except Exception:
                pass
        import pywintypes
        # pythoncom.Empty (no Missing): pywin32 corta la lista de argumentos en el primer Missing.
        e = self._pythoncom.Empty
        before = self._open_project_names() or set()
        try:
            # Name, ReadOnly, ..., NoAuto, ..., openPool, ..., IgnoreReadOnlyRecommended, ..., DoNotLoadFromEnterprise
            ok = app.FileOpenEx(full, True, e, e, e, e, True, e, e, e, e, PJ_POOL_READ_ONLY, e, e, True, e, True)
        except (pywintypes.com_error, TypeError):
            ok = app.FileOpenEx(full, True)
        if not ok:
            raise RuntimeError(f"MS Project no pudo abrir {full}.")
        self.project = app.ActiveProject
        self.opened_by_us = True
        self._opened_names = (self._open_project_names() or set()) - before

    def _open_project_names(self) -> set[str] | None:
        """Rutas (en minúsculas) de los proyectos abiertos en MS Project; None si no se pueden leer."""
        try:
            projects = self.app.Projects
            return {str(_get(projects.Item(i), "FullName", "")).lower() for i in range(1, int(projects.Count) + 1)}
        except Exception:
            return None

    def _activate(self) -> None:
        """Las llamadas a nivel de aplicación (cálculo, alias de campos) actúan sobre el proyecto activo."""
        try:
            self.project.Activate()
        except Exception:
            pass

    def _cleanup(self) -> None:
        try:
            if self.project is not None and self._status_changed and not self.opened_by_us:
                try:
                    self.project.StatusDate = self._orig_status if self._orig_status is not None else "NA"
                except Exception:
                    pass
            if self.project is not None and self.opened_by_us:
                try:
                    self.project.Activate()
                except Exception:
                    pass
                try:
                    self.app.FileCloseEx(PJ_DO_NOT_SAVE)
                except Exception:
                    try:
                        self.app.FileClose(PJ_DO_NOT_SAVE)
                    except Exception:
                        log.warning("No se pudo cerrar el proyecto en MS Project")
            if self.app is not None and not self.started_by_us:
                for prop, value in self._saved_app_props.items():
                    if value is not None:
                        try:
                            setattr(self.app, prop, value)
                        except Exception:
                            pass
                if self._prev_active is not None:
                    try:
                        self._prev_active.Activate()
                    except Exception:
                        pass
            if self.app is not None and self.started_by_us:
                # MS Project iniciado por el programa: se cierra si solo quedan abiertos proyectos que abrió el
                # programa (p. ej. el fondo de recursos que se abre junto con el archivo). Si el usuario abrió
                # otros proyectos en esa instancia (visible), se deja abierta para no descartar su trabajo.
                remaining = self._open_project_names()
                if remaining is not None and remaining <= self._opened_names:
                    # Se liberan antes las referencias a objetos de MS Project: liberarlas con el proceso ya
                    # cerrado produce errores de RPC (0x800706BA, "el servidor RPC no está disponible").
                    self._release_objects()
                    try:
                        self.app.Quit(PJ_DO_NOT_SAVE)
                    except Exception:
                        pass
        finally:
            self._release_objects()
            self._opened_names = set()
            self.app = None
            gc.collect()
            if self._pythoncom is not None:
                try:
                    self._pythoncom.CoUninitialize()
                except Exception:
                    pass
                self._pythoncom = None

    def _release_objects(self) -> None:
        """Libera las referencias a tareas y proyectos de MS Project (no a la aplicación)."""
        self.project = None
        self._prev_active = None
        gc.collect()

    # -- campos ---------------------------------------------------------------
    def _resolve_field(self, name: str):
        """('attr', 'Text10') o ('fid', id) para un nombre de campo o su alias."""
        if not name:
            return None
        b = builtin_kind(name)
        if b:
            return ("attr", b)
        parsed = parse_custom_name(name)
        if parsed:
            return ("attr", f"{parsed[0]}{parsed[1]}")
        for kind, table in (("Text", TASK_TEXT_IDS), ("Number", TASK_NUMBER_IDS)):
            for idx, fid in table.items():
                try:
                    alias = self.app.CustomFieldGetName(fid)
                except Exception:
                    continue
                if alias and same_name(str(alias), name):
                    return ("attr", f"{kind}{idx}")
        if is_spanish_wbs_name(name):
            return ("attr", "WBS")
        try:
            return ("fid", int(self.app.FieldNameToFieldConstant(name, 0)))
        except Exception:
            return None

    @staticmethod
    def _field_value(t, ref) -> object:
        if ref is None:
            return None
        kind, key = ref
        try:
            return getattr(t, key) if kind == "attr" else t.GetField(key)
        except Exception:
            return None

    # -- lectura --------------------------------------------------------------
    def set_status_date(self, d: date) -> None:
        self._activate()
        p = self.project
        self._orig_status = _get(p, "StatusDate")
        p.StatusDate = _com_date(datetime(d.year, d.month, d.day, 23, 59))
        self._status_changed = True
        for method in ("CalculateProject", "CalculateAll"):
            try:
                getattr(self.app, method)()
                break
            except Exception:
                continue

    def read_project(self, cfg: Config, status_date: date | None = None) -> ProjectData:
        p = self.project
        self._activate()
        if status_date is not None:
            self.set_status_date(status_date)
        warnings: list[str] = []
        # Línea base: la indicada en la configuración o la que el proyecto usa para el valor acumulado
        ev_baseline = int(_f(_get(p, "EarnedValueBaseline", 0)))
        n = cfg.baseline if cfg.baseline is not None else (ev_baseline if 0 <= ev_baseline <= 10 else 0)
        sfx = "" if n == 0 else str(n)

        wbs_ref = self._resolve_field(cfg.wbs_field)
        if wbs_ref is None:
            warnings.append(f"No se encontró el campo EDT '{cfg.wbs_field}' en MS Project; se usa el campo EDT (WBS).")
            wbs_ref = ("attr", "WBS")
        found: dict[str, str] = {}
        method_ref = self._resolve_field(cfg.method_field) if cfg.method_field else None
        if cfg.method_field and method_ref is None:
            warnings.append(f"No se encontró el campo de método '{cfg.method_field}' en MS Project.")
        elif method_ref is not None:
            found["metodo"] = cfg.method_field
        weight_ref = self._resolve_field(cfg.milestone_weight_field) if cfg.milestone_weight_field else None
        if cfg.milestone_weight_field and weight_ref is None:
            warnings.append(f"No se encontró el campo de peso de hitos '{cfg.milestone_weight_field}'.")
        elif weight_ref is not None:
            found["peso_hito"] = cfg.milestone_weight_field
        visible_ref = self._resolve_field(cfg.visible_field) if cfg.visible_field else None
        if cfg.visible_field_explicit and visible_ref is None:
            warnings.append(f"No se encontró el campo Visible '{cfg.visible_field}' en MS Project: las filas del "
                            "libro siguen el nivel máximo.")
        company_ref = split_ref = None
        if cfg.multi_enabled:
            company_ref = self._resolve_field(cfg.company_field)
            if company_ref is None:
                warnings.append(f"No se encontró el campo de empresa '{cfg.company_field}' en MS Project: todas las "
                                f"tareas se tratan como '{cfg.company_both}'.")
            else:
                found["campo_empresa"] = cfg.company_field
            split_ref = self._resolve_field(cfg.split_field)
            if split_ref is None:
                warnings.append(f"No se encontró el campo de reparto '{cfg.split_field}' en MS Project: las tareas "
                                f"'{cfg.company_both}' se reparten al {cfg.split_default:g} %.")

        tasks: list[Task] = []
        any_baseline = False
        skipped_external = 0
        for t in p.Tasks:
            if t is None:                 # fila en blanco
                continue
            if bool(_get(t, "ExternalTask", False)):
                skipped_external += 1
                continue
            uid = int(t.UniqueID)
            bl_start = _dt(_get(t, f"Baseline{sfx}Start"))
            bl_finish = _dt(_get(t, f"Baseline{sfx}Finish"))
            has_bl = bl_start is not None or bl_finish is not None
            any_baseline = any_baseline or has_bl
            wbs = self._field_value(t, wbs_ref)
            weight = self._field_value(t, weight_ref)
            company = self._field_value(t, company_ref)
            split = self._field_value(t, split_ref)
            visible = self._field_value(t, visible_ref)
            task = Task(
                uid=uid, id=int(_get(t, "ID", 0) or 0), name=str(_get(t, "Name", "") or ""),
                outline_level=int(_get(t, "OutlineLevel", 1) or 1),
                outline_number=str(_get(t, "OutlineNumber", "") or ""),
                wbs=str(wbs or "").strip(),          # sin código EDT => la tarea no tiene fila propia
                summary=bool(_get(t, "Summary", False)), milestone=bool(_get(t, "Milestone", False)),
                active=bool(_get(t, "Active", True)),
                baseline_start=bl_start, baseline_finish=bl_finish, has_baseline=has_bl,
                start=_dt(_get(t, "Start")), finish=_dt(_get(t, "Finish")),
                percent_complete=_f(_get(t, "PercentComplete")),
                physical_percent_complete=_f(_get(t, "PhysicalPercentComplete")),
                actual_start=_dt(_get(t, "ActualStart")), actual_finish=_dt(_get(t, "ActualFinish")),
                duration=_number(_get(t, "Duration")), estimated=bool(_get(t, "Estimated", False)),
                manual=bool(_get(t, "Manual", False)),
                constraint_type=int(_number(_get(t, "ConstraintType", CONSTRAINT_ASAP)) or CONSTRAINT_ASAP),
                constraint_date=_dt(_get(t, "ConstraintDate")), predecessors=_predecessors(t, uid),
                method=str(self._field_value(t, method_ref) or "").strip(),
                milestone_weight=None if weight in (None, "") else parse_number(weight),
                company=str(company or "").strip(),
                split_a=None if split in (None, "") else parse_number(split),
                visible="" if visible is None else str(visible).strip(),
            )
            tasks.append(task)
        if skipped_external:
            warnings.append(f"Se omitieron {skipped_external} tareas externas (vínculos entre proyectos).")

        saved = None
        for getter in (lambda: p.BaselineSavedDate(n), lambda: p.GetBaselineSavedDate(n)):
            try:
                saved = _dt(getter())
                break
            except Exception:
                continue
        any_baseline = any_baseline or saved is not None

        summary = _get(p, "ProjectSummaryTask")
        title = str(_get(summary, "Name", "") or "") if summary is not None else ""
        status = _dt(_get(p, "StatusDate"))
        project = ProjectData(
            name=title or str(_get(p, "Name", "") or Path(self.origen).stem),
            source=str(_get(p, "FullName", "") or self.origen), backend="com", tasks=tasks,
            status_date=status,
            currency_symbol=str(_get(p, "CurrencyCode", "") or _get(p, "CurrencySymbol", "") or ""),
            project_start=_dt(_get(p, "ProjectStart")),
            baseline_label="Línea base" if n == 0 else f"Línea base {n}",
            warnings=warnings, fields_found=found,
        )
        if not any_baseline:
            use_current_schedule(project)
        return project
