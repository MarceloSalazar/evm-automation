"""Lector de archivos XML de MS Project (Archivo > Guardar como > XML).

No requiere MS Project instalado. Formato: esquema MSPDI
(http://schemas.microsoft.com/project). En este formato los montos de costo
se guardan multiplicados por 100 (centavos).
"""
from __future__ import annotations

import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

from evm.config import Config
from evm.fields import builtin_kind, field_id, is_spanish_wbs_name, parse_custom_name, parse_number, same_name
from evm.model import MonthKey, ProjectData, Task, use_current_schedule
from evm.timeline import split_segment
from evm.workcalendar import WorkCalendar, calendar_from_xml

log = logging.getLogger(__name__)

# Tipo de dato escalonado "costo de línea base de la tarea" según la línea base:
# 0 -> 10; N (1..10) -> 19 + 6 * (N - 1)
def timephased_baseline_cost_type(baseline: int) -> int:
    return 10 if baseline == 0 else 19 + 6 * (baseline - 1)


# FixedCostAccrual en MSPDI: 1 = comienzo, 2 = fin, 3 = prorrateado
_ACCRUAL = {"1": "inicio", "2": "fin", "3": "prorrateado"}


class XmlFormatError(ValueError):
    pass


class _Reader:
    def __init__(self, root: ET.Element):
        self.root = root
        tag = root.tag
        self.ns = tag[: tag.index("}") + 1] if tag.startswith("{") else ""
        if not tag.endswith("Project"):
            raise XmlFormatError("El archivo no es un XML de MS Project (falta el elemento <Project>).")

    def q(self, name: str) -> str:
        return self.ns + name

    def text(self, el: ET.Element | None, name: str, default: str = "") -> str:
        if el is None:
            return default
        child = el.find(self.q(name))
        if child is None or child.text is None:
            return default
        return child.text.strip()

    def date(self, el, name) -> datetime | None:
        t = self.text(el, name)
        if not t:
            return None
        try:
            d = datetime.fromisoformat(t)
        except ValueError:
            return None
        return d.replace(tzinfo=None) if d.year >= 1984 else None

    def num(self, el, name, default: float = 0.0) -> float:
        t = self.text(el, name)
        if not t:
            return default
        try:
            return float(t)
        except ValueError:
            return default

    def money(self, el, name) -> float:
        """Montos de costo: en MSPDI se guardan multiplicados por 100."""
        return self.num(el, name) / 100.0

    def money_opt(self, el, name) -> float | None:
        t = self.text(el, name)
        if not t:
            return None
        try:
            return float(t) / 100.0
        except ValueError:
            return None

    def flag(self, el, name, default: bool = False) -> bool:
        t = self.text(el, name)
        if not t:
            return default
        return t in ("1", "true", "True")


def _resolve(name: str, defs: dict[int, tuple[str, str]]) -> tuple[str, object] | None:
    """Devuelve ('elem', 'WBS'|'OutlineNumber') o ('ext', FieldID) para el nombre configurado."""
    if not name:
        return None
    b = builtin_kind(name)
    if b:
        return ("elem", b)
    for fid, (fname, alias) in defs.items():
        if (alias and same_name(alias, name)) or same_name(fname, name):
            return ("ext", fid)
    parsed = parse_custom_name(name)
    if parsed:
        fid = field_id(*parsed)
        if fid:
            return ("ext", fid)
    if is_spanish_wbs_name(name):
        return ("elem", "WBS")
    return None


def read_xml(path: str | Path, cfg: Config) -> ProjectData:
    path = Path(path)
    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        raise XmlFormatError(f"El archivo XML no es válido: {exc}") from exc
    r = _Reader(tree.getroot())
    root = r.root
    warnings: list[str] = []

    title = ""
    currency = r.text(root, "CurrencyCode") or r.text(root, "CurrencySymbol")
    status = r.date(root, "StatusDate")

    # Definiciones de campos personalizados: FieldID -> (FieldName, Alias)
    defs: dict[int, tuple[str, str]] = {}
    ext_parent = root.find(r.q("ExtendedAttributes"))
    if ext_parent is not None:
        for ea in ext_parent.findall(r.q("ExtendedAttribute")):
            try:
                fid = int(r.text(ea, "FieldID"))
            except ValueError:
                continue
            defs[fid] = (r.text(ea, "FieldName"), r.text(ea, "Alias"))

    wbs_ref = _resolve(cfg.wbs_field, defs) or ("elem", "WBS")
    if cfg.wbs_field and _resolve(cfg.wbs_field, defs) is None:
        warnings.append(f"No se encontró el campo EDT '{cfg.wbs_field}' en el XML; se usa el campo EDT (WBS).")
    method_ref = _resolve(cfg.method_field, defs) if cfg.method_field else None
    if cfg.method_field and method_ref is None:
        warnings.append(f"No se encontró el campo de método '{cfg.method_field}' en el XML.")
    weight_ref = _resolve(cfg.milestone_weight_field, defs) if cfg.milestone_weight_field else None
    if cfg.milestone_weight_field and weight_ref is None:
        warnings.append(f"No se encontró el campo de peso de hitos '{cfg.milestone_weight_field}' en el XML.")
    company_ref = split_ref = None
    if cfg.multi_enabled:
        company_ref = _resolve(cfg.company_field, defs)
        if company_ref is None:
            warnings.append(f"No se encontró el campo de empresa '{cfg.company_field}' en el XML: todas las tareas "
                            f"se tratan como '{cfg.company_both}'.")
        split_ref = _resolve(cfg.split_field, defs)
        if split_ref is None:
            warnings.append(f"No se encontró el campo de reparto '{cfg.split_field}' en el XML: las tareas "
                            f"'{cfg.company_both}' se reparten al {cfg.split_default:g} %.")

    # Calendario del proyecto: reparte los tramos escalonados que cruzan de un mes a otro
    try:
        cal_uid = int(r.text(root, "CalendarUID"))
    except ValueError:
        cal_uid = None
    calendar = calendar_from_xml(r, root, cal_uid) or WorkCalendar()
    calendar.add_holidays(cfg.holidays)

    # Línea base: la indicada en la configuración o la usada por el proyecto para el valor acumulado
    try:
        ev_baseline = int(r.text(root, "BaselineForEarnedValue") or 0)
    except ValueError:
        ev_baseline = 0
    ev_baseline = ev_baseline if 0 <= ev_baseline <= 10 else 0
    n = cfg.baseline if cfg.baseline is not None else ev_baseline
    if n != ev_baseline:
        warnings.append(f"El PV y el BAC usan la línea base {n}, pero MS Project calcula su valor acumulado (CPTR) "
                        f"con la línea base {ev_baseline}: el método 'CPTR de MS Project' no es comparable.")
    tp_type = str(timephased_baseline_cost_type(n))
    tasks_el = root.find(r.q("Tasks"))
    if tasks_el is None:
        raise XmlFormatError("El XML no contiene tareas (<Tasks>).")

    tasks: list[Task] = []
    summary_bl_cost = None
    summary_cost = None
    summary_acwp = None
    any_baseline = False
    skipped_external = 0

    for el in tasks_el.findall(r.q("Task")):
        if r.flag(el, "IsNull"):
            continue
        try:
            uid = int(r.text(el, "UID"))
            tid = int(r.text(el, "ID", "0") or 0)
            level = int(r.text(el, "OutlineLevel", "0") or 0)
        except ValueError:
            continue

        ext: dict[int, str] = {}
        for ea in el.findall(r.q("ExtendedAttribute")):
            try:
                ext[int(r.text(ea, "FieldID"))] = r.text(ea, "Value")
            except ValueError:
                continue

        def value_of(ref) -> str:
            if ref is None:
                return ""
            kind, key = ref
            return r.text(el, key) if kind == "elem" else ext.get(key, "")

        bl = None
        ev_bl_cost = None
        for b in el.findall(r.q("Baseline")):
            if r.text(b, "Number") == str(n):
                bl = b
            if n != ev_baseline and r.text(b, "Number") == str(ev_baseline):
                ev_bl_cost = r.money(b, "Cost")
        if n != ev_baseline and ev_bl_cost is None:
            ev_bl_cost = 0.0
        bl_start = r.date(bl, "Start") if bl is not None else None
        bl_finish = r.date(bl, "Finish") if bl is not None else None
        bl_cost = r.money(bl, "Cost") if bl is not None else 0.0
        has_bl = bl_start is not None or bl_finish is not None
        any_baseline = any_baseline or has_bl or bool(bl_cost)

        if level == 0:
            summary_bl_cost = bl_cost
            summary_cost = r.money(el, "Cost")
            summary_acwp = r.money_opt(el, "ACWP")
            if summary_acwp is None:
                summary_acwp = r.money_opt(el, "ActualCost")
            # Mismo nombre que entrega COM (ProjectSummaryTask.Name), para reconocer el proyecto
            title = r.text(el, "Name")
            continue
        if r.flag(el, "ExternalTask"):
            skipped_external += 1
            continue

        # Costo de línea base mensual (datos escalonados en el tiempo)
        tp: dict[MonthKey, float] | None = None
        for seg in el.findall(r.q("TimephasedData")):
            if r.text(seg, "Type") != tp_type:
                continue
            s, f = r.date(seg, "Start"), r.date(seg, "Finish")
            amount = r.money(seg, "Value")
            if s is None or f is None or not amount:
                continue
            tp = tp or {}
            for k, v in split_segment(amount, s, f, calendar).items():
                tp[k] = tp.get(k, 0.0) + v

        weight = None
        if weight_ref is not None:
            weight = parse_number(value_of(weight_ref))
        company = value_of(company_ref).strip() if company_ref is not None else ""
        split_a = parse_number(value_of(split_ref)) if split_ref is not None else None

        task = Task(
            uid=uid, id=tid, name=r.text(el, "Name"), outline_level=level,
            outline_number=r.text(el, "OutlineNumber"),
            wbs=value_of(wbs_ref).strip(),         # sin código EDT => la tarea no tiene fila propia
            summary=r.flag(el, "Summary"), milestone=r.flag(el, "Milestone"),
            active=r.flag(el, "Active", True),
            baseline_cost=bl_cost, baseline_start=bl_start, baseline_finish=bl_finish, has_baseline=has_bl,
            start=r.date(el, "Start"), finish=r.date(el, "Finish"), cost=r.money(el, "Cost"),
            fixed_cost=r.money(el, "FixedCost"),
            fixed_cost_accrual=_ACCRUAL.get(r.text(el, "FixedCostAccrual"), "prorrateado"),
            percent_complete=r.num(el, "PercentComplete"),
            physical_percent_complete=r.num(el, "PhysicalPercentComplete"),
            actual_start=r.date(el, "ActualStart"), actual_finish=r.date(el, "ActualFinish"),
            actual_cost=r.money(el, "ActualCost"), acwp=r.money_opt(el, "ACWP"), bcwp=r.money_opt(el, "BCWP"),
            method=value_of(method_ref).strip(), milestone_weight=weight, ev_baseline_cost=ev_bl_cost,
            company=company, split_a=split_a,
            timephased_pv=tp,
        )
        tasks.append(task)

    tasks.sort(key=lambda t: t.id)
    if not title:
        name = r.text(root, "Name")
        title = r.text(root, "Title") or (Path(name).stem if name else path.stem)
    if skipped_external:
        warnings.append(f"Se omitieron {skipped_external} tareas externas (vínculos entre proyectos).")

    label = "Línea base" if n == 0 else f"Línea base {n}"
    project = ProjectData(
        name=title, source=str(path), backend="xml", tasks=tasks, status_date=status,
        currency_symbol=currency, project_start=r.date(root, "StartDate"), baseline_label=label,
        baseline_found=any_baseline, summary_baseline_cost=summary_bl_cost, summary_cost=summary_cost,
        summary_acwp=summary_acwp, ev_baseline=ev_baseline,
        warnings=warnings, calendar=calendar,
    )
    if not any_baseline:
        use_current_schedule(project)
    return project
