"""Fichas de tareas (comando generar-fichas): documento Markdown con una ficha por cada elemento EDT del informe.

Las fichas son las filas del libro (nivel máximo, planificación, campo Visible, tareas inactivas y varios archivos),
sin las filas «-D». Cada ficha tiene las fechas y la duración de la línea base (las del % planificado) y sus
predecesoras y sucesoras: cada vínculo de MS Project se traslada a los elementos con ficha que contienen sus dos
tareas. No se muestran los vínculos internos de un elemento, los repetidos ni los vínculos con tareas sin ficha (de
otra planificación, inactivas, con Visible = No o de otro proyecto).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from evm.config import Config
from evm.fields import fmt_number
from evm.model import ProjectData, Task
from evm.percent import calendar_for, task_dates
from evm.wbs import KIND_RESIDUAL, Report, ReportPart, ReportRow
from evm.workcalendar import WorkCalendar


@dataclass
class Card:
    """Ficha de un elemento EDT."""

    code: str
    name: str
    level: int                          # nivel de la fila en el libro (1 = elemento de nivel superior)
    start: datetime | None
    finish: datetime | None
    days: float | None                  # duración en días laborables (None = sin datos)
    milestone: bool = False
    current: bool = False               # la tarea no tiene línea base: fechas y duración del programa actual
    description: str = ""
    predecessors: list["Card"] = field(default_factory=list, repr=False)
    successors: list["Card"] = field(default_factory=list, repr=False)


@dataclass
class Cards:
    report: Report
    cards: list[Card]
    links: int = 0                      # vínculos entre fichas (pares predecesora-sucesora distintos)
    hidden: int = 0                     # vínculos de MS Project con una tarea sin ficha: no se muestran

    @property
    def current(self) -> int:
        """Fichas sin línea base."""
        return sum(1 for c in self.cards if c.current)


def _days(t: Task, project: ProjectData, calendar: WorkCalendar, start: datetime | None, finish: datetime | None,
          from_baseline: bool) -> float | None:
    """Duración en días laborables: la de MS Project (línea base o programa actual) dividida por los minutos de un día
    del proyecto; si el archivo no la tiene, el tiempo laborable del calendario entre las dos fechas."""
    minutes = t.baseline_duration if from_baseline else t.duration
    if minutes is None and not t.milestone and start is not None and finish is not None and finish > start:
        minutes = sum(calendar.minutes_by_month(start, finish).values())
    if minutes is None:
        return 0.0 if t.milestone else None
    return minutes / project.minutes_per_day


def _card(row: ReportRow, project: ProjectData, calendar: WorkCalendar) -> Card:
    t = row.task
    start, finish, from_baseline = task_dates(t)
    return Card(code=row.code, name=row.name, level=row.level, start=start, finish=finish,
                days=_days(t, project, calendar, start, finish, from_baseline), milestone=t.milestone,
                current=not from_baseline, description=t.description)


def build_cards(report: Report, parts: list[ReportPart], cfg: Config) -> Cards:
    """Fichas de las filas del informe (sin las filas «-D») con sus predecesoras y sucesoras."""
    projects = {s.name: s.project for s in report.sources}
    calendars: dict[str, WorkCalendar] = {}
    out = Cards(report, [])
    card_of: dict[tuple[str, int], Card] = {}          # (archivo, UniqueID) -> ficha de la tarea
    for row in report.rows:
        if row.kind == KIND_RESIDUAL:
            continue
        if row.source not in calendars:
            calendars[row.source] = calendar_for(projects[row.source], cfg)
        card = _card(row, projects[row.source], calendars[row.source])
        out.cards.append(card)
        card_of[(row.source, row.task.uid)] = card
    order = {id(c): i for i, c in enumerate(out.cards)}

    # Ficha de cada tarea exportada: la suya o la de su tarea superior más cercana con ficha. Las tareas que quedan
    # fuera del árbol (otra planificación, inactivas, Visible = No) o sin elemento superior con ficha no tienen
    owner: dict[tuple[str, int], Card] = {}
    for part in parts:
        name = part.source.name
        stack: list[tuple[Task, Card | None]] = [(t, None) for t in (part.scope.roots if part.scope else [])]
        while stack:
            t, current = stack.pop()
            current = card_of.get((name, t.uid), current)
            if current is not None:
                owner[(name, t.uid)] = current
            stack.extend((c, current) for c in t.children)

    pairs: set[tuple[int, int]] = set()
    for part in parts:
        name = part.source.name
        for t in part.source.project.tasks:
            for link in t.predecessors:
                succ, pred = owner.get((name, t.uid)), owner.get((name, link.uid))
                if succ is None and pred is None:
                    continue
                if succ is None or pred is None:
                    out.hidden += 1
                elif succ is not pred:
                    pairs.add((order[id(pred)], order[id(succ)]))
    for p, s in sorted(pairs):
        out.cards[s].predecessors.append(out.cards[p])
        out.cards[p].successors.append(out.cards[s])
    out.links = len(pairs)
    return out


# ----------------------------------------------------------------------------
# Markdown
# ----------------------------------------------------------------------------
_SPECIAL = re.compile(r"([\\`*_\[\]<>#|~])")
_LIST_NUMBER = re.compile(r"^(\d+)([.)])")


def _text(value: str) -> str:
    """Texto literal en Markdown, en una línea: sin énfasis, enlaces, HTML ni tablas."""
    return _SPECIAL.sub(r"\\\1", " ".join(str(value or "").split()))


def _paragraph(line: str) -> str:
    """Línea de la descripción como párrafo: tampoco puede empezar una lista ni un título."""
    text = _text(line)
    if text[:1] in ("-", "+", "="):
        return "\\" + text
    return _LIST_NUMBER.sub(r"\1\\\2", text)


def _date(d: datetime | None) -> str:
    return f"{d:%d/%m/%Y}" if d else "sin fecha"


def _duration(c: Card) -> str:
    if c.days is None:
        return "sin datos"
    v = round(c.days, 2)
    text = fmt_number(v, trim=True)
    return f"{text} {'día' if v == 1 else 'días'}" + (" (hito)" if c.milestone else "")


def render(cards: Cards, cfg: Config, today: date) -> str:
    """Documento Markdown: datos del informe y una ficha por elemento EDT, con un título del nivel de su fila."""
    rep = cards.report
    lines = [f"# Fichas de tareas: {_text(rep.title)}", ""]
    if rep.multi_file:
        lines.append("- Archivos de MS Project:")
        for s in rep.sources:
            plan = f"planificación {_text(s.plan)}; " if s.plan else ""
            lines.append(f"    - {_text(s.project.name)}: {plan}fechas planificadas: {s.project.baseline_label}")
    else:
        if rep.plan:
            lines.append(f"- Planificación: {_text(rep.plan)}")
        lines.append(f"- Fechas planificadas: {rep.project.baseline_label}")
    lines += [
        "- Duración: en días laborables",
        f"- Nivel de desglose EDT: {cfg.max_level}",
        f"- Fichas: {len(cards.cards)}",
        f"- Fecha de emisión: {today:%d/%m/%Y}",
        "",
    ]
    if not cards.cards:
        lines += ["Ninguna tarea tiene código EDT: no hay fichas.", ""]
    for c in cards.cards:
        lines += [f"{'#' * min(c.level + 1, 6)} {_text(c.code)} {_text(c.name)}", ""]
        for line in c.description.splitlines():
            if line.strip():
                lines += [_paragraph(line), ""]
        lines += [f"- Inicio planificado: {_date(c.start)}",
                  f"- Fin planificado: {_date(c.finish)}",
                  f"- Duración: {_duration(c)}"]
        for label, linked in (("Predecesoras", c.predecessors), ("Sucesoras", c.successors)):
            if linked:
                lines.append(f"- {label}:")
                lines += [f"    - {_text(x.code)} - {_text(x.name)}" for x in linked]
            else:
                lines.append(f"- {label}: ninguna")
        if c.current:
            lines.append("- Observación: la tarea no tiene línea base; las fechas y la duración son las del programa "
                         "actual.")
        lines.append("")
    return "\n".join(lines)


def write_cards(cards: Cards, cfg: Config, path: str | Path, today: date | None = None) -> None:
    Path(path).write_text(render(cards, cfg, today or date.today()), encoding="utf-8")
