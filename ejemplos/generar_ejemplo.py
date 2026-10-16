"""Genera un proyecto de ejemplo en formato XML de MS Project (MSPDI).

Sirve para probar el exportador sin MS Project. El archivo imita lo que escribe
MS Project al usar Archivo > Guardar como > XML: costos multiplicados por 100,
línea base 0, costo de línea base mensualizado (TimephasedData tipo 10),
campos personalizados con alias ("Método EVM" en Texto10 y "Peso hito" en
Número1) y, opcionalmente, avance a una fecha de estado.

Uso:
    python ejemplos/generar_ejemplo.py                       # solo línea base
    python ejemplos/generar_ejemplo.py --fecha-estado 2027-06-30 --salida ejemplos/avance_2027-06.xml
"""
from __future__ import annotations

import argparse
import calendar
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

TEXT10 = 188743750      # Texto10 (tarea)
NUMBER1 = 188743767     # Número1 (tarea)


@dataclass
class T:
    name: str
    level: int
    start: date | None = None
    finish: date | None = None
    cost: float = 0.0
    method: str = ""
    milestone: bool = False
    weight: float | None = None
    children: list["T"] = field(default_factory=list)
    # calculados
    uid: int = 0
    wbs: str = ""
    pct: float = 0.0
    actual_cost: float = 0.0
    actual_start: date | None = None
    actual_finish: date | None = None


def wdays(a: date, b: date) -> list[date]:
    out, d = [], a
    while d <= b:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


def add_wd(d: date, n: int) -> date:
    while d.weekday() >= 5:
        d += timedelta(days=1)
    k = 1
    while k < n:
        d += timedelta(days=1)
        if d.weekday() < 5:
            k += 1
    return d


def build_program() -> list[T]:
    """Programa de 4 elementos de nivel 1, hasta nivel 4, ~5 años."""
    start = date(2026, 1, 5)

    def wp(name, offset_m, dur_d, cost, method="", acts=None):
        s = add_wd(start + timedelta(days=int(offset_m * 30.4)), 1)
        t = T(name, 3, method=method)
        if acts:
            cur = s
            for (aname, adur, acost, *rest) in acts:
                if adur == 0:  # hito
                    t.children.append(T(aname, 4, cur, cur, acost, milestone=True, weight=rest[0] if rest else None))
                else:
                    f = add_wd(cur, adur)
                    t.children.append(T(aname, 4, cur, f, acost))
                    cur = add_wd(f + timedelta(days=1), 1)
        else:
            t.start, t.finish, t.cost = s, add_wd(s, dur_d), cost
        return t

    prog = [
        T("Gestión del programa", 1, children=[
            T("Dirección y control", 2, children=[
                wp("Oficina de gestión del programa (PMO)", 0, 1250, 1_800_000, "Avance lineal (LOE)"),
                wp("Control de proyectos y EVM", 0, 1250, 650_000, "Avance lineal (LOE)"),
                wp("Gestión de riesgos", 1, 1200, 300_000, "Avance lineal (LOE)"),
            ]),
            T("Gestión de contratos", 2, children=[
                wp("Licitación de ingeniería", 0, 80, 120_000, "0/100"),
                wp("Licitación de obras", 8, 100, 180_000, "50/50"),
                wp("Administración de contratos", 10, 1000, 420_000, "Avance lineal (LOE)"),
            ]),
        ]),
        T("Ingeniería", 1, children=[
            T("Ingeniería básica", 2, children=[
                wp("Estudios de terreno", 2, 60, 350_000, "50/50"),
                wp("Diseño conceptual", 3, 0, 0, "Hitos ponderados", acts=[
                    ("Bases de diseño", 25, 180_000), ("Hito: bases aprobadas", 0, 0, 30),
                    ("Alternativas y selección", 40, 260_000), ("Hito: alternativa seleccionada", 0, 0, 30),
                    ("Informe conceptual", 30, 160_000), ("Hito: informe aprobado", 0, 0, 40),
                ]),
                wp("Ingeniería básica civil", 6, 140, 900_000, "% Completado"),
                wp("Ingeniería básica eléctrica", 6, 140, 750_000, "% Completado"),
            ]),
            T("Ingeniería de detalle", 2, children=[
                wp("Detalle civil", 12, 220, 1_600_000, "% Físico completado"),
                wp("Detalle mecánico", 13, 220, 1_450_000, "% Físico completado"),
                wp("Detalle eléctrico y control", 14, 240, 1_300_000, "% Completado"),
                wp("Revisión y aprobación", 22, 0, 0, "Hitos ponderados", acts=[
                    ("Revisión 60 %", 30, 120_000), ("Hito: revisión 60 % aprobada", 0, 0, 40),
                    ("Revisión 90 %", 30, 120_000), ("Hito: emitido para construcción", 0, 0, 60),
                ]),
            ]),
        ]),
        T("Construcción", 1, children=[
            T("Obras civiles", 2, children=[
                wp("Movimiento de tierras", 18, 160, 3_200_000, "Unidades completadas"),
                wp("Fundaciones", 22, 220, 5_400_000, "Unidades completadas"),
                wp("Estructuras", 28, 300, 7_800_000, "% Físico completado"),
            ]),
            T("Montaje electromecánico", 2, children=[
                wp("Suministro de equipos", 20, 0, 0, "Hitos ponderados", acts=[
                    ("Fabricación de equipos", 200, 6_000_000), ("Hito: equipos en fábrica (FAT)", 0, 0, 40),
                    ("Transporte", 40, 900_000), ("Hito: equipos en obra", 0, 0, 60),
                ]),
                wp("Montaje mecánico", 32, 280, 4_300_000, "% Completado"),
                wp("Montaje eléctrico", 34, 300, 3_900_000, "% Completado"),
            ]),
        ]),
        T("Puesta en marcha", 1, children=[
            T("Pruebas", 2, children=[
                wp("Pruebas en frío", 45, 60, 650_000, "0/100"),
                wp("Pruebas en caliente", 48, 60, 850_000, "20/80"),
            ]),
            T("Cierre", 2, children=[
                wp("Capacitación y entrega", 50, 60, 400_000, "50/50"),
                wp("Cierre administrativo", 52, 60, 150_000, ""),
            ]),
        ]),
    ]
    return prog


def walk(nodes, prefix=""):
    for i, n in enumerate(nodes, start=1):
        n.wbs = f"{prefix}{i}"
        yield n
        yield from walk(n.children, n.wbs + ".")


def rollup(n: T) -> None:
    if not n.children:
        return
    for c in n.children:
        rollup(c)
    n.start = min(c.start for c in n.children)
    n.finish = max(c.finish for c in n.children)
    n.cost = sum(c.cost for c in n.children)


def monthly(n: T) -> dict[tuple[int, int], float]:
    """Costo de línea base por mes (lineal por días hábiles), como lo hace MS Project."""
    if n.children:
        out: dict[tuple[int, int], float] = {}
        for c in n.children:
            for k, v in monthly(c).items():
                out[k] = out.get(k, 0.0) + v
        return out
    if not n.cost:
        return {}
    days = wdays(n.start, n.finish) or [n.start]
    out = {}
    for d in days:
        out[(d.year, d.month)] = out.get((d.year, d.month), 0.0) + n.cost / len(days)
    return out


def progress(n: T, status: date) -> None:
    """Avance simulado: atraso del 15 % y sobrecosto del 8 %."""
    if n.children:
        for c in n.children:
            progress(c, status)
        done_cost = sum(c.cost * c.pct for c in n.children)
        tot = sum(c.cost for c in n.children)
        dur = sum(max(1, len(wdays(c.start, c.finish))) * c.pct for c in n.children)
        dtot = sum(max(1, len(wdays(c.start, c.finish))) for c in n.children)
        n.pct = (dur / dtot) if dtot else 0.0
        n.actual_cost = sum(c.actual_cost for c in n.children)
        starts = [c.actual_start for c in n.children if c.actual_start]
        n.actual_start = min(starts) if starts else None
        n.actual_finish = max(c.actual_finish for c in n.children) if all(c.actual_finish for c in n.children) else None
        _ = done_cost, tot
        return
    if status < n.start:
        return
    total = max(1, len(wdays(n.start, n.finish)))
    elapsed = len(wdays(n.start, min(status, n.finish)))
    pct = min(1.0, elapsed / total * 0.85)
    if n.milestone:
        pct = 1.0 if status >= n.finish + timedelta(days=20) else 0.0
    n.pct = round(pct, 2)
    n.actual_start = n.start if n.pct > 0 or n.milestone and n.pct else None
    n.actual_finish = n.finish + timedelta(days=20) if n.pct >= 1.0 else None
    n.actual_cost = round(n.cost * n.pct * 1.08, 2)


def dt(d: date, end: bool = False) -> str:
    return datetime.combine(d, time(17, 0) if end else time(8, 0)).isoformat()


def money(v: float) -> str:
    """Montos en MSPDI: multiplicados por 100."""
    text = f"{v * 100:.2f}"
    return text[:-3] if text.endswith(".00") else text


def write(nodes: list[T], path: Path, status: date | None) -> None:
    tasks = list(walk(nodes))
    for r in nodes:
        rollup(r)
    for uid, t in enumerate(tasks, start=1):
        t.uid = uid
    if status:
        for r in nodes:
            progress(r, status)
    start = min(t.start for t in tasks)
    finish = max(t.finish for t in tasks)
    total = sum(r.cost for r in nodes)
    total_ac = sum(r.actual_cost for r in nodes)
    total_ev = sum(t.cost * t.pct for t in tasks if not t.children)

    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<Project xmlns="http://schemas.microsoft.com/project">',
           "\t<SaveVersion>14</SaveVersion>",
           "\t<Name>programa_ejemplo.xml</Name>",
           "\t<Title>Programa de ejemplo EVM</Title>",
           f"\t<StartDate>{dt(start)}</StartDate>",
           f"\t<FinishDate>{dt(finish, True)}</FinishDate>",
           "\t<CurrencyDigits>2</CurrencyDigits>",
           "\t<CurrencySymbol>$</CurrencySymbol>",
           "\t<CurrencyCode>USD</CurrencyCode>",
           "\t<DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>",
           "\t<BaselineForEarnedValue>0</BaselineForEarnedValue>"]
    if status:
        out.append(f"\t<StatusDate>{dt(status, True)}</StatusDate>")
    out += ["\t<DefaultTaskEVMethod>0</DefaultTaskEVMethod>",
            "\t<ExtendedAttributes>",
            f"\t\t<ExtendedAttribute><FieldID>{TEXT10}</FieldID><FieldName>Text10</FieldName>"
            "<Alias>Método EVM</Alias></ExtendedAttribute>",
            f"\t\t<ExtendedAttribute><FieldID>{NUMBER1}</FieldID><FieldName>Number1</FieldName>"
            "<Alias>Peso hito</Alias></ExtendedAttribute>",
            "\t</ExtendedAttributes>",
            "\t<Tasks>"]

    def task_xml(uid, tid, name, level, wbs, t: T | None):
        summary = bool(t and t.children)
        s = t.start if t else start
        f = t.finish if t else finish
        cost = t.cost if t else total
        ac = t.actual_cost if t else total_ac
        pct = (t.pct if t else (total_ev / total if total else 0))
        ev = (sum(x.cost * x.pct for x in walk_leaves(t)) if t else total_ev)
        lines = ["\t\t<Task>", f"\t\t\t<UID>{uid}</UID>", f"\t\t\t<ID>{tid}</ID>",
                 f"\t\t\t<Name>{escape(name)}</Name>", "\t\t\t<Active>1</Active>", "\t\t\t<IsNull>0</IsNull>",
                 f"\t\t\t<WBS>{wbs}</WBS>", f"\t\t\t<OutlineNumber>{wbs}</OutlineNumber>",
                 f"\t\t\t<OutlineLevel>{level}</OutlineLevel>",
                 f"\t\t\t<Start>{dt(s)}</Start>", f"\t\t\t<Finish>{dt(f, True)}</Finish>",
                 f"\t\t\t<Milestone>{1 if t and t.milestone else 0}</Milestone>",
                 f"\t\t\t<Summary>{1 if summary or t is None else 0}</Summary>",
                 "\t\t\t<ExternalTask>0</ExternalTask>",
                 "\t\t\t<FixedCost>0</FixedCost>", "\t\t\t<FixedCostAccrual>3</FixedCostAccrual>",
                 f"\t\t\t<PercentComplete>{int(round(pct * 100))}</PercentComplete>",
                 f"\t\t\t<Cost>{money(cost)}</Cost>"]
        if t and t.actual_start:
            lines.append(f"\t\t\t<ActualStart>{dt(t.actual_start)}</ActualStart>")
        if t and t.actual_finish:
            lines.append(f"\t\t\t<ActualFinish>{dt(t.actual_finish, True)}</ActualFinish>")
        lines += [f"\t\t\t<ActualCost>{money(ac)}</ActualCost>",
                  f"\t\t\t<ACWP>{round(ac * 100, 2):.2f}</ACWP>",
                  f"\t\t\t<BCWP>{round(ev * 100, 2):.2f}</BCWP>",
                  f"\t\t\t<PhysicalPercentComplete>{int(round(pct * 100))}</PhysicalPercentComplete>",
                  "\t\t\t<EarnedValueMethod>0</EarnedValueMethod>"]
        if t and t.method:
            lines.append(f"\t\t\t<ExtendedAttribute><FieldID>{TEXT10}</FieldID><Value>{escape(t.method)}</Value>"
                         "</ExtendedAttribute>")
        if t and t.weight is not None:
            lines.append(f"\t\t\t<ExtendedAttribute><FieldID>{NUMBER1}</FieldID><Value>{t.weight:g}</Value>"
                         "</ExtendedAttribute>")
        lines += ["\t\t\t<Baseline>", "\t\t\t\t<Number>0</Number>",
                  f"\t\t\t\t<Start>{dt(s)}</Start>", f"\t\t\t\t<Finish>{dt(f, True)}</Finish>",
                  f"\t\t\t\t<Cost>{money(cost)}</Cost>", "\t\t\t</Baseline>"]
        months = monthly(t) if t else {}
        if t is None:
            for r in nodes:
                for k, v in monthly(r).items():
                    months[k] = months.get(k, 0.0) + v
        for (y, m), v in sorted(months.items()):
            a = max(s, date(y, m, 1))
            b = min(f, date(y, m, calendar.monthrange(y, m)[1]))
            lines += ["\t\t\t<TimephasedData>", "\t\t\t\t<Type>10</Type>", f"\t\t\t\t<UID>{uid}</UID>",
                      f"\t\t\t\t<Start>{dt(a)}</Start>", f"\t\t\t\t<Finish>{dt(b, True)}</Finish>",
                      "\t\t\t\t<Unit>1</Unit>", f"\t\t\t\t<Value>{money(v)}</Value>",
                      "\t\t\t</TimephasedData>"]
        lines.append("\t\t</Task>")
        return lines

    out += task_xml(0, 0, "Programa de ejemplo EVM", 0, "0", None)
    for i, t in enumerate(tasks, start=1):
        out += task_xml(t.uid, i, t.name, t.level, t.wbs, t)
    out += ["\t</Tasks>", "</Project>"]
    path.write_text("\n".join(out) + "\n", encoding="utf-8")


def walk_leaves(t: T):
    if not t.children:
        yield t
    for c in t.children:
        yield from walk_leaves(c)


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--salida", default=str(Path(__file__).with_name("programa_ejemplo.xml")))
    p.add_argument("--fecha-estado", help="AAAA-MM-DD: agrega avance simulado a esa fecha")
    a = p.parse_args()
    status = date.fromisoformat(a.fecha_estado) if a.fecha_estado else None
    write(build_program(), Path(a.salida), status)
    print(f"Archivo generado: {a.salida}")


if __name__ == "__main__":
    main()
