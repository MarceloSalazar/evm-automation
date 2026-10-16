"""Conservación de EV, AC y métodos de un libro anterior al regenerar (--migrar-desde)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from evm.model import MonthKey
from evm.wbs import KIND_LEAF, SEV_ERROR, SEV_INFO, SEV_WARN, Issue, Report, ReportRow, fmt_money, iter_subtree
from evm.workbook import SHEET_LISTS, row_key


@dataclass
class PriorRow:
    key: str                     # UniqueID (o UniqueID-D)
    code: str
    name: str
    uid: int | None
    residual: bool
    ev: dict[MonthKey, float]
    ac: dict[MonthKey, float]
    method: str
    origin_method: str | None = None     # método al generarse el libro (None: libro antiguo)

    @property
    def method_edited(self) -> bool:
        """True si el método de la columna C fue modificado en Excel."""
        if not self.method:
            return False
        if self.origin_method is not None:
            return self.method != self.origin_method
        return self.method.strip().lower() not in ("por definir", "a definir", "")

    @property
    def has_values(self) -> bool:
        return bool(self.ev or self.ac)


@dataclass
class PriorData:
    rows: list[PriorRow]
    status_date: date | None
    project_name: str
    max_level: int | None
    methods_list: list[str]
    same_project: bool = True
    # Resultado de plan_migration: clave de la fila nueva -> datos a escribir
    values: dict[str, dict[str, dict[MonthKey, float]]] = field(default_factory=dict)
    methods: dict[str, str] = field(default_factory=dict)
    matched: int = 0
    unmatched: list[PriorRow] = field(default_factory=list)

    def values_for(self, key: str, kind: str) -> dict[MonthKey, float]:
        return self.values.get(key, {}).get(kind, {})

    def method_for(self, key: str) -> str | None:
        return self.methods.get(key)


def prior_from_book(path, project_name: str) -> PriorData:
    """Lee EV, AC, métodos y fecha de corte de un libro generado anteriormente."""
    from evm.update import read_book
    book = read_book(path)
    rows = []
    for s in book.input_rows:
        uid_text = s.key[:-2] if s.key.endswith("-D") else s.key
        rows.append(PriorRow(
            key=s.key, code=s.code, name=s.name, uid=int(uid_text) if uid_text.isdigit() else None,
            residual=s.key.endswith("-D"),
            ev={book.horizon.key(i): v for i, v in enumerate(book.ev[s.row]) if v},
            ac={book.horizon.key(i): v for i, v in enumerate(book.ac[s.row]) if v},
            method=book.method[s.row], origin_method=s.origin_method))
    methods_list = []
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    if SHEET_LISTS in wb.sheetnames:
        for (value,) in wb[SHEET_LISTS].iter_rows(min_row=2, max_col=1, values_only=True):
            if not isinstance(value, str) or not value.strip():
                break                   # la lista termina en la primera fila vacía
            if len(value.strip()) <= 60:
                methods_list.append(value.strip())
    wb.close()
    level = book.meta.get("nivel_maximo")
    return PriorData(rows=rows, status_date=book.status_date,
                     project_name=str(book.meta.get("proyecto") or ""),
                     max_level=int(level) if level else None, methods_list=methods_list)


def plan_migration(report: Report, prior: PriorData, max_level: int) -> list[Issue]:
    """Asigna cada fila del libro anterior a una fila del libro nuevo (sin duplicar historiales).

    Mismo proyecto: por UniqueID; si la fila anterior quedó dentro de un paquete
    nuevo (p. ej. al bajar el nivel máximo), su historial se suma a ese paquete.
    Proyecto distinto: por código EDT, una fila anterior por fila nueva.
    """
    issues: list[Issue] = []
    prior.same_project = prior.project_name == report.project.name
    if not prior.same_project:
        issues.append(Issue(SEV_WARN, "", "",
                            f"El libro anterior corresponde al proyecto '{prior.project_name}' y el nuevo a "
                            f"'{report.project.name}': los EV/AC se asocian por código EDT."))
    if prior.max_level and prior.max_level != max_level:
        issues.append(Issue(SEV_WARN, "", "",
                            f"El libro anterior usaba nivel máximo {prior.max_level} y el nuevo {max_level}. Las "
                            "filas que ahora se dividen en varias no pueden repartir su historial (ver errores)."))

    inputs = [r for r in report.rows if r.is_input]
    by_key = {row_key(r): r for r in inputs}
    by_code: dict[str, ReportRow] = {}
    for r in inputs:
        by_code.setdefault(r.code, r)
    contains: dict[int, ReportRow] = {}
    for r in inputs:
        if r.kind == KIND_LEAF:
            for n in iter_subtree(r.task):
                contains.setdefault(n.uid, r)

    used_codes: set[str] = set()
    sources: dict[str, list[tuple[PriorRow, str]]] = {}
    for old in prior.rows:
        target, how = None, ""
        if prior.same_project:
            if old.key in by_key:
                target, how = by_key[old.key], "uid"
            elif old.uid is not None and old.uid in contains:
                target, how = contains[old.uid], "rama"
        else:
            cand = by_code.get(old.code)
            if cand is not None and row_key(cand) not in used_codes:
                target, how = cand, "codigo"
                used_codes.add(row_key(cand))
        if target is None:
            if old.has_values:
                prior.unmatched.append(old)
            continue
        sources.setdefault(row_key(target), []).append((old, how))

    horizon = report.horizon
    clamped = 0.0
    for key, lst in sources.items():
        target = by_key[key]
        merged: dict[str, dict[MonthKey, float]] = {"EV": {}, "AC": {}}
        for old, how in lst:
            for kind, data in (("EV", old.ev), ("AC", old.ac)):
                for k, v in data.items():
                    i = horizon.index(k)
                    if i >= horizon.months:
                        clamped += abs(v)
                    i = min(max(i, 0), horizon.months - 1)
                    nk = horizon.key(i)
                    merged[kind][nk] = merged[kind].get(nk, 0.0) + v
            if how == "rama" and old.has_values:
                issues.append(Issue(SEV_INFO, target.code, target.name,
                                    f"Se sumó el historial de EV/AC de la fila anterior {old.code} '{old.name}'."))
            if how == "codigo" and old.has_values and old.name != target.name:
                issues.append(Issue(SEV_WARN, target.code, target.name,
                                    f"Recibió el historial de EV/AC de la fila anterior {old.code} '{old.name}' "
                                    "(asociada por código EDT)."))
        if merged["EV"] or merged["AC"]:
            prior.values[key] = merged
            prior.matched += 1
        # El método editado en Excel en el libro anterior se conserva si la asociación es uno a uno.
        # Si el método cambió en MS Project (y no se editó en Excel), se usa el nuevo.
        if len(lst) == 1 and lst[0][1] in ("uid", "codigo"):
            old = lst[0][0]
            if old.method_edited and old.method != target.method:
                prior.methods[key] = old.method

    if prior.methods:
        legacy = any(o.origin_method is None for o in prior.rows)
        issues.append(Issue(SEV_INFO, "", "",
                            f"Se conservaron {len(prior.methods)} métodos de medición de la columna C del libro anterior"
                            + (" (el libro anterior no indica si se editaron en Excel o cambiaron en MS Project; "
                               "revíselos)." if legacy else ", editados en Excel.")))
        # Recalcular el aviso de filas sin método
        undefined = sum(1 for r in inputs
                        if (prior.methods.get(row_key(r)) or r.method).strip().lower() in ("", "por definir",
                                                                                          "a definir"))
        report.issues[:] = [i for i in report.issues if "sin método de medición" not in i.message]
        if undefined:
            issues.append(Issue(SEV_INFO, "", "", f"{undefined} filas sin método de medición definido (columna C)."))
    if clamped > 0.005:
        issues.append(Issue(SEV_WARN, "", "",
                            f"Hay EV/AC por {fmt_money(clamped)} cargados después del último mes del nuevo horizonte: "
                            "se acumularon en el último mes."))
    for old in prior.unmatched:
        issues.append(Issue(SEV_ERROR, old.code, old.name,
                            f"No se pudo trasladar su historial (EV {fmt_money(sum(old.ev.values()))}, AC "
                            f"{fmt_money(sum(old.ac.values()))}): la tarea no existe en el proyecto nuevo o se dividió "
                            "en varias filas. Cárguelo a mano en las filas que correspondan."))
    return issues
