"""generar-cliente con un MS Project simulado: copia del origen, borrado de tareas, vínculos, campo del cliente,
fechas fijadas, archivo de solo lectura y registro de diagnóstico.

El MS Project simulado hace lo mínimo: borra una tarea con su rama, guarda los vínculos de cada tarea, calcula el %
completado de los resúmenes con la duración de sus tareas de detalle (como MS Project) y, al recalcular, lleva al
comienzo del proyecto las tareas sin predecesoras, sin comenzar y sin delimitación (o a la fecha de su delimitación «No
comenzar antes del»). Las pruebas con MS Project real están en test_msproject_cliente.py.
"""
from __future__ import annotations

import os
import stat
from datetime import date, datetime, timedelta, timezone

import logging

import pytest

from evm.cli import main
from evm.client_plan import render_control
from evm.client_writer import PJ_CALENDARS, PJ_MODULES, _pred_uids, com_error_text, generate_client
from evm.config import Config
from evm.fields import TASK_TEXT_IDS
from evm.messages import EvmError
from evm.model import LINK_FF, LINK_FS, LINK_SF, LINK_SS
from evm.reader_com import ComSession

from compound import mpp_with_macros

UTC = timezone.utc
FIELD = TASK_TEXT_IDS[15]
# Borrado de las tareas que no se muestran (actividad_interna = false); ANON: «Otras tareas» (por defecto)
CFG = Config(plan="INTERNO", client_field="Etiqueta Cliente", client_internal_activity=False)
ANON = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="Text5")
PROJECT_START = datetime(2026, 1, 5, 8, tzinfo=UTC)
# Errores de MS Project simulados: citan el nombre de una tarea, que no debe llegar al registro ni a los mensajes
DELETE_ERROR = "No se puede borrar «Secreta»."
LINK_ERROR = "La tarea 'Secreta' crea un ciclo."


def d(day: int, hour: int = 8) -> datetime:
    return datetime(2026, 1, day, hour, tzinfo=UTC)


# Tipos de vínculo como los escribe MS Project en español en el texto de las predecesoras
LINK_TEXT = {LINK_FS: "FC", LINK_SS: "CC", LINK_FF: "FF", LINK_SF: "CF"}


class FakeList(list):
    """Colección o lista de MS Project: Count e Item(1..Count)."""

    @property
    def Count(self):
        return len(self)

    def Item(self, i):
        return self[i - 1]


class FakeCalendar:
    def __init__(self, project, name):
        self.project, self.Name = project, name

    def Delete(self):
        if self.Name in self.project.calendar_fails:
            raise RuntimeError(-2147352567, "Ocurrió una excepción.", (0, "MSProject", "No se puede.", None, 0, 0),
                               None)
        self.project.calendars.remove(self)


class FakeCode:
    def __init__(self, lines, declarations):
        self.CountOfLines, self.CountOfDeclarationLines = lines, declarations

    def DeleteLines(self, start, count):
        self.CountOfLines = 0


class FakeComponent:
    """Módulo de VBA: Type 1 = módulo estándar, 100 = ThisProject."""

    def __init__(self, name, kind=1, lines=10, declarations=1):
        self.Name, self.Type, self.CodeModule = name, kind, FakeCode(lines, declarations)


class FakeComponents(FakeList):
    def Remove(self, component):
        self.remove(component)


class FakeDep:
    def __init__(self, project, frm, to, kind=LINK_FS, lag=0):
        self.project, self.From, self.To, self.Type, self.Lag = project, frm, to, kind, lag

    def Delete(self):
        self.project.links.remove(self)


class FakeDeps:
    """TaskDependencies: los vínculos en que la tarea es predecesora o sucesora."""

    def __init__(self, project, task):
        self.project, self.task = project, task

    def __iter__(self):
        return iter([x for x in self.project.links if x.From is self.task or x.To is self.task])

    def Add(self, From, Type=LINK_FS, Lag=0):
        if (From.UniqueID, self.task.UniqueID) in self.project.reject:
            raise RuntimeError(-2147352567, "Ocurrió una excepción.", (0, "MSProject", LINK_ERROR, None, 0, 0), None)
        dep = FakeDep(self.project, From, self.task, Type, Lag)
        self.project.links.append(dep)
        return dep


class FakeTask:
    def __init__(self, project, uid, level, name, wbs="", mark="", start=d(5), finish=d(9, 17), duration=2400.0,
                 pct=0, actual_start=None):
        self.project = project
        self.UniqueID, self.ID, self.Name, self.OutlineLevel = uid, uid, name, level
        self.OutlineNumber = self.WBS = wbs
        self.Milestone, self.Active, self.ExternalTask = False, True, False
        self.Start, self.Finish, self.Duration = start, finish, duration
        self.BaselineStart, self.BaselineFinish, self.BaselineDuration = start, finish, duration
        self.Baseline1Start, self.Baseline1Finish = d(1), d(2, 17)
        self._pct, self.PhysicalPercentComplete = pct, 0
        self.ActualStart = actual_start or "NA"
        self.ActualFinish = "NA"
        self.ConstraintType, self.ConstraintDate = 0, "NA"
        self.Text15, self.Text5 = mark, wbs
        self.span = finish - start
        self.CalendarObject = None              # sin calendario propio
        self.pred_text = None                   # texto de las predecesoras que no se reconoce (p. ej. otro proyecto)

    @property
    def Summary(self):
        tasks = self.project.live
        i = tasks.index(self)
        return i + 1 < len(tasks) and tasks[i + 1].OutlineLevel > self.OutlineLevel

    @property
    def PercentComplete(self):
        """Resumen: duración real de sus tareas de detalle / su duración, redondeado (como MS Project)."""
        if not self.Summary:
            return self._pct
        leaves = self.project.leaves(self)
        total = sum(x.Duration for x in leaves)
        if not total:
            return 100 if leaves and all(x.PercentComplete >= 100 for x in leaves) else 0
        return round(sum(x.Duration * x.PercentComplete for x in leaves) / total)

    @PercentComplete.setter
    def PercentComplete(self, value):
        if self.project.reject_percent:
            raise RuntimeError("no se puede")
        self._pct = value

    @property
    def TaskDependencies(self):
        return FakeDeps(self.project, self)

    @property
    def UniqueIDPredecessors(self):
        """Como MS Project en español: separadas por punto y coma, con la posición en días y coma decimal."""
        if self.pred_text is not None:
            return self.pred_text
        parts = []
        for x in self.project.links:
            if x.To is not self:
                continue
            kind = LINK_TEXT[x.Type]
            if x.Lag:
                days = f"{abs(x.Lag) / 480:g}".replace(".", ",")
                parts.append(f"{x.From.UniqueID}{kind}{'+' if x.Lag > 0 else '-'}{days} días")
            else:
                parts.append(str(x.From.UniqueID) if kind == "FC" else f"{x.From.UniqueID}{kind}")
        return ";".join(parts)

    def Delete(self):
        self.project.delete(self)

    def GetField(self, fid):
        return self.Text15 if fid == FIELD else ""

    def SetField(self, fid, value):
        if fid == FIELD:
            self.Text15 = value


class FakeTasks:
    def __init__(self, project):
        self.project = project

    def __iter__(self):
        return iter(list(self.project.live))

    @property
    def Count(self):
        return len(self.project.live)


class FakeProject:
    def __init__(self):
        self.live: list[FakeTask] = []
        self.links: list[FakeDep] = []
        self.reject: set[tuple[int, int]] = set()
        self.AutoLinkTasks = True
        self.autolink_always = False            # simula que no se puede desactivar la vinculación automática
        self.fail_delete = False
        self.reject_percent = False
        self.StatusDate = "NA"
        self.ProjectStart = PROJECT_START
        self.FullName = self.Name = "copia.mpp"
        self.ProjectSummaryTask = type("S", (), {"Name": "Programa"})()
        self.calendars = [FakeCalendar(self, "Estándar")]
        self.calendar_fails: set[str] = set()   # calendarios que no se dejan quitar
        self.Resources = FakeList()
        self.vba: FakeComponents | None = None   # proyecto VBA (None: MS Project no da acceso a él)
        self.has_vba = False                    # el archivo guardado tiene un proyecto VBA (con sus módulos)

    @property
    def BaseCalendars(self):
        return FakeList(self.calendars)

    @property
    def Calendar(self):
        return self.calendars[0]

    @property
    def VBProject(self):
        if self.vba is None:
            raise RuntimeError(-2146827284, "No se confía en el acceso mediante programación al proyecto de Visual "
                                            "Basic.")
        return type("VBProject", (), {"VBComponents": self.vba})()

    def modules(self) -> list[str]:
        """Módulos (sin ThisProject) que quedarían al guardar."""
        return [c.Name for c in (self.vba or []) if c.Type != 100] + self.organizer_modules

    @property
    def Tasks(self):
        return FakeTasks(self)

    def add(self, *args, **kw) -> FakeTask:
        t = FakeTask(self, *args, **kw)
        self.live.append(t)
        return t

    def link(self, a, b, kind=LINK_FS, lag=0):
        self.links.append(FakeDep(self, a, b, kind, lag))

    def BaselineSavedDate(self, n):
        return "NA"

    def leaves(self, task):
        """Tareas de detalle de un resumen."""
        i = self.live.index(task)
        out = []
        for t in self.live[i + 1:]:
            if t.OutlineLevel <= task.OutlineLevel:
                break
            if not t.Summary:
                out.append(t)
        return out

    def delete(self, task):
        if self.fail_delete:
            raise RuntimeError(-2147352567, "Ocurrió una excepción.", (0, "MSProject", DELETE_ERROR, None, 0, 0), None)
        i = self.live.index(task)
        j = i + 1
        while j < len(self.live) and self.live[j].OutlineLevel > task.OutlineLevel:
            j += 1
        gone = self.live[i:j]
        preds = [x.From for x in self.links if x.To in gone and x.From not in gone]
        succs = [x.To for x in self.links if x.From in gone and x.To not in gone]
        self.links = [x for x in self.links if x.From not in gone and x.To not in gone]
        del self.live[i:j]
        if self.AutoLinkTasks or self.autolink_always:          # vinculación automática de MS Project
            for a in preds:
                for b in succs:
                    self.link(a, b)

    def calculate(self):
        has_pred = {x.To.UniqueID for x in self.links}
        for t in self.live:
            if t.Summary or t.ActualStart != "NA" or t.UniqueID in has_pred:
                continue
            if t.ConstraintType == 4:
                t.Start = t.ConstraintDate
            elif t.ConstraintType == 0:
                t.Start = PROJECT_START
            t.Finish = t.Start + t.span


class FakeProjects:
    def __init__(self, names):
        self.names = names

    @property
    def Count(self):
        return len(self.names)

    def Item(self, i):
        return type("P", (), {"FullName": self.names[i - 1]})()


class FakeApp:
    Version, Build = "16.0", "prueba"

    def __init__(self, project, source_open=""):
        self.project = project
        self.Calculation = -1
        self.names = {FIELD: "Etiqueta Cliente"}
        self.saved = 0
        self.rename_fails = False
        self.filters, self.applied, self.filter_fails = {}, [], False
        self.formulas = {}                      # id del campo -> fórmula
        self.organizer_fails = False
        self.organized = []                     # elementos quitados con el Organizador: (tipo, nombre)
        self.Projects = FakeProjects(["copia.mpp"] + ([source_open] if source_open else []))
        project.organizer_modules = []          # módulos que solo ve el Organizador (sin acceso al proyecto VBA)

    def OrganizerDeleteItem(self, Type, FileName, Name):
        if self.organizer_fails or FileName != self.project.Name:
            raise RuntimeError("no se puede")
        if Type == PJ_MODULES and Name in self.project.organizer_modules:
            self.project.organizer_modules.remove(Name)
        elif Type == PJ_CALENDARS and any(c.Name == Name for c in self.project.calendars):
            self.project.calendars = [c for c in self.project.calendars if c.Name != Name]
        else:
            raise RuntimeError("no existe")
        self.organized.append((Type, Name))

    def CustomFieldGetName(self, fid):
        return self.names.get(fid, "")

    def CustomFieldRename(self, fid, name):
        if self.rename_fails:
            raise RuntimeError("no se puede")
        self.names[fid] = name

    def FieldNameToFieldConstant(self, name, kind):
        raise RuntimeError("no existe")

    def CustomFieldGetFormula(self, fid):
        return self.formulas.get(fid, "")

    def FilterEdit(self, **kw):
        if self.filter_fails:
            raise RuntimeError("no se puede")
        self.filters[kw["Name"]] = kw

    def FilterApply(self, Name):
        self.applied.append(Name)

    def CalculateProject(self):
        self.project.calculate()

    def FileSave(self):
        self.saved += 1


def fake_session(app):
    class Session(ComSession):
        def __init__(self, origen, visible=False, read_only=True):
            super().__init__(origen, visible, read_only)
            self.app, self.project = app, app.project

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            pass

        def save(self):
            super().save()
            with open(self.origen, "wb") as f:          # lo que guarda MS Project en la copia
                p = app.project
                f.write(mpp_with_macros(p.modules()) if p.has_vba else b"archivo del cliente")

    return Session


def build(source_open=""):
    """INTERNO (con una tarea oculta en medio de una cadena, un resumen sin tareas visibles y una tarea que solo
    dependía de una oculta) y CLIENTE, con un vínculo entre las dos planificaciones."""
    p = FakeProject()
    t = {}
    t["root"] = p.add(1, 1, "INTERNO", mark="No")
    t["fase"] = p.add(2, 2, "Fase", "1", "Sí", start=d(5), finish=d(13, 17))
    t["a1"] = p.add(3, 3, "A1", "1.1", "Sí", start=d(5), finish=d(6, 17), duration=960)
    t["oculta"] = p.add(4, 3, "Oculta", "1.2", "No", start=d(7), finish=d(9, 17), duration=1440)
    t["a3"] = p.add(5, 3, "A3", "1.3", "si", start=d(12), finish=d(13, 17), duration=960)
    t["res"] = p.add(6, 2, "Resumen", "2", "S", start=d(5), finish=d(8, 17), duration=1920)
    t["r1"] = p.add(7, 3, "R1", "2.1", "", start=d(5), finish=d(6, 17), duration=960)
    t["r2"] = p.add(8, 3, "R2", "2.2", "No", start=d(7), finish=d(8, 17), duration=960)
    t["b"] = p.add(9, 2, "B", "3", "Sí", start=d(9), finish=d(9, 17), duration=480)
    t["sola"] = p.add(10, 2, "Solo oculta", "4", "No", start=d(7), finish=d(7, 17), duration=480)
    t["sola"].ConstraintType, t["sola"].ConstraintDate = 4, d(7)      # sin predecesoras: comienza el 7 por su delimitación
    t["c"] = p.add(11, 2, "C", "5", "Sí", start=d(8), finish=d(8, 17), duration=480)
    t["cliente"] = p.add(12, 1, "CLIENTE", mark="Sí")
    t["copia"] = p.add(13, 2, "Copia", "1", "Sí", start=d(5), finish=d(6, 17), duration=960)
    p.link(t["a1"], t["oculta"])
    p.link(t["oculta"], t["a3"])
    p.link(t["r1"], t["r2"])
    p.link(t["r2"], t["b"])
    p.link(t["sola"], t["c"])                 # C solo depende de una tarea oculta: sin ella, se adelantaría
    p.link(t["a1"], t["copia"], LINK_SS)      # vínculo entre planificaciones
    return FakeApp(p, source_open), t


def run(tmp_path, app, source_bytes=b"planificacion interna", **kw):
    source = tmp_path / "interno.mpp"
    source.write_bytes(source_bytes)
    destino = tmp_path / "cliente.mpp"
    plan, result = generate_client(source, destino, kw.pop("cfg", CFG), session_cls=fake_session(app), **kw)
    return source, destino, plan, result


def test_generates_the_client_file_from_a_copy(tmp_path):
    app, t = build()
    source, destino, plan, result = run(tmp_path, app)
    p = app.project
    assert source.read_bytes() == b"planificacion interna"                      # el origen no se toca
    assert destino.read_bytes() == b"archivo del cliente" and app.saved == 1
    assert not os.access(destino, os.W_OK)                                      # de solo lectura
    assert [x.name for x in tmp_path.iterdir() if x.name.startswith("~")] == []  # sin la copia temporal
    assert [x.UniqueID for x in p.live] == [1, 2, 3, 5, 6, 9, 11]
    assert t["root"].Name == "CLIENTE"
    assert all(x.Text15 == "" for x in p.live) and app.names[FIELD] == "" and result.field_name_removed
    links = sorted((x.From.UniqueID, x.To.UniqueID, x.Type, x.Lag) for x in p.links)
    assert links == [(3, 5, LINK_FS, 1440), (6, 9, LINK_FS, 0)]               # puente y vínculo del resumen
    assert result.removed_links == []
    # C quedó sin predecesoras: MS Project la llevaba al comienzo del proyecto; se fija su fecha
    assert [x.uid for x in result.fixed] == [11] and t["c"].ConstraintType == 4
    assert t["c"].Start == d(8) and result.differences == []
    assert (app.Calculation, p.AutoLinkTasks) == (-1, True)                    # opciones restauradas
    # Sin las tareas ocultas, el libro EVM del cliente pierde sus filas: se avisa
    assert {i.id for i in result.issues(plan)} == {"CLI-013", "CLI-024"}
    assert result.comparison.generated and result.comparison.book.only_internal


def test_inactive_task_that_moves_is_fixed(tmp_path):
    """MS Project sigue programando una tarea inactiva con sus predecesoras: sin la oculta de la que depende, se
    adelanta, y se fija como las demás."""
    app, t = build()
    t["c"].Active = False
    _, _, plan, result = run(tmp_path, app)
    assert 11 in plan.shown
    assert [x.uid for x in result.fixed] == [11] and t["c"].ConstraintType == 4
    assert t["c"].Start == d(8) and result.differences == []


def test_single_plan_without_plan_task(tmp_path):
    """Una sola planificación, sin tarea de planificación: [exportacion] planificacion vacía, todo el proyecto."""
    p = FakeProject()
    fase = p.add(1, 1, "Fase", "1", "Sí", start=d(5), finish=d(13, 17))
    a1 = p.add(2, 2, "A1", "1.1", "Sí", start=d(5), finish=d(6, 17), duration=960)
    oculta = p.add(3, 2, "Oculta", "1.2", "No", start=d(7), finish=d(9, 17), duration=1440)
    a3 = p.add(4, 2, "A3", "1.3", "Sí", start=d(12), finish=d(13, 17), duration=960)
    p.add(5, 1, "Interna", "2", "", start=d(5), finish=d(6, 17))
    p.add(6, 2, "Interna 1", "2.1", "Sí", start=d(5), finish=d(6, 17))     # Sí debajo de una tarea sin Sí
    p.add(7, 1, "B", "3", "S", start=d(5), finish=d(5, 17), duration=480)
    p.link(a1, oculta)
    p.link(oculta, a3)
    app = FakeApp(p)
    _, destino, plan, result = run(tmp_path, app, cfg=Config(client_field="Etiqueta Cliente", client_internal_activity=False))
    assert plan.root is None and plan.outside == []
    assert [x.UniqueID for x in p.live] == [1, 2, 4, 7] and fase.Name == "Fase"
    assert [(x.From.UniqueID, x.To.UniqueID, x.Lag) for x in p.links] == [(2, 4, 1440)]
    assert all(x.Text15 == "" for x in p.live) and destino.read_bytes() == b"archivo del cliente"
    assert {i.id for i in plan.issues()} >= {"CLI-004"} and result.differences == []


def test_collapsed_summary_keeps_its_baselines_and_progress(tmp_path):
    app, t = build()
    t["r1"].PercentComplete, t["r1"].ActualStart, t["res"].ActualStart = 100, d(5), d(5)
    assert t["res"].PercentComplete == 50           # el resumen tiene el avance de sus dos subtareas ocultas
    original = (t["res"].Duration, t["res"].BaselineStart, t["res"].Baseline1Finish)

    def forget(task):                           # MS Project cambia los datos del resumen al quedarse sin subtareas
        FakeProject.delete(app.project, task)
        if t["res"] in app.project.live and not t["res"].Summary:
            t["res"].Duration, t["res"].BaselineStart, t["res"].Baseline1Finish = 480, "NA", d(3, 17)

    app.project.delete = forget
    _, _, plan, result = run(tmp_path, app)
    assert (t["res"].Duration, t["res"].BaselineStart, t["res"].Baseline1Finish) == original
    assert not t["res"].Summary and t["res"].PercentComplete == 50     # ahora es una tarea, con el mismo avance
    assert result.progress_tasks == [] and "CLI-021" not in {i.id for i in result.issues(plan)}


def test_summary_with_hidden_progress_is_reported(tmp_path):
    """El avance de una subtarea oculta no cuenta en el resumen del cliente: se avisa (CLI-020) con el % previsto y
    el que calcula MS Project."""
    app, t = build()
    t["oculta"].PercentComplete, t["oculta"].ActualStart = 100, d(7)
    t["a1"].PercentComplete, t["a1"].ActualStart = 100, d(5)
    t["fase"].ActualStart = d(5)
    # Fase: A1 (2 días, 100 %), Oculta (3 días, 100 %), A3 (2 días, 0 %) -> 5/7 = 71 %; sin la oculta, 2/4 = 50 %
    assert t["fase"].PercentComplete == 71
    _, _, plan, result = run(tmp_path, app)
    [change] = plan.progress_changes
    assert (change.task.uid, change.task.percent_complete, change.client, change.hidden) == (2, 71, 50.0, 1)
    assert result.summary_percent[2] == 50 and result.progress_tasks == []
    assert "CLI-020" in {i.id for i in plan.issues()}


def test_task_that_loses_its_progress_is_reported(tmp_path):
    app, t = build()
    t["r1"].PercentComplete, t["r1"].ActualStart, t["res"].ActualStart = 100, d(5), d(5)
    app.project.reject_percent = True               # MS Project no acepta el avance del resumen convertido
    _, _, plan, result = run(tmp_path, app)
    [change] = result.progress_tasks
    assert (change.task.uid, change.task.percent_complete, change.client) == (6, 50, 0)
    assert {"CLI-019", "CLI-021"} <= {i.id for i in result.issues(plan)}


def test_links_created_by_ms_project_are_removed(tmp_path):
    app, _ = build()
    app.project.autolink_always = True
    _, _, plan, result = run(tmp_path, app)
    assert result.removed_links == [(3, 5)] or (3, 5) in result.removed_links
    links = sorted((x.From.UniqueID, x.To.UniqueID, x.Lag) for x in app.project.links)
    assert links == [(3, 5, 1440), (6, 9, 0)]
    assert "CLI-016" in {i.id for i in result.issues(plan)}


def test_rejected_link_is_reported(tmp_path):
    app, _ = build()
    app.project.reject = {(3, 5)}
    _, _, plan, result = run(tmp_path, app)
    [(link, error)] = result.rejected
    assert (link.pred.uid, link.succ.uid) == (3, 5)
    assert "Secreta" not in error and "«…»" in error                    # sin el nombre de la tarea
    assert "CLI-015" in {i.id for i in result.issues(plan)}


def test_without_fixing_dates_they_are_reported(tmp_path):
    app, t = build()
    cfg = Config(plan="INTERNO", client_field="Etiqueta Cliente", client_fix_dates=False,
                 client_internal_activity=False)
    _, _, plan, result = run(tmp_path, app, cfg=cfg)
    assert result.fixed == [] and t["c"].ConstraintType == 0
    [change] = result.differences
    assert change.task.uid == 11 and change.start == datetime(2026, 1, 5, 8) and "fijar_fechas" in change.note
    assert "CLI-014" in {i.id for i in result.issues(plan)}


def test_failure_leaves_no_file_and_the_source_untouched(tmp_path):
    app, _ = build()
    app.project.fail_delete = True
    with pytest.raises(EvmError) as exc:
        run(tmp_path, app)
    assert exc.value.id == "CLI-018" and "borrado de tareas" in str(exc.value) and "Secreta" not in str(exc.value)
    assert sorted(x.name for x in tmp_path.iterdir()) == ["interno.mpp"]
    assert (tmp_path / "interno.mpp").read_bytes() == b"planificacion interna"


def test_existing_destination_needs_overwrite(tmp_path):
    app, _ = build()
    run(tmp_path, app)
    app, _ = build()
    with pytest.raises(EvmError) as exc:
        run(tmp_path, app)
    assert exc.value.id == "CFG-042"
    app, _ = build()
    run(tmp_path, app, overwrite=True)                  # reemplaza el archivo de solo lectura
    assert not os.access(tmp_path / "cliente.mpp", os.W_OK)


def test_source_open_in_ms_project_is_reported(tmp_path):
    source = tmp_path / "interno.mpp"
    app, _ = build(source_open=str(source.resolve()).lower())
    _, _, plan, result = run(tmp_path, app)
    assert result.source_open and "CLI-012" in {i.id for i in result.issues(plan)}


def test_field_name_that_cannot_be_removed(tmp_path):
    app, _ = build()
    app.rename_fails = True
    _, _, plan, result = run(tmp_path, app)
    assert not result.field_name_removed and all(x.Text15 == "" for x in app.project.live)
    assert "CLI-017" in {i.id for i in result.issues(plan)}


# ----------------------------------------------------------------------------
# actividad_interna (por defecto): las tareas que no se muestran quedan como «Otras tareas»
# ----------------------------------------------------------------------------
def test_internal_activity_keeps_hidden_tasks_as_other_tasks(tmp_path):
    app, t = build()
    t["oculta"].Notes, t["oculta"].Text3, t["oculta"].Text10 = "Detalle interno", "Responsable", "0/100"
    t["oculta"].Hyperlink = "Oferta"
    t["oculta"]._pct, t["oculta"].ActualStart = 100, d(7)
    t["a1"].Notes = "Nota visible"
    cfg = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="Text5", method_field="Text10")
    _, destino, plan, result = run(tmp_path, app, cfg=cfg)
    p = app.project
    assert [x.UniqueID for x in p.live] == list(range(1, 12))          # sin la planificación CLIENTE
    hidden = {4, 7, 8, 10}
    assert {x.UniqueID for x in p.live if x.Name == "Otras tareas"} == hidden
    assert all(x.Text5 == "" for x in p.live if x.UniqueID in hidden)
    assert (t["a1"].Name, t["a1"].Text5, t["a1"].Notes) == ("A1", "1.1", "Nota visible")    # las que se muestran
    o = t["oculta"]
    assert (o.Notes, o.Text3, o.Hyperlink, o.Text10, o.PercentComplete) == ("", "", "", "0/100", 100)
    assert t["root"].Name == "CLIENTE" and all(x.Text15 == "" for x in p.live)
    # Se conservan todos los vínculos de la planificación (el que iba a la otra planificación desaparece con ella)
    links = sorted((x.From.UniqueID, x.To.UniqueID) for x in p.links)
    assert links == [(3, 4), (4, 5), (7, 8), (8, 9), (10, 11)] and result.removed_links == []
    assert result.fixed == [] and result.anonymized == 4 and result.codes_removed and result.not_cleared == {}
    comp = result.comparison
    assert comp.generated and comp.percent_internal == comp.percent_client and comp.tasks == []
    assert comp.checked == 11
    ids = {i.id for i in plan.issues() + result.issues(plan)}
    assert "CLI-022" in ids and not ids & {"CLI-023", "CLI-026", "CLI-027", "CLI-028"}
    # Libro EVM (se avisa, CLI-024 y CLI-025): 1.2 pasa a la fila 1.0.0-D; el resumen 2 pierde los códigos de sus dos
    # subtareas y su fila reúne las de 2.1 y 2.2; 4 (de primer nivel) queda sin código y fuera del libro del cliente
    book = comp.book
    assert [(a.code, b.code) for a, b in book.moved] == [("1.2.0", "1.0.0-D")]
    assert book.where == {"2.1.0": ["2.0.0"], "2.2.0": ["2.0.0"], "4.0.0": [], "2.0.0": ["2.1.0", "2.2.0"]}
    assert (book.leaves_internal, book.leaves_client) == (8, 7)
    assert "CLI-024" in ids and "CLI-025" in ids and destino.read_bytes() == b"archivo del cliente"


def test_internal_activity_removes_inactive_hidden_tasks(tmp_path):
    app, t = build()
    t["sola"].Active = False
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert [x.uid for x in plan.dropped] == [10] and 10 not in {x.UniqueID for x in app.project.live}
    assert result.anonymized == 3 and result.comparison.tasks == []


def test_internal_activity_reports_tasks_that_change(tmp_path):
    """Si una tarea del archivo del cliente no conserva sus fechas, la comparación final lo indica (CLI-026)."""
    app, t = build()
    calculate = app.project.calculate

    def late():
        calculate()
        t["b"].Start = d(12)

    app.project.calculate = late
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    [diff] = result.comparison.tasks
    assert diff.task.uid == 9 and diff.fields[0][0] == "Comienzo"
    assert "CLI-026" in {i.id for i in result.issues(plan)}


def _formula_task(t):
    class Formula(FakeTask):
        @property
        def Text7(self):
            return "calculado"

        @Text7.setter
        def Text7(self, value):
            raise RuntimeError(-2147352567, "Ocurrió una excepción.", (0, "MSProject", "Campo calculado", None, 0, 0),
                               None)

    t["oculta"].__class__ = Formula


def test_internal_activity_field_that_cannot_be_cleared(tmp_path):
    app, t = build()
    _formula_task(t)
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert result.not_cleared["Text7"][0] == 1 and "Campo calculado" in result.not_cleared["Text7"][1]
    assert t["oculta"].Name == "Otras tareas"
    assert "Text7 (1)" in {i.id: i.message for i in result.issues(plan)}["CLI-027"]
    assert result.calculated == {}


def test_internal_activity_calculated_field(tmp_path):
    """Un campo con fórmula no se puede vaciar: se informa aparte (CLI-030), no como dato que no se quitó."""
    app, t = build()
    _formula_task(t)
    app.formulas[TASK_TEXT_IDS[7]] = '[Duration]/480 & " días"'
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert result.not_cleared == {} and result.calculated == {"Text7": 1}
    issues = {i.id: i.message for i in result.issues(plan)}
    assert "CLI-027" not in issues and "Text7 (1)" in issues["CLI-030"]
    report = render_control(plan, "origen.mpp", date(2026, 11, 8), "Prueba", outcome=result)
    assert "Campos calculados de las «Otras tareas»" in report and "| Text7 | 1 |" in report


def test_comparison_error_still_delivers_the_file(tmp_path, monkeypatch):
    import evm.client_writer as writer

    def broken(*args, **kw):
        raise ZeroDivisionError("prueba")

    monkeypatch.setattr(writer, "compare", broken)
    app, _ = build()
    _, destino, plan, result = run(tmp_path, app, cfg=ANON)
    assert destino.read_bytes() == b"archivo del cliente"
    assert "ZeroDivisionError" in result.comparison.book.error
    assert "CLI-029" in {i.id for i in result.issues(plan)}


def test_view_filter(tmp_path):
    app, _ = build()
    cfg = Config(plan="INTERNO", client_field="Etiqueta Cliente", wbs_field="Text5", client_filter=True)
    _, _, plan, result = run(tmp_path, app, cfg=cfg)
    assert result.filter_applied and app.filters["Tareas del cliente"]["Value"] == "Otras tareas"
    assert app.applied == ["Tareas del cliente"]
    app, _ = build()
    app.filter_fails = True
    _, _, plan, result = run(tmp_path, app, cfg=cfg, overwrite=True)
    assert result.filter_applied is False and "CLI-028" in {i.id for i in result.issues(plan)}


def test_cli_generates_other_tasks(tmp_path, monkeypatch, capsys):
    import evm.client_writer as writer
    app, _ = build()
    for x in app.project.live:
        if x.Name not in ("INTERNO", "CLIENTE"):
            x.Name = f"Confidencial{x.UniqueID}"
    monkeypatch.setattr(writer, "ComSession", fake_session(app))
    source = tmp_path / "interno.mpp"
    source.write_bytes(b"planificacion interna")
    cfg = tmp_path / "c.toml"
    cfg.write_text('[campos]\nedt = "Text5"\n\n[exportacion]\nplanificacion = "INTERNO"\n\n[cliente]\n'
                   'campo = "Etiqueta Cliente"\n', encoding="utf-8")
    log = tmp_path / "registro.log"
    assert main(["--registro", str(log), "generar-cliente", "--origen", str(source), "--config", str(cfg)]) == 0
    out = capsys.readouterr().out
    assert "Tareas que no se muestran: 4 como «Otras tareas», 0 inactivas se quitan" in out
    assert "% completado del proyecto: 0 % (interno) y 0 % (cliente)" in out and "Tareas comparadas: 11" in out
    report = (tmp_path / "interno_cliente_control.md").read_text(encoding="utf-8")
    assert "Archivo del cliente leído en MS Project después de guardarlo." in report
    assert "| Tareas que quedan como «Otras tareas» | 4 |" in report
    text = log.read_text(encoding="utf-8")
    assert "Paso: tareas internas" in text and "Comparación: % completado del proyecto" in text
    assert "Confidencial" not in text and "Confidencial" not in report


def test_com_error_text_hides_quoted_names():
    exc = RuntimeError(-2147352567, "Ocurrió una excepción.", (0, "MSProject", "La tarea 'Montaje' y «Pruebas» "
                                                               "forman un ciclo.", None, 0, 0), None)
    assert com_error_text(exc) == "0x80020009 La tarea «…» y «…» forman un ciclo."


def test_cli_generates_with_log_file(tmp_path, monkeypatch, capsys):
    import evm.client_writer as writer
    app, _ = build()
    for x in app.project.live:
        if x.Name not in ("INTERNO", "CLIENTE"):
            x.Name = f"Confidencial{x.UniqueID}"
    monkeypatch.setattr(writer, "ComSession", fake_session(app))
    source = tmp_path / "interno.mpp"
    source.write_bytes(b"planificacion interna")
    cfg = tmp_path / "c.toml"
    cfg.write_text('[exportacion]\nplanificacion = "INTERNO"\n\n[cliente]\ncampo = "Etiqueta Cliente"\nactividad_interna = false\n',
                   encoding="utf-8")
    log = tmp_path / "registro.log"
    assert main(["--registro", str(log), "generar-cliente", "--origen", str(source), "--config", str(cfg)]) == 0
    out = capsys.readouterr()
    assert "Archivo del cliente generado" in out.out and "1 fijadas" in out.out
    assert "DEBUG" not in out.err                       # sin -v, la consola no muestra el diagnóstico
    destino = tmp_path / "interno_cliente.mpp"
    assert destino.read_bytes() == b"archivo del cliente"
    report = (tmp_path / "interno_cliente_control.md").read_text(encoding="utf-8")
    assert "## Escritura con MS Project" in report and "### Fechas fijadas" in report
    text = log.read_text(encoding="utf-8")
    assert "Paso: vínculos" in text and "Vínculo nuevo 1.1 (UID 3) -> 1.3 (UID 5)" in text
    assert "MS Project 16.0" in text and "Fecha fijada 5 (UID 11)" in text
    assert "Confidencial" not in text and "Confidencial" not in report
    # Con XML solo se puede simular
    assert main(["generar-cliente", "--origen", str(tmp_path / "x.xml"), "--config", str(cfg)]) == 2
    assert "CLI-010" in capsys.readouterr().err
    # El destino no puede ser el origen
    assert main(["generar-cliente", "--origen", str(source), "--config", str(cfg), "--destino", str(source)]) == 2
    assert "CLI-011" in capsys.readouterr().err


def test_ms_project_error_is_logged_without_names(tmp_path, monkeypatch, capsys):
    import evm.client_writer as writer
    app, _ = build()
    app.project.fail_delete = True
    monkeypatch.setattr(writer, "ComSession", fake_session(app))
    source = tmp_path / "interno.mpp"
    source.write_bytes(b"planificacion interna")
    cfg = tmp_path / "c.toml"
    cfg.write_text('[exportacion]\nplanificacion = "INTERNO"\n\n[cliente]\ncampo = "Etiqueta Cliente"\nactividad_interna = false\n',
                   encoding="utf-8")
    log = tmp_path / "registro.log"
    assert main(["--registro", str(log), "generar-cliente", "--origen", str(source), "--config", str(cfg)]) == 2
    err = capsys.readouterr().err
    assert "CLI-018" in err and "borrado de tareas" in err and "Secreta" not in err
    text = log.read_text(encoding="utf-8")
    assert "Error en el paso borrado de tareas" in text and "client_writer.py" in text and "Secreta" not in text
    assert sorted(x.name for x in tmp_path.iterdir()) == ["c.toml", "interno.mpp", "registro.log"]


def test_log_file_records_error_details(tmp_path, capsys):
    log = tmp_path / "registro.log"
    assert main(["--registro", str(log), "generar-cliente", "--origen", str(tmp_path / "no.mpp"), "--simular"]) == 2
    assert "CFG-054" in capsys.readouterr().err
    text = log.read_text(encoding="utf-8")
    assert "Exportador EVM" in text and "Argumentos:" in text and "Traceback" in text


# ----------------------------------------------------------------------------
# Vínculos: se leen de una vez (Id. único de predecesoras)
# ----------------------------------------------------------------------------
@pytest.mark.parametrize("text, uids", [
    ("", set()), ("12", {12}), ("12;15FC+2,5 días;7CC-1 sem.", {12, 15, 7}), ("12,15FF+50%", {12, 15}),
    ("3FC+2 días?", {3}), ("4CF-3 edías;9", {4, 9}), ("12FS+2.5 days,15SS", {12, 15}), (" 8 ; 9 ", {8, 9}),
])
def test_predecessor_text(text, uids):
    assert _pred_uids(text) == uids


@pytest.mark.parametrize("text", [r"C:\Planes\otro.mpp\12", "12FC+2,15", "Doce", None, 12])
def test_predecessor_text_not_recognized(text):
    """Un vínculo con otro proyecto o un texto que no se entiende: se recorren los vínculos de la tarea."""
    assert _pred_uids(text) is None


def test_links_are_read_from_the_predecessor_text(tmp_path, caplog):
    caplog.set_level(logging.INFO, logger="evm.client_writer")
    app, t = build()
    t["a3"].pred_text = r"C:\Planes\otro.mpp\4"            # no se reconoce: se recorren sus vínculos
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert "tareas revisadas vínculo a vínculo: 1 de 11" in caplog.text
    assert result.removed_links == [] and len(app.project.links) == 5


def test_link_check_finds_the_links_ms_project_adds(tmp_path, caplog):
    """Con «Otras tareas» solo se borran las inactivas: si MS Project vincula sus vecinas, se quita el vínculo."""
    caplog.set_level(logging.INFO, logger="evm.client_writer")
    app, t = build()
    t["sola"].Active = False
    app.project.autolink_always = True
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert result.removed_links == [] and "tareas revisadas vínculo a vínculo: 0 de 10" in caplog.text
    app, t = build()
    t["oculta"].Active = False
    app.project.autolink_always = True
    _, _, plan, result = run(tmp_path, app, cfg=ANON, overwrite=True)
    assert result.removed_links == [(3, 5)] and "tareas revisadas vínculo a vínculo: 1 de 10" in caplog.text
    assert "CLI-016" in {i.id for i in result.issues(plan)}


# ----------------------------------------------------------------------------
# Contenido del archivo: calendarios sin uso y macros
# ----------------------------------------------------------------------------
def with_calendars(app, t):
    p = app.project
    for name in ("Sede A", "Sede B", "Sede C", "Sede D", "Sede O", "Sede R",
                 "Para calend. de línea base de Microsoft Project 98"):
        p.calendars.append(FakeCalendar(p, name))
    cal = {c.Name: c for c in p.calendars}
    t["c"].CalendarObject = cal["Sede A"]                   # tarea que se muestra
    t["oculta"].CalendarObject = cal["Sede O"]              # «Otras tareas» (con actividad_interna) o se borra
    p.Resources = FakeList([None, type("R", (), {"BaseCalendar": "Sede R"})()])
    p.calendar_fails = {"Sede C", "Sede D"}                 # no se dejan quitar con Delete
    return p


def test_unused_calendars_are_removed(tmp_path):
    app, t = build()
    p = with_calendars(app, t)
    original = app.OrganizerDeleteItem

    def organizer(Type, FileName, Name):                    # el Organizador quita Sede C, pero no Sede D
        if Name == "Sede D":
            raise RuntimeError("no se puede")
        original(Type, FileName, Name)

    app.OrganizerDeleteItem = organizer
    p.Title, p.Author = "Programa", "Responsable"
    p.ViewList = FakeList(["Diagrama de Gantt", "Tareas del cliente"])
    p.TaskFilterList = FakeList(["Todas las tareas"])
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    c = result.content
    assert c.checked
    assert c.calendars == [("Estándar", "calendario del proyecto"), ("Sede A", "1 tarea"), ("Sede O", "1 tarea"),
                           ("Sede R", "1 recurso"),
                           ("Para calend. de línea base de Microsoft Project 98",
                            "de MS Project (líneas base de Project 98)")]
    assert c.removed_calendars == ["Sede B", "Sede C"] and [n for n, _ in c.failed_calendars] == ["Sede D"]
    assert [x.Name for x in p.calendars] == ["Estándar", "Sede A", "Sede D", "Sede O", "Sede R",
                                             "Para calend. de línea base de Microsoft Project 98"]
    assert c.properties == ["Título", "Autor"]
    assert c.lists == {"Vistas": ["Diagrama de Gantt", "Tareas del cliente"], "Tablas de tareas": [],
                       "Filtros de tareas": ["Todas las tareas"], "Grupos de tareas": []}
    issues = {i.id: i.message for i in result.issues(plan)}
    assert "2 calendarios sin uso" in issues["CLI-032"] and "1 calendario sin uso" in issues["CLI-033"]
    report = render_control(plan, "origen.mpp", date(2026, 11, 8), "Prueba", outcome=result)
    assert "### Contenido del archivo del cliente" in report and "| Sede B | Sin uso: quitado |" in report
    assert "| Calendarios | 5 se conservan, 2 sin uso quitados, 1 sin uso que no se pudieron quitar |" in report
    assert "| Propiedades del archivo con valor | Título, Autor (se ven en Archivo > Información) |" in report
    assert "| Vistas | 2: Diagrama de Gantt, Tareas del cliente |" in report
    assert "Programa" not in report.split("### Contenido del archivo del cliente")[1]     # sin los valores


def test_calendar_of_deleted_tasks_is_removed(tmp_path):
    """actividad_interna = false: el calendario que solo usaba una tarea que se borra queda sin uso."""
    app, t = build()
    p = with_calendars(app, t)
    p.calendar_fails = set()
    _, _, plan, result = run(tmp_path, app)
    assert "Sede O" in result.content.removed_calendars
    assert ("Sede A", "1 tarea") in result.content.calendars


def test_calendars_are_kept_without_the_project_calendar(tmp_path):
    """Sin el calendario del proyecto no se sabe qué calendarios se usan: no se quita ninguno."""
    app, t = build()
    p = with_calendars(app, t)
    p.calendars[0].Name = ""
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    assert result.content.removed_calendars == [] and len(p.calendars) == 8
    assert all("no se comprobó" in use for name, use in result.content.calendars if name.startswith("Sede B"))


def test_macros_are_removed_with_the_vba_project(tmp_path):
    app, _ = build()
    p = app.project
    p.vba = FakeComponents([FakeComponent("ThisProject", 100, lines=3, declarations=0), FakeComponent("Módulo1"),
                            FakeComponent("Clase1", 2)])
    p.has_vba = True
    _, destino, plan, result = run(tmp_path, app, source_bytes=mpp_with_macros(["Módulo1"], ["Clase1"]), cfg=ANON)
    c = result.content
    assert c.macros == ["Módulo1", "Clase1"] and c.macro_method == "proyecto VBA"
    assert c.removed_macros == ["código de ThisProject", "Módulo1", "Clase1"]
    assert c.macros_left == [] and c.remaining_macros == [] and not os.access(destino, os.W_OK)
    assert p.vba[0].CodeModule.CountOfLines == 0                    # ThisProject se queda, sin código
    issues = {i.id: i.message for i in result.issues(plan)}
    assert "las macros (código de ThisProject, Módulo1, Clase1)" in issues["CLI-032"] and "CLI-033" not in issues
    report = render_control(plan, "origen.mpp", date(2026, 11, 8), "Prueba", outcome=result)
    assert ("| Macros | Quitadas (proyecto VBA): código de ThisProject, Módulo1, Clase1; comprobado en el archivo "
            "guardado |") in report


def test_macros_are_removed_with_the_organizer(tmp_path):
    """Sin acceso al proyecto VBA (Centro de confianza de MS Project), se quitan con el Organizador por el nombre de
    los módulos leídos del archivo."""
    app, _ = build()
    p = app.project
    p.organizer_modules, p.has_vba = ["Módulo1"], True
    _, _, plan, result = run(tmp_path, app, source_bytes=mpp_with_macros(["Módulo1"]), cfg=ANON)
    c = result.content
    assert c.macro_method == "Organizador" and c.removed_macros == ["Módulo1"] and c.macros_left == []
    assert app.organized == [(PJ_MODULES, "Módulo1")] and "No se confía" in c.macro_error
    assert "CLI-032" in {i.id for i in result.issues(plan)}


def test_macros_that_remain_are_reported(tmp_path):
    app, _ = build()
    p = app.project
    p.organizer_modules, p.has_vba = ["Módulo1"], True
    app.organizer_fails = True
    _, _, plan, result = run(tmp_path, app, source_bytes=mpp_with_macros(["Módulo1"]), cfg=ANON)
    assert result.content.macros_left == ["Módulo1"]
    issues = {i.id: i.message for i in result.issues(plan)}
    assert "macros (Módulo1)" in issues["CLI-033"] and "CLI-032" not in issues
    report = render_control(plan, "origen.mpp", date(2026, 11, 8), "Prueba", outcome=result)
    assert "| Macros | El archivo del cliente las conserva: Módulo1 |" in report


def test_macros_that_cannot_be_checked(tmp_path):
    """Ni el proyecto VBA ni el archivo se pueden leer (en las demás pruebas, el archivo simulado no es un .mpp)."""
    app, _ = build()
    _, _, plan, result = run(tmp_path, app, cfg=ANON)
    c = result.content
    assert c.macros is None and c.macros_left is None and c.remaining_macros == []
    assert not {"CLI-032", "CLI-033"} & {i.id for i in result.issues(plan)}
    report = render_control(plan, "origen.mpp", date(2026, 11, 8), "Prueba", outcome=result)
    assert "| Macros | No se pudo comprobar (0x" in report
