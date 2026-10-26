"""Modelo de datos común a los lectores (COM y XML)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

from evm.messages import Issue

# Clave de mes: (año, mes)
MonthKey = tuple[int, int]

# Tipos de vínculo (PjTaskLinkType de VBA = <Type> de <PredecessorLink> en el XML)
LINK_FF, LINK_FS, LINK_SF, LINK_SS = 0, 1, 2, 3

# Delimitaciones (PjConstraint de VBA = <ConstraintType> en el XML)
CONSTRAINT_ASAP = 0             # Lo antes posible
CONSTRAINT_MSO = 2              # Debe comenzar el
CONSTRAINT_MFO = 3              # Debe finalizar el
CONSTRAINT_SNLT = 5             # No comenzar después del
CONSTRAINT_FNLT = 7             # No finalizar después del


@dataclass
class Link:
    """Vínculo con una tarea predecesora."""

    uid: int                            # UniqueID de la predecesora
    kind: int = LINK_FS                 # LINK_FF, LINK_FS, LINK_SF o LINK_SS
    lag: float = 0.0                    # posición (negativa = adelanto); solo se usa su signo


@dataclass
class Task:
    """Tarea leída de MS Project (independiente del motor de lectura)."""

    uid: int
    id: int
    name: str
    outline_level: int
    outline_number: str = ""
    wbs: str = ""                       # código EDT según el campo configurado
    summary: bool = False
    milestone: bool = False
    active: bool = True

    # Línea base seleccionada (sus fechas definen el % planificado)
    baseline_start: datetime | None = None
    baseline_finish: datetime | None = None
    has_baseline: bool = False          # la tarea tiene fechas de línea base guardadas
    baseline_duration: float | None = None      # duración de la línea base en minutos (None = no se leyó)

    # Programa actual
    start: datetime | None = None
    finish: datetime | None = None
    total_slack: float | None = None    # holgura total en minutos (None = no se leyó)
    critical: bool = False              # tarea crítica según MS Project

    # Avance (para el comando actualizar)
    percent_complete: float = 0.0       # 0..100
    physical_percent_complete: float = 0.0
    actual_start: datetime | None = None
    actual_finish: datetime | None = None

    # Programación (para las validaciones)
    duration: float | None = None       # duración del programa actual en minutos (None = no se leyó)
    estimated: bool = False             # duración estimada ("5 días?")
    manual: bool = False                # programada manualmente
    constraint_type: int = CONSTRAINT_ASAP
    constraint_date: datetime | None = None
    predecessors: list[Link] = field(default_factory=list)

    # Campos personalizados configurados
    method: str = ""                    # método de medición
    milestone_weight: float | None = None
    company: str = ""                   # multi-empresa: empresa responsable (columna "Empresa")
    split_a: float | None = None        # multi-empresa: % de la empresa A en tareas "Ambas"
    visible: str = ""                   # campo "Visible": "Sí", "No" o vacío (texto tal como está en MS Project)
    description: str = ""               # campo de descripción de las fichas de tareas (p. ej. las notas)

    # Estructura (la completa wbs.build_tree)
    parent: "Task | None" = field(default=None, repr=False)
    children: list["Task"] = field(default_factory=list, repr=False)

    @property
    def is_leaf(self) -> bool:
        return not self.children


@dataclass
class ProjectData:
    """Proyecto leído de MS Project."""

    name: str
    source: str                          # ruta o descripción del origen
    backend: str                         # "com" | "xml"
    tasks: list[Task]                    # en orden de ID (orden del esquema)
    status_date: datetime | None = None
    currency_symbol: str = ""
    project_start: datetime | None = None
    baseline_label: str = "Línea base"
    baseline_found: bool = True          # False => se usaron las fechas actuales
    warnings: list[Issue] = field(default_factory=list)   # avisos de la lectura (LEC-...)
    calendar: object | None = None                # WorkCalendar del proyecto (motor XML)
    minutes_per_day: float = 480.0                # minutos de un día de duración (Opciones > Programación)
    # Campos configurados que existen en el archivo ([campos] y [multiempresa]): clave de la configuración ->
    # nombre del campo. Sirve para avisar de los campos que existen pero están vacíos en todas las tareas
    fields_found: dict[str, str] = field(default_factory=dict)


def use_current_schedule(project: ProjectData) -> None:
    """Sin línea base guardada: el % planificado se calcula con las fechas actuales."""
    project.baseline_found = False
    project.baseline_label = f"Programa actual (sin {project.baseline_label.lower()})"
    for t in project.tasks:
        t.baseline_start = t.start
        t.baseline_finish = t.finish
        t.baseline_duration = t.duration
