"""Modelo de datos común a los lectores (COM y XML)."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

# Clave de mes: (año, mes)
MonthKey = tuple[int, int]


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

    # Programa actual
    start: datetime | None = None
    finish: datetime | None = None

    # Avance (para el comando actualizar)
    percent_complete: float = 0.0       # 0..100
    physical_percent_complete: float = 0.0
    actual_start: datetime | None = None
    actual_finish: datetime | None = None

    # Campos personalizados configurados
    method: str = ""                    # método de medición
    milestone_weight: float | None = None
    company: str = ""                   # multi-empresa: empresa responsable (columna "Empresa")
    split_a: float | None = None        # multi-empresa: % de la empresa A en tareas "Ambas"
    visible: str = ""                   # campo "Visible": "Sí", "No" o vacío (texto tal como está en MS Project)

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
    warnings: list[str] = field(default_factory=list)
    calendar: object | None = None                # WorkCalendar del proyecto (motor XML)


def use_current_schedule(project: ProjectData) -> None:
    """Sin línea base guardada: el % planificado se calcula con las fechas actuales."""
    project.baseline_found = False
    project.baseline_label = f"Programa actual (sin {project.baseline_label.lower()})"
    for t in project.tasks:
        t.baseline_start = t.start
        t.baseline_finish = t.finish
