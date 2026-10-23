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

    # Línea base seleccionada (la que define el PV / BAC)
    baseline_cost: float = 0.0
    baseline_start: datetime | None = None
    baseline_finish: datetime | None = None
    has_baseline: bool = False          # la tarea tiene fechas de línea base guardadas

    # Programa actual
    start: datetime | None = None
    finish: datetime | None = None
    cost: float = 0.0
    fixed_cost: float = 0.0
    fixed_cost_accrual: str = "prorrateado"   # "inicio" | "prorrateado" | "fin"

    # Avance (para el comando actualizar)
    percent_complete: float = 0.0       # 0..100
    physical_percent_complete: float = 0.0
    actual_start: datetime | None = None
    actual_finish: datetime | None = None
    actual_cost: float = 0.0
    acwp: float | None = None           # costo real a la fecha de estado (MS Project)
    bcwp: float | None = None           # valor acumulado calculado por MS Project
    ev_baseline_cost: float | None = None   # costo de la línea base que MS Project usa para el CPTR,
                                            # solo si difiere de la línea base del informe

    # Campos personalizados configurados
    method: str = ""                    # método de medición
    milestone_weight: float | None = None
    company: str = ""                   # multi-empresa: empresa responsable (columna "Empresa")
    split_a: float | None = None        # multi-empresa: % de la empresa A en tareas "Ambas"
    visible: str = ""                   # campo "Visible": "Sí", "No" o vacío (texto tal como está en MS Project)

    # Distribución mensual del costo propio de la tarea (PV), si el lector la
    # obtiene directamente de MS Project (datos escalonados en el tiempo).
    timephased_pv: dict[MonthKey, float] | None = None

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
    baseline_found: bool = True          # False => se usó el programa actual
    summary_baseline_cost: float | None = None   # costo de línea base de la tarea resumen del proyecto
    summary_cost: float | None = None
    summary_acwp: float | None = None            # costo real (CRTR) del proyecto a la fecha de estado
    ev_baseline: int = 0                         # línea base que MS Project usa para el valor acumulado
    warnings: list[str] = field(default_factory=list)
    calendar: object | None = None                # WorkCalendar del proyecto (motor XML)


def use_current_schedule(project: ProjectData) -> None:
    """Sin línea base guardada: el PV se calcula con el costo y las fechas actuales."""
    project.baseline_found = False
    project.baseline_label = f"Programa actual (sin {project.baseline_label.lower()})"
    for t in project.tasks:
        t.baseline_cost = t.cost
        t.baseline_start = t.start
        t.baseline_finish = t.finish
        t.timephased_pv = None
    if project.summary_cost is not None:
        project.summary_baseline_cost = project.summary_cost
