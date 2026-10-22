"""Configuración del exportador (archivo TOML opcional + valores por defecto)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from datetime import date, datetime
from pathlib import Path

try:  # Python 3.11+
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore[no-redef]


# Modo de cálculo del PV y del EV.
# porcentaje: MS Project da el % planificado y el % ganado de cada fila (por la duración de la línea base de sus
#             tareas y el peso de sus hitos); el importe es el BAC que se ingresa en el libro multiplicado por el %.
# costo:      el PV y el BAC salen del costo de línea base de MS Project y el EV es el BAC × la fracción ganada.
MODE_PERCENT = "porcentaje"
MODE_COST = "costo"
MODES = (MODE_PERCENT, MODE_COST)
DEFAULT_MODE = MODE_PERCENT

# Método que solo existe en el modo costo (el CPTR de MS Project necesita costos).
METHOD_BCWP = "CPTR de MS Project"

# Lista de métodos de medición ofrecida en el desplegable de la columna C.
DEFAULT_METHODS = [
    "Por definir",
    "0/100",
    "50/50",
    "25/75",
    "20/80",
    "Hitos ponderados",
    "% Completado",
    "% Físico completado",
    "Unidades completadas",
    "Avance lineal (LOE)",
    METHOD_BCWP,
]


def parse_mode(text: str) -> str:
    """Modo de cálculo: 'porcentaje' o 'costo' (también 'coste')."""
    value = (text or "").strip().lower()
    if value == "coste":
        value = MODE_COST
    if value not in MODES:
        raise ConfigError(f"modo = '{text}' no es válido: debe ser \"porcentaje\" (por defecto) o \"costo\".")
    return value


def _default_mode() -> str:
    return DEFAULT_MODE


class ConfigError(ValueError):
    """Error en el archivo de configuración o en los parámetros."""


INVALID_SHEET_CHARS = set("[]:*?/\\")


def parse_month(text: str) -> date:
    """Convierte 'AAAA-MM' (o 'AAAA-MM-DD') al primer día de ese mes."""
    m = re.fullmatch(r"\s*(\d{4})-(\d{1,2})(?:-(\d{1,2}))?\s*", text or "")
    if not m:
        raise ConfigError(f"Mes inválido '{text}'. El formato debe ser AAAA-MM, por ejemplo 2026-01.")
    year, month = int(m.group(1)), int(m.group(2))
    if not 1 <= month <= 12:
        raise ConfigError(f"Mes inválido '{text}'.")
    return date(year, month, 1)


def parse_date(text: str) -> date:
    """Convierte 'AAAA-MM-DD' (o 'DD/MM/AAAA') a fecha."""
    text = (text or "").strip()
    parts = None
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if m:
        parts = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        parts = (int(m.group(3)), int(m.group(2)), int(m.group(1)))
    if parts:
        try:
            return date(*parts)
        except ValueError:
            pass
    raise ConfigError(f"Fecha inválida '{text}'. El formato debe ser AAAA-MM-DD, por ejemplo 2026-10-31.")


@dataclass
class ProjectSpec:
    """Archivo de MS Project de un informe con varios archivos (sección [[proyectos]])."""

    name: str                       # nombre corto y fijo del archivo en el libro (p. ej. "P1")
    origin: str                     # archivo .mpp o .xml
    plan: str | None = None         # planificación de este archivo; None = la de [exportacion]
    baseline: int | None = None     # línea base de este archivo; None = la de [exportacion]


MAX_PROJECT_NAME = 30


@dataclass
class Config:
    # [campos] -- campos de MS Project a leer
    wbs_field: str = "WBS"                # WBS | OutlineNumber | Text1..Text30 | alias
    method_field: str = ""                # campo con el método de medición (p.ej. Text10)
    milestone_weight_field: str = ""      # campo numérico con el peso de cada hito (p.ej. Number1)

    # [horizonte]
    start_month: date | None = None       # primer mes (mes 1); None = mes de inicio de la línea base
    months: int = 120

    # [exportacion]
    mode: str = field(default_factory=_default_mode)   # "porcentaje" (por defecto) | "costo"
    plan: str = ""                        # código EDT de la planificación (p. ej. INTERNO); "" = todas
    max_level: int = 3
    baseline: int | None = None           # 0 = Línea base; 1..10 = Línea base N; None = la del proyecto
    pad_wbs: bool = True                  # "1" -> "1.0.0", "1.2" -> "1.2.0"
    include_inactive: bool = False
    distribution: str = "habiles"         # tareas sin datos mensuales: "habiles" | "calendario"
    holidays: set[date] = field(default_factory=set)
    default_method: str = "Por definir"
    currency: str = "EUR"                 # "" = código de moneda del proyecto en MS Project
    decimals: int = 0

    # [metodos] -- None = lista por defecto del modo (sin "CPTR de MS Project" en el modo porcentaje)
    methods_list: list[str] | None = None

    # [semaforo] -- umbrales comunes de SPI y CPI. Rojo: menor que el rojo inferior o mayor que el rojo
    # superior; amarillo: menor que el amarillo inferior o mayor que el amarillo superior; verde: entre ambos
    threshold_red: float = 0.90           # rojo inferior
    threshold_yellow: float = 0.95        # amarillo inferior
    threshold_yellow_upper: float = 1.05  # amarillo superior
    threshold_red_upper: float = 1.10     # rojo superior

    # [multiempresa] -- reparto entre dos empresas (desactivado por defecto)
    multi_enabled: bool = False
    company_field: str = "Empresa"        # columna de MS Project con la empresa responsable
    company_a: str = "EmpresaA"           # valor en la columna y sufijo de las hojas (máx. 15 caracteres)
    company_b: str = "EmpresaB"
    company_both: str = "Ambas"           # valor que indica tarea de ambas empresas
    split_field: str = "% Reparto"        # % de la empresa A en las tareas "Ambas"
    split_default: float = 50.0
    ac_from_project: bool = False         # AC por empresa desde MS Project (si no, se carga a mano)

    # [[proyectos]] -- varios archivos de MS Project en un mismo libro (vacío = un solo archivo)
    projects: list[ProjectSpec] = field(default_factory=list)
    title: str = ""                       # [exportacion] titulo: título del informe ("" = nombre del proyecto)

    source_file: str = ""                 # ruta del TOML leído (informativo)

    @property
    def percent_mode(self) -> bool:
        return self.mode == MODE_PERCENT

    @property
    def methods(self) -> list[str]:
        """Métodos del desplegable de la columna C: los de [metodos] lista o los del modo."""
        if self.methods_list is not None:
            return self.methods_list
        return [m for m in DEFAULT_METHODS if not (self.percent_mode and m == METHOD_BCWP)]

    def project_config(self, name: str) -> "Config":
        """Configuración para leer un archivo de [[proyectos]]: su planificación y su línea base (la general si el
        archivo no las indica)."""
        spec = next((p for p in self.projects if p.name == name), None)
        if spec is None:
            return self
        return replace(self, plan=self.plan if spec.plan is None else spec.plan,
                       baseline=self.baseline if spec.baseline is None else spec.baseline)

    def validate(self) -> None:
        self.mode = parse_mode(self.mode)
        if not 1 <= self.max_level <= 10:
            raise ConfigError("nivel_maximo debe estar entre 1 y 10.")
        if not 1 <= self.months <= 600:
            raise ConfigError("meses debe estar entre 1 y 600.")
        if self.baseline is not None and not 0 <= self.baseline <= 10:
            raise ConfigError("linea_base debe estar entre 0 y 10.")
        if self.distribution not in ("habiles", "calendario"):
            raise ConfigError("distribucion debe ser 'habiles' o 'calendario'.")
        if not 0 <= self.decimals <= 4:
            raise ConfigError("decimales debe estar entre 0 y 4.")
        if not (0 < self.threshold_red <= self.threshold_yellow < self.threshold_yellow_upper
                <= self.threshold_red_upper):
            raise ConfigError("Los umbrales del semáforo deben cumplir 0 < rojo_inferior <= amarillo_inferior < "
                              "amarillo_superior <= rojo_superior (por defecto 0,90 / 0,95 / 1,05 / 1,10).")
        if self.default_method and self.default_method not in self.methods:
            self.methods_list = [self.default_method] + self.methods
        if self.multi_enabled:
            from evm.fields import same_name
            for key, name in (("empresa_a", self.company_a), ("empresa_b", self.company_b)):
                if not name:
                    raise ConfigError(f"[multiempresa] {key} no puede estar vacío.")
                if len(name) > 15:
                    raise ConfigError(f"[multiempresa] {key} = '{name}' supera los 15 caracteres permitidos.")
                bad = [ch for ch in name if ch in INVALID_SHEET_CHARS or ch == "'"]
                if bad:
                    raise ConfigError(f"[multiempresa] {key} = '{name}' contiene caracteres no admitidos en "
                                      f"nombres de hoja: {' '.join(sorted(set(bad)))}")
            if same_name(self.company_a, self.company_b):
                raise ConfigError("[multiempresa] empresa_a y empresa_b deben ser distintas.")
            if not self.company_both or any(same_name(self.company_both, n)
                                            for n in (self.company_a, self.company_b)):
                raise ConfigError("[multiempresa] ambas debe tener un valor distinto de los nombres de las empresas.")
            if not 0 <= self.split_default <= 100:
                raise ConfigError("[multiempresa] reparto_por_defecto debe estar entre 0 y 100.")
        if self.projects:
            from evm.fields import same_name
            seen: list[str] = []
            for p in self.projects:
                if not p.name:
                    raise ConfigError("[[proyectos]]: cada archivo debe tener un 'nombre' (p. ej. \"P1\").")
                if len(p.name) > MAX_PROJECT_NAME or any(ch in p.name for ch in ":;"):
                    raise ConfigError(f"[[proyectos]] nombre = '{p.name}': debe tener como máximo {MAX_PROJECT_NAME} "
                                      "caracteres, sin ':' ni ';'.")
                if any(same_name(p.name, n) for n in seen):
                    raise ConfigError(f"[[proyectos]]: el nombre '{p.name}' está repetido.")
                seen.append(p.name)
                if not p.origin:
                    raise ConfigError(f"[[proyectos]] '{p.name}': falta 'origen' (archivo .mpp o .xml).")
                if p.origin.strip().lower() == "activo":
                    raise ConfigError(f"[[proyectos]] '{p.name}': con varios archivos no se admite 'activo'; se debe "
                                      "indicar la ruta del archivo .mpp o .xml.")
                if p.baseline is not None and not 0 <= p.baseline <= 10:
                    raise ConfigError(f"[[proyectos]] '{p.name}': linea_base debe estar entre 0 y 10.")


def _get(section: dict, key: str, default, kind):
    value = section.get(key, default)
    if value is None:
        return default
    if kind is bool and not isinstance(value, bool):
        raise ConfigError(f"'{key}' debe ser true o false.")
    if kind in (int, float) and (isinstance(value, bool) or not isinstance(value, (int, float))):
        raise ConfigError(f"'{key}' debe ser numérico.")
    if kind is int and isinstance(value, float) and not value.is_integer():
        raise ConfigError(f"'{key}' debe ser un número entero (sin decimales).")
    if kind is str and not isinstance(value, str):
        raise ConfigError(f"'{key}' debe ser texto.")
    return kind(value)


def _section(data: dict, name: str) -> dict:
    value = data.get(name, {})
    if not isinstance(value, dict):
        raise ConfigError(f"[{name}] debe ser una sección: una línea [{name}] seguida de sus claves "
                          f"(se encontró '{name} = ...').")
    return value


def _projects(data: dict, base: Path) -> list[ProjectSpec]:
    """Secciones [[proyectos]]: un archivo de MS Project por sección. Las rutas relativas se toman desde la
    carpeta del archivo de configuración."""
    items = data.get("proyectos")
    if items is None:
        return []
    if not isinstance(items, list) or not all(isinstance(x, dict) for x in items):
        raise ConfigError("Los archivos de MS Project se indican con una sección [[proyectos]] (con dos corchetes) "
                          "por archivo, cada una con 'nombre' y 'origen'.")
    out = []
    for item in items:
        name = _get(item, "nombre", "", str).strip()
        origin = _get(item, "origen", "", str).strip()
        if origin and origin.lower() != "activo" and not Path(origin).is_absolute():
            origin = str(base / origin)
        plan = _get(item, "planificacion", None, str)
        baseline = None
        if item.get("linea_base") not in (None, ""):
            baseline = _get(item, "linea_base", 0, int)
        out.append(ProjectSpec(name, origin, plan.strip() if plan is not None else None, baseline))
    return out


def load_config(path: str | Path | None) -> Config:
    """Lee el archivo TOML (si se indica) y devuelve la configuración."""
    cfg = Config()
    if not path:
        return cfg
    path = Path(path)
    if not path.is_file():
        raise ConfigError(f"No se encontró el archivo de configuración: {path}")
    try:
        text = path.read_text(encoding="utf-8-sig")       # admite UTF-8 con BOM (Bloc de notas)
    except UnicodeDecodeError as exc:
        raise ConfigError(f"El archivo de configuración {path} no está guardado en UTF-8. Se debe guardar con "
                          "codificación UTF-8 (en el Bloc de notas: Guardar como > Codificación: UTF-8).") from exc
    try:
        data = tomllib.loads(text)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"Error de sintaxis en {path}: {exc}") from exc

    campos = _section(data, "campos")
    cfg.wbs_field = _get(campos, "edt", cfg.wbs_field, str).strip() or "WBS"
    cfg.method_field = _get(campos, "metodo", cfg.method_field, str).strip()
    cfg.milestone_weight_field = _get(campos, "peso_hito", cfg.milestone_weight_field, str).strip()

    hor = _section(data, "horizonte")
    start = _get(hor, "mes_inicio", "", str).strip()
    cfg.start_month = parse_month(start) if start else None
    cfg.months = _get(hor, "meses", cfg.months, int)

    exp = _section(data, "exportacion")
    cfg.mode = parse_mode(_get(exp, "modo", cfg.mode, str))
    cfg.plan = _get(exp, "planificacion", cfg.plan, str).strip()
    cfg.max_level = _get(exp, "nivel_maximo", cfg.max_level, int)
    if exp.get("linea_base") not in (None, ""):
        cfg.baseline = _get(exp, "linea_base", 0, int)
    cfg.pad_wbs = _get(exp, "rellenar_edt", cfg.pad_wbs, bool)
    cfg.include_inactive = _get(exp, "incluir_inactivas", cfg.include_inactive, bool)
    cfg.distribution = _get(exp, "distribucion", cfg.distribution, str).strip().lower()
    feriados = exp.get("feriados", [])
    if not isinstance(feriados, list):
        raise ConfigError("'feriados' debe ser una lista de fechas.")
    cfg.holidays = {f.date() if isinstance(f, datetime) else f if isinstance(f, date) else parse_date(str(f))
                    for f in feriados}
    cfg.default_method = _get(exp, "metodo_por_defecto", cfg.default_method, str).strip()
    cfg.currency = _get(exp, "moneda", cfg.currency, str).strip()
    cfg.decimals = _get(exp, "decimales", cfg.decimals, int)
    cfg.title = _get(exp, "titulo", cfg.title, str).strip()

    met = _section(data, "metodos")
    lista = met.get("lista")
    if lista is not None:
        if not isinstance(lista, list) or not all(isinstance(x, str) for x in lista):
            raise ConfigError("[metodos] lista debe ser una lista de textos.")
        cfg.methods_list = [x.strip() for x in lista if x.strip()]

    sem = _section(data, "semaforo")
    for key, alias in (("rojo_inferior", "rojo"), ("amarillo_inferior", "amarillo")):
        if key in sem and alias in sem:
            raise ConfigError(f"[semaforo]: se indicó '{key}' y '{alias}', que son la misma clave ('{alias}' es el "
                              f"nombre anterior). Se debe dejar solo '{key}'.")
    cfg.threshold_red = _get(sem, "rojo_inferior" if "rojo_inferior" in sem else "rojo", cfg.threshold_red, float)
    cfg.threshold_yellow = _get(sem, "amarillo_inferior" if "amarillo_inferior" in sem else "amarillo",
                                cfg.threshold_yellow, float)
    cfg.threshold_yellow_upper = _get(sem, "amarillo_superior", cfg.threshold_yellow_upper, float)
    cfg.threshold_red_upper = _get(sem, "rojo_superior", cfg.threshold_red_upper, float)

    mul = _section(data, "multiempresa")
    cfg.multi_enabled = _get(mul, "activar", cfg.multi_enabled, bool)
    cfg.company_field = _get(mul, "campo_empresa", cfg.company_field, str).strip()
    cfg.company_a = _get(mul, "empresa_a", cfg.company_a, str).strip()
    cfg.company_b = _get(mul, "empresa_b", cfg.company_b, str).strip()
    cfg.company_both = _get(mul, "ambas", cfg.company_both, str).strip()
    cfg.split_field = _get(mul, "campo_reparto", cfg.split_field, str).strip()
    cfg.split_default = _get(mul, "reparto_por_defecto", cfg.split_default, float)
    cfg.ac_from_project = _get(mul, "ac_desde_project", cfg.ac_from_project, bool)

    cfg.projects = _projects(data, path.parent)

    cfg.source_file = str(path)
    cfg.validate()
    return cfg
