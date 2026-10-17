"""Configuración del exportador (archivo TOML opcional + valores por defecto)."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

try:  # Python 3.11+
    import tomllib
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore[no-redef]


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
    "CPTR de MS Project",
]


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
class Config:
    # [campos] -- campos de MS Project a leer
    wbs_field: str = "WBS"                # WBS | OutlineNumber | Text1..Text30 | alias
    method_field: str = ""                # campo con el método de medición (p.ej. Text10)
    milestone_weight_field: str = ""      # campo numérico con el peso de cada hito (p.ej. Number1)

    # [horizonte]
    start_month: date | None = None       # primer mes (mes 1); None = mes de inicio de la línea base
    months: int = 120

    # [exportacion]
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

    # [metodos]
    methods: list[str] = field(default_factory=lambda: list(DEFAULT_METHODS))

    # [semaforo] -- umbrales para SPI / CPI
    threshold_red: float = 0.90
    threshold_yellow: float = 0.95

    # [multiempresa] -- reparto entre dos empresas (desactivado por defecto)
    multi_enabled: bool = False
    company_field: str = "Empresa"        # columna de MS Project con la empresa responsable
    company_a: str = "EmpresaA"           # valor en la columna y sufijo de las hojas (máx. 15 caracteres)
    company_b: str = "EmpresaB"
    company_both: str = "Ambas"           # valor que indica tarea de ambas empresas
    split_field: str = "% Reparto"        # % de la empresa A en las tareas "Ambas"
    split_default: float = 50.0
    ac_from_project: bool = False         # AC por empresa desde MS Project (si no, se carga a mano)

    source_file: str = ""                 # ruta del TOML leído (informativo)

    def validate(self) -> None:
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
        if not 0 < self.threshold_red <= self.threshold_yellow:
            raise ConfigError("Los umbrales deben cumplir 0 < rojo <= amarillo.")
        if self.default_method and self.default_method not in self.methods:
            self.methods.insert(0, self.default_method)
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

    met = _section(data, "metodos")
    lista = met.get("lista")
    if lista is not None:
        if not isinstance(lista, list) or not all(isinstance(x, str) for x in lista):
            raise ConfigError("[metodos] lista debe ser una lista de textos.")
        cfg.methods = [x.strip() for x in lista if x.strip()]

    sem = _section(data, "semaforo")
    cfg.threshold_red = _get(sem, "rojo", cfg.threshold_red, float)
    cfg.threshold_yellow = _get(sem, "amarillo", cfg.threshold_yellow, float)

    mul = _section(data, "multiempresa")
    cfg.multi_enabled = _get(mul, "activar", cfg.multi_enabled, bool)
    cfg.company_field = _get(mul, "campo_empresa", cfg.company_field, str).strip()
    cfg.company_a = _get(mul, "empresa_a", cfg.company_a, str).strip()
    cfg.company_b = _get(mul, "empresa_b", cfg.company_b, str).strip()
    cfg.company_both = _get(mul, "ambas", cfg.company_both, str).strip()
    cfg.split_field = _get(mul, "campo_reparto", cfg.split_field, str).strip()
    cfg.split_default = _get(mul, "reparto_por_defecto", cfg.split_default, float)
    cfg.ac_from_project = _get(mul, "ac_desde_project", cfg.ac_from_project, bool)

    cfg.source_file = str(path)
    cfg.validate()
    return cfg
