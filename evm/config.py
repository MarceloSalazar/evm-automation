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


# Cálculo del PV y del EV: MS Project da el % planificado y el % ganado de cada fila (por la duración de la línea
# base de sus tareas y el peso de sus hitos); el importe es el BAC que se ingresa en el libro multiplicado por el %.
# Libros de versiones anteriores calculados con los costos de MS Project: _meta modo = "costo" (no se admiten).
MODE_PERCENT = "porcentaje"

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
    # Campo de texto "Visible": Sí = fila propia (también por debajo del nivel máximo); No = la tarea y su rama
    # quedan fuera del PV, el EV y el AC (si todas las subtareas tienen No, las representa su tarea superior);
    # vacío = según el nivel máximo. "" = no se lee
    visible_field: str = "Visible"
    visible_field_explicit: bool = False  # la clave se indicó en la configuración (avisar si no existe el campo)
    # Descripción de cada ficha (generar-fichas): "Notas" = notas de la tarea, o un campo de texto; "" = no se lee
    description_field: str = ""

    # [horizonte]
    start_month: date | None = None       # primer mes (mes 1); None = mes de inicio de la línea base
    months: int = 120

    # [exportacion]
    plan: str = ""                        # código EDT de la planificación (p. ej. INTERNO); "" = todas
    max_level: int = 3
    baseline: int | None = None           # 0 = Línea base; 1..10 = Línea base N; None = la del proyecto
    pad_wbs: bool = True                  # "1" -> "1.0.0", "1.2" -> "1.2.0"
    # lista_edt: archivo con los códigos EDT que tienen fila en el libro (reemplaza a max_level y al campo Visible
    # para decidir las filas); wbs_list: sus códigos, tal como están en el archivo (None = sin lista)
    wbs_list_file: str = ""
    wbs_list: list[str] | None = None
    include_inactive: bool = False
    distribution: str = "habiles"         # tareas sin datos mensuales: "habiles" | "calendario"
    holidays: set[date] = field(default_factory=set)
    default_method: str = "Por definir"
    currency: str = "EUR"                 # "" = código de moneda del proyecto en MS Project
    decimals: int = 0
    # Depuración de costes: filas «(coste)» con la previsión de una empresa, pegada
    # en la hoja PV-depuración, para alinear las tareas e hitos de MS Project con esa previsión
    cost_review: bool = False             # depuracion_costes
    cost_review_company: str = ""         # depuracion_empresa: empresa_a o empresa_b ("" = empresa_a)
    cost_review_alignment: bool = False   # depuracion_alineamiento: indicador de alineamiento
    cost_review_threshold: float = 10.0   # depuracion_umbral: % del BAC a partir del cual se resalta un mes

    # [metodos] -- None = lista por defecto
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

    # [validacion] -- método de medición frente a la duración de la línea base del paquete: meses a partir de los
    # cuales se avisa que un paquete con 0/100 (o con 50/50, 25/75, 20/80...) es largo para ese método; 0 = no se avisa
    max_months_0_100: float = 1.0         # meses_0_100
    max_months_fixed: float = 2.0         # meses_formula_fija

    # [[proyectos]] -- varios archivos de MS Project en un mismo libro (vacío = un solo archivo)
    projects: list[ProjectSpec] = field(default_factory=list)
    title: str = ""                       # [exportacion] titulo: título del informe ("" = nombre del proyecto)

    source_file: str = ""                 # ruta del TOML leído (informativo)

    @property
    def cost_review_role(self) -> str | None:
        """Hoja que se depura: "A" o "B" (PV de esa empresa), "" (hoja PV de un libro de una sola empresa) o None si
        la depuración de costes está desactivada."""
        if not self.cost_review:
            return None
        if not self.multi_enabled:
            return ""
        from evm.fields import same_name
        return "B" if same_name(self.cost_review_company, self.company_b) else "A"

    @property
    def methods(self) -> list[str]:
        """Métodos del desplegable de la columna C: los de [metodos] lista o los del programa."""
        if self.methods_list is not None:
            return self.methods_list
        return list(DEFAULT_METHODS)

    def project_config(self, name: str) -> "Config":
        """Configuración para leer un archivo de [[proyectos]]: su planificación y su línea base (la general si el
        archivo no las indica)."""
        spec = next((p for p in self.projects if p.name == name), None)
        if spec is None:
            return self
        return replace(self, plan=self.plan if spec.plan is None else spec.plan,
                       baseline=self.baseline if spec.baseline is None else spec.baseline)

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
            if self.cost_review:
                company = self.cost_review_company.strip()
                if not company:
                    self.cost_review_company = self.company_a
                elif same_name(company, self.company_a):
                    self.cost_review_company = self.company_a
                elif same_name(company, self.company_b):
                    self.cost_review_company = self.company_b
                else:
                    raise ConfigError(f"depuracion_empresa = '{company}' no es una de las empresas de [multiempresa]: "
                                      f"debe ser '{self.company_a}' o '{self.company_b}'.")
        for key, value in (("meses_0_100", self.max_months_0_100), ("meses_formula_fija", self.max_months_fixed)):
            if not 0 <= value <= 120:
                raise ConfigError(f"[validacion] {key} debe estar entre 0 y 120 meses (0 = no se controla).")
        if not 0 < self.cost_review_threshold <= 100:
            raise ConfigError("depuracion_umbral debe ser un % mayor que 0 y como máximo 100 (por defecto 10).")
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


def read_wbs_list(path: str | Path) -> list[str]:
    """Códigos EDT del archivo de lista_edt, en su orden y sin repetir: uno por línea. Se ignoran las líneas vacías,
    las que empiezan con '#' y el texto que sigue al código (por ejemplo, el nombre del elemento). Admite UTF-8 (con
    o sin BOM), UTF-16 (texto Unicode de Excel) y ANSI."""
    path = Path(path)
    if not path.is_file():
        raise ConfigError(f"No se encontró el archivo de la lista EDT (lista_edt): {path}")
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise ConfigError(f"No se pudo leer el archivo de la lista EDT (lista_edt) {path}: {exc}") from exc
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = data.decode("utf-16", errors="replace")
    else:
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("cp1252", errors="replace")
    codes: list[str] = []
    seen: set[str] = set()
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        code = re.split(r"[\s;,]+", line, maxsplit=1)[0]
        if code and code not in seen:
            seen.add(code)
            codes.append(code)
    if not codes:
        raise ConfigError(f"El archivo de la lista EDT (lista_edt) no tiene códigos: {path}. Se debe escribir un "
                          "código EDT por línea (por ejemplo 1.0.0, 1.1.0, 1.1.1).")
    return codes


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
    if "visible" in campos:
        cfg.visible_field = _get(campos, "visible", cfg.visible_field, str).strip()
        cfg.visible_field_explicit = bool(cfg.visible_field)
    cfg.description_field = _get(campos, "descripcion", cfg.description_field, str).strip()

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
    wbs_list = _get(exp, "lista_edt", "", str).strip()
    if wbs_list:
        list_path = Path(wbs_list) if Path(wbs_list).is_absolute() else path.parent / wbs_list
        cfg.wbs_list_file = str(list_path)
        cfg.wbs_list = read_wbs_list(list_path)
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
    cfg.cost_review = _get(exp, "depuracion_costes", cfg.cost_review, bool)
    cfg.cost_review_company = _get(exp, "depuracion_empresa", cfg.cost_review_company, str).strip()
    cfg.cost_review_alignment = _get(exp, "depuracion_alineamiento", cfg.cost_review_alignment, bool)
    cfg.cost_review_threshold = _get(exp, "depuracion_umbral", cfg.cost_review_threshold, float)

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

    val = _section(data, "validacion")
    cfg.max_months_0_100 = _get(val, "meses_0_100", cfg.max_months_0_100, float)
    cfg.max_months_fixed = _get(val, "meses_formula_fija", cfg.max_months_fixed, float)

    mul = _section(data, "multiempresa")
    cfg.multi_enabled = _get(mul, "activar", cfg.multi_enabled, bool)
    cfg.company_field = _get(mul, "campo_empresa", cfg.company_field, str).strip()
    cfg.company_a = _get(mul, "empresa_a", cfg.company_a, str).strip()
    cfg.company_b = _get(mul, "empresa_b", cfg.company_b, str).strip()
    cfg.company_both = _get(mul, "ambas", cfg.company_both, str).strip()
    cfg.split_field = _get(mul, "campo_reparto", cfg.split_field, str).strip()
    cfg.split_default = _get(mul, "reparto_por_defecto", cfg.split_default, float)

    cfg.projects = _projects(data, path.parent)

    cfg.source_file = str(path)
    cfg.validate()
    return cfg
