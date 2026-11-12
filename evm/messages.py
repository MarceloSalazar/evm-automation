"""Catálogo de los mensajes del programa (errores, advertencias e informativas), cada uno con su código.

Cada mensaje que el programa muestra en la consola o en la hoja Validación tiene un código estable de la forma
ÁREA-NNN (p. ej. FEC-003): el área agrupa los mensajes por tema y el número no cambia aunque cambien la severidad o
el texto. CATALOG guarda la severidad, la explicación y qué hacer de cada código; docs/codigos.md se genera a partir
de él (python docs/herramientas/codigos.py) y las pruebas comprueban que está al día y que todo mensaje usa un código
del catálogo.

Un código no se reutiliza ni se renumera: si un mensaje deja de existir, su código queda en el catálogo como
retirado, con el código que lo informa desde entonces; silenciarlo no tiene efecto y se avisa con CFG-047.
Metodología para agregar un mensaje: docs/desarrollo.md, sección 5.
"""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field, replace

SEV_FATAL = "Error fatal"       # detiene el programa: no es una observación de la hoja Validación
SEV_ERROR = "Error"
SEV_WARN = "Advertencia"
SEV_INFO = "Info"

# Temas de las observaciones: agrupan la salida del comando validar y la columna Tema de la hoja Validación
TOPIC_READ = "Lectura y configuración"
TOPIC_WBS = "Estructura y códigos EDT"
TOPIC_VISIBLE = "Campo Visible"
TOPIC_DATES = "Fechas y duración"
TOPIC_WEIGHTS = "Hitos y pesos"
TOPIC_METHOD = "Método de medición"
TOPIC_COMPANY = "Empresas"
TOPIC_PROGRESS = "Avance a la fecha de estado"
TOPIC_LOGIC = "Lógica del cronograma"
TOPIC_BOOK = "Libro"
TOPICS = [TOPIC_READ, TOPIC_WBS, TOPIC_VISIBLE, TOPIC_DATES, TOPIC_WEIGHTS, TOPIC_METHOD, TOPIC_COMPANY,
          TOPIC_PROGRESS, TOPIC_LOGIC, TOPIC_BOOK]

ID_PATTERN = re.compile(r"[A-Z]{3}-\d{3}")

# Alcance de una observación: a qué parte del libro se refiere. La hoja Validación y la consola las ordenan así,
# para resolver primero las de las filas que muestran las hojas PV, EV y AC
TARGET_GENERAL = "General"          # sin tarea: el proyecto, la configuración o el libro
TARGET_ROW = "Fila EDT"             # una fila de las hojas PV, EV y AC (paquete de trabajo o tarea resumen)
TARGET_DETAIL = "Detalle"           # tarea que se suma en una fila (nivel mayor, fuera de la lista EDT o sin código)
TARGET_OUT = "Fuera del libro"      # tarea que no está en ninguna fila (omitida, inactiva o con Visible = No)
TARGETS = [TARGET_GENERAL, TARGET_ROW, TARGET_DETAIL, TARGET_OUT]
SEVERITIES = [SEV_ERROR, SEV_WARN, SEV_INFO]


# ----------------------------------------------------------------------------
# Observaciones y errores con código
# ----------------------------------------------------------------------------
@dataclass
class Issue:
    """Observación de la consola y de la hoja Validación. La severidad y el tema, si no se indican, son los del
    código en el catálogo."""
    id: str                     # código del catálogo (p. ej. FEC-003)
    code: str = ""              # código EDT de la fila o la tarea ("" = observación general)
    name: str = ""              # nombre de la fila o la tarea
    message: str = ""
    severity: str = ""          # Error, Advertencia o Info (los códigos de severidad variable la indican)
    topic: str = ""
    uid: int | None = None      # UniqueID de la tarea de la observación (None = sin tarea o se busca por su código)
    target: str = ""            # alcance (TARGETS); "" = sin asignar (lo asignan wbs.locate_issues y wbs.located)
    row: str = ""               # código EDT de la fila del libro de la tarea: la suya o la fila que la suma
    file: str = ""              # varios archivos ([[proyectos]]): nombre del archivo de la observación
    tasks: int = 0              # línea agrupada (grouped): cantidad de observaciones que reúne
    variant: str = ""           # caso del código (p. ej. "sin sucesoras" en LOG-003): el informe de mejoras agrupa
    #                             por él las tareas de un código

    def __post_init__(self) -> None:
        entry = CATALOG.get(self.id)
        if not self.severity:
            self.severity = entry.issue_severity if entry else SEV_WARN
        if not self.topic:
            area = AREAS.get(self.id[:3])
            self.topic = area.topic if area else TOPIC_BOOK


def by_priority(issues: list[Issue]) -> list[Issue]:
    """Las observaciones por severidad (los errores primero) y, en cada severidad, por alcance (TARGETS: general,
    filas EDT, tareas de detalle y fuera del libro); las del mismo grupo, en su orden."""
    severity = {s: k for k, s in enumerate(SEVERITIES)}
    target = {t: k for k, t in enumerate(TARGETS)}
    return sorted(issues, key=lambda i: (severity.get(i.severity, 1), target.get(i.target or TARGET_GENERAL, 0)))


def grouped(issues: list[Issue], hint: str) -> list[Issue]:
    """Las observaciones de las tareas de detalle y de las que están fuera del libro, en una línea por código y fila
    del libro con la cantidad de tareas (si son dos o más); las demás, sin cambios. Cada línea va en el lugar de la
    primera observación de su grupo; hint indica dónde se ven una por una."""
    groups: dict[tuple, list[Issue]] = {}
    order: list = []
    for i in issues:
        if i.target in (TARGET_DETAIL, TARGET_OUT):
            key = (i.id, i.severity, i.target, i.row, i.file)
            if key not in groups:
                groups[key] = []
                order.append(key)
            groups[key].append(i)
        else:
            order.append(i)
    out: list[Issue] = []
    for x in order:
        found = [x] if isinstance(x, Issue) else groups[x]
        if len(found) == 1:
            out.append(found[0])
            continue
        first, n = found[0], len(found)
        entry = CATALOG.get(first.id)
        names = [i.code or i.name for i in found]
        shown = ", ".join(names[:10]) + (f" y {n - 10} más" if n > 10 else "")
        where = f"de detalle de la fila {first.row}" if first.target == TARGET_DETAIL else "fuera del libro"
        tag = f"[{first.file}] " if first.file else ""
        out.append(replace(first, code="", name=f"{n} tareas", uid=None, tasks=n,
                           message=f"{tag}{entry.title if entry else first.id}: {n} tareas {where} ({shown}). "
                                   f"{hint}"))
    return out


class EvmError(Exception):
    """Error que detiene el programa, con su código del catálogo (p. ej. CFG-003)."""

    def __init__(self, id: str, message: str):
        super().__init__(message)
        self.id = id


class InputError(EvmError, ValueError):
    """Datos que impiden continuar (fecha de corte, valores del libro...)."""


class MissingFileError(EvmError, FileNotFoundError):
    """No existe el archivo indicado."""


class SaveError(EvmError, PermissionError):
    """No se puede reemplazar un archivo (por ejemplo, el libro está abierto en Excel)."""


class ProjectAppError(EvmError, RuntimeError):
    """MS Project no se puede usar por automatización COM."""


# ----------------------------------------------------------------------------
# Catálogo
# ----------------------------------------------------------------------------
# Comandos en que aparecen los mensajes (columna "Dónde" del documento)
W_ALL = "Todos los comandos"
W_SOURCE = "generar, actualizar, validar y generar-fichas"
W_READ = "generar, actualizar, validar y generar-fichas (lectura de MS Project)"
W_GV = "generar (hoja Validación) y validar"
W_G = "generar (hoja Validación)"
W_V = "validar"
W_OUT = "generar, generar-fichas, generar-cliente y validar --informe"
W_BOOK = "actualizar, analizar, procesar-costes y generar --migrar-desde (lectura del libro)"
W_MIG = "generar --migrar-desde (hoja Validación)"
W_REVIEW = "generar con depuracion_costes (hoja Validación)"
W_PROC = "procesar-costes"
W_ACT = "actualizar"
W_AN = "actualizar y analizar"
W_ROW = "actualizar (observaciones por fila)"
W_FIC = "generar-fichas"
W_CLI = "generar-cliente"
W_CLI_V = "generar-cliente y validar (con [cliente] campo)"


@dataclass(frozen=True)
class Area:
    prefix: str
    title: str
    topic: str                  # tema de la hoja Validación y del comando validar
    where: str                  # comandos en que aparecen sus mensajes (salvo que el código indique otros)
    summary: str


AREAS = {a.prefix: a for a in [
    Area("CFG", "Configuración y opciones", TOPIC_READ, W_ALL,
         "Archivo de configuración, opciones de la línea de comandos, horizonte y códigos silenciados."),
    Area("LEC", "Lectura de MS Project", TOPIC_READ, W_READ,
         "Apertura del archivo .mpp o .xml y campos configurados que no se encuentran o están vacíos."),
    Area("EDT", "Estructura y códigos EDT", TOPIC_WBS, W_GV,
         "Formato, dependencia, orden y duplicados de los códigos EDT, planificación, lista EDT y filas del libro."),
    Area("VIS", "Campo Visible", TOPIC_VISIBLE, W_GV,
         "Valores del campo Visible y ramas que quedan fuera del cálculo."),
    Area("FEC", "Fechas y duración", TOPIC_DATES, W_GV,
         "Fechas de la línea base y del programa actual, hitos y duración de las tareas."),
    Area("HIT", "Hitos y pesos", TOPIC_WEIGHTS, W_GV,
         "Peso de los hitos, método «Hitos ponderados» y filas sin % planificado."),
    Area("MET", "Método de medición", TOPIC_METHOD, W_GV,
         "Método de medición del EV de cada fila."),
    Area("EMP", "Empresas", TOPIC_COMPANY, W_GV,
         "Multi-empresa: empresa y % de reparto de cada tarea."),
    Area("AVA", "Avance a la fecha de estado", TOPIC_PROGRESS, W_V,
         "Fechas reales, % completado, avance actualizado a la fecha de estado y fuera de secuencia, e índices "
         "DCMA del avance: BEI, tareas incumplidas y CPLI (solo validar)."),
    Area("LOG", "Lógica del cronograma", TOPIC_LOGIC, W_V,
         "Vínculos, delimitaciones, holgura y duración de las tareas, según los criterios de la evaluación DCMA de "
         "14 puntos (solo validar)."),
    Area("LIB", "Libro", TOPIC_BOOK, W_BOOK,
         "Libro Excel: recordatorio del BAC y libros dañados, modificados o abiertos en Excel."),
    Area("MIG", "Migración (--migrar-desde)", TOPIC_BOOK, W_MIG,
         "Traslado del BAC, el % ganado, el AC y los métodos de un libro anterior."),
    Area("DEP", "Depuración de costes", TOPIC_BOOK, W_REVIEW,
         "Hoja PV-depuración, su traslado con --migrar-desde y la revisión del plan de costes (procesar-costes)."),
    Area("ACT", "Actualizar y analizar", TOPIC_BOOK, W_ACT,
         "Carga del % ganado del mes: fecha de corte, filas del libro, BAC y observaciones por fila."),
    Area("FIC", "Fichas de tareas", TOPIC_BOOK, W_FIC,
         "Comando generar-fichas."),
    Area("CLI", "Planificación para el cliente", TOPIC_BOOK, W_CLI,
         "Comando generar-cliente: tareas que se muestran al cliente, «Otras tareas», resúmenes convertidos, "
         "vínculos nuevos y comparación con la planificación interna."),
    Area("SIS", "Errores del sistema", TOPIC_BOOK, W_ALL,
         "Errores de archivos del sistema y errores inesperados del programa."),
]}


@dataclass(frozen=True)
class Code:
    id: str
    severities: tuple[str, ...]     # la primera es la habitual; las demás, según el caso
    title: str
    cause: str
    action: str
    where: str = ""                 # "" = los comandos del área
    retired: str = ""               # versión en que se retiró ("" = vigente)
    by: str = ""                    # código retirado: el código que lo informa desde entonces ("" = ninguno)

    @property
    def area(self) -> Area:
        return AREAS[self.id[:3]]

    @property
    def issue_severity(self) -> str:
        """Severidad de la observación (los errores que detienen el programa también pueden ser observaciones:
        el comando validar los informa como errores y sigue)."""
        return next((s for s in self.severities if s != SEV_FATAL), SEV_ERROR)

    @property
    def silenceable(self) -> bool:
        return SEV_FATAL not in self.severities and self.id not in UNSILENCEABLE

    @property
    def location(self) -> str:
        return self.where or self.area.where


def _code(id: str, severity: str | tuple[str, ...], title: str, cause: str, action: str, where: str = "",
          retired: str = "", by: str = "") -> Code:
    severities = severity if isinstance(severity, tuple) else (severity,)
    return Code(id, severities, title, cause, action, where, retired, by)


F, E, W, I = SEV_FATAL, SEV_ERROR, SEV_WARN, SEV_INFO      # noqa: E741 (abreviaturas del catálogo)

_CODES = [
    # --- CFG: configuración y opciones --------------------------------------------------------------------------
    _code("CFG-001", F, "No se encontró el archivo de configuración",
          "El archivo indicado con --config no existe en esa ruta.",
          "Revisar la ruta (las rutas relativas se toman desde la carpeta en que se ejecuta el programa) o crear el "
          "archivo a partir de config.ejemplo.toml."),
    _code("CFG-002", F, "El archivo de configuración no está en UTF-8",
          "El archivo se guardó con otra codificación (por ejemplo ANSI) y tiene acentos o eñes.",
          "Guardarlo con codificación UTF-8 (en el Bloc de notas: Guardar como > Codificación: UTF-8)."),
    _code("CFG-003", F, "Error de sintaxis en el archivo de configuración",
          "El archivo no es TOML válido: comillas sin cerrar, una clave repetida, una sección mal escrita... El "
          "mensaje indica la línea y la columna.",
          "Corregir la línea indicada. Los textos van entre comillas, los números y true/false sin comillas."),
    _code("CFG-004", F, "Clave con un valor del tipo incorrecto",
          "Una clave espera true o false, un número, un número entero o un texto, y tiene otro tipo de valor (por "
          "ejemplo, meses = \"24\" con comillas).",
          "Corregir el valor de la clave indicada (ver config.ejemplo.toml)."),
    _code("CFG-005", F, "Sección escrita como clave",
          "Una sección (por ejemplo [exportacion]) se escribió como clave: 'exportacion = ...'.",
          "Escribir la sección en una línea propia, [nombre], seguida de sus claves."),
    _code("CFG-006", F, "Archivos de MS Project mal indicados",
          "La clave proyectos no es una lista de secciones [[proyectos]].",
          "Indicar cada archivo en una sección [[proyectos]] (con dos corchetes), con 'nombre' y 'origen'."),
    _code("CFG-007", F, "Mes con formato no válido",
          "Un mes (mes_inicio o --mes-inicio) no tiene el formato AAAA-MM o el mes no existe.",
          "Indicar el mes como AAAA-MM, por ejemplo 2026-01."),
    _code("CFG-008", F, "Fecha con formato no válido",
          "Una fecha (--fecha-corte, feriados...) no tiene el formato AAAA-MM-DD o la fecha no existe.",
          "Indicar la fecha como AAAA-MM-DD, por ejemplo 2026-10-31."),
    _code("CFG-009", F, "Nivel máximo fuera de rango",
          "nivel_maximo (o --nivel-max) no está entre 1 y 10.",
          "Indicar un nivel entre 1 y 10 (por defecto, 3)."),
    _code("CFG-010", F, "Cantidad de meses fuera de rango",
          "meses (o --meses) no está entre 1 y 600.",
          "Indicar una cantidad de meses entre 1 y 600 (por defecto, 120)."),
    _code("CFG-011", F, "Línea base fuera de rango",
          "linea_base (general, de un archivo de [[proyectos]] o --linea-base) no está entre 0 y 10.",
          "Indicar 0 para «Línea base» o de 1 a 10 para «Línea base N»."),
    _code("CFG-012", F, "Distribución no válida",
          "distribucion no es 'habiles' ni 'calendario'.",
          "Indicar distribucion = \"habiles\" (por defecto) o \"calendario\"."),
    _code("CFG-013", F, "Decimales fuera de rango",
          "decimales no está entre 0 y 4.",
          "Indicar una cantidad de decimales entre 0 y 4."),
    _code("CFG-014", F, "Umbrales del semáforo incoherentes",
          "Los umbrales de [semaforo] no cumplen 0 < rojo_inferior <= amarillo_inferior < amarillo_superior <= "
          "rojo_superior.",
          "Corregir los umbrales (por defecto 0,90 / 0,95 / 1,05 / 1,10)."),
    _code("CFG-015", F, "Umbral del semáforo indicado dos veces",
          "En [semaforo] se indicó una clave y también su nombre anterior, que es la misma clave.",
          "Dejar solo el nombre actual de la clave."),
    _code("CFG-016", F, "Umbral de [validacion] fuera de rango",
          "meses_0_100 o meses_formula_fija no está entre 0 y 120 meses.",
          "Indicar un valor entre 0 y 120 (0 = no se controla)."),
    _code("CFG-017", F, "Feriados mal indicados",
          "feriados no es una lista de fechas.",
          "Indicar los feriados como lista: feriados = [\"2026-12-25\", \"2027-01-01\"]."),
    _code("CFG-018", F, "Lista de métodos mal indicada",
          "[metodos] lista no es una lista de textos.",
          "Indicar los métodos como lista de textos entre comillas."),
    _code("CFG-019", F, "Nombre de empresa vacío",
          "En [multiempresa], empresa_a o empresa_b está vacío.",
          "Indicar el nombre de las dos empresas."),
    _code("CFG-020", F, "Nombre de empresa demasiado largo",
          "El nombre de una empresa supera los 15 caracteres: no cabría en el nombre de sus hojas (PV-EmpresaA...).",
          "Usar un nombre o una abreviatura de hasta 15 caracteres."),
    _code("CFG-021", F, "Nombre de empresa con caracteres no admitidos",
          "El nombre de una empresa tiene caracteres que Excel no admite en el nombre de una hoja ([ ] : * ? / \\ "
          "o comillas simples).",
          "Quitar esos caracteres del nombre."),
    _code("CFG-022", F, "Las dos empresas tienen el mismo nombre",
          "empresa_a y empresa_b son iguales (sin distinguir mayúsculas).",
          "Indicar dos nombres distintos."),
    _code("CFG-023", F, "Valor de 'ambas' no válido",
          "El valor que indica las tareas de las dos empresas (ambas) está vacío o es el nombre de una empresa.",
          "Indicar un valor distinto de los nombres de las empresas (por defecto, \"Ambas\")."),
    _code("CFG-024", F, "Reparto por defecto fuera de rango",
          "reparto_por_defecto no está entre 0 y 100.",
          "Indicar el % de la empresa A en las tareas compartidas, entre 0 y 100 (por defecto, 50)."),
    _code("CFG-025", F, "Empresa de la depuración no reconocida",
          "depuracion_empresa no es empresa_a ni empresa_b.",
          "Indicar el nombre de una de las dos empresas (o dejarlo vacío para depurar empresa_a)."),
    _code("CFG-026", F, "Umbral de la depuración fuera de rango",
          "depuracion_umbral no es mayor que 0 y como máximo 100.",
          "Indicar un % mayor que 0 y como máximo 100 (por defecto, 10)."),
    _code("CFG-027", F, "Archivo de [[proyectos]] sin nombre",
          "Una sección [[proyectos]] no tiene 'nombre'.",
          "Indicar un nombre breve para cada archivo (p. ej. \"P1\")."),
    _code("CFG-028", F, "Nombre de archivo de [[proyectos]] no válido",
          "El nombre de un archivo de [[proyectos]] es demasiado largo o tiene ':' o ';'.",
          "Usar un nombre breve, sin ':' ni ';'."),
    _code("CFG-029", F, "Nombre de archivo de [[proyectos]] repetido",
          "Dos secciones [[proyectos]] tienen el mismo nombre.",
          "Indicar un nombre distinto para cada archivo."),
    _code("CFG-030", F, "Archivo de [[proyectos]] sin origen",
          "Una sección [[proyectos]] no indica 'origen'.",
          "Indicar la ruta del archivo .mpp o .xml."),
    _code("CFG-031", F, "'activo' con varios archivos",
          "Con varios archivos ([[proyectos]]), un origen es 'activo' (el proyecto abierto en MS Project).",
          "Indicar la ruta del archivo .mpp o .xml de cada archivo."),
    _code("CFG-032", F, "No se encontró el archivo de la lista EDT",
          "El archivo indicado en lista_edt no existe.",
          "Revisar la ruta (las rutas relativas se toman desde la carpeta del archivo de configuración)."),
    _code("CFG-033", F, "No se pudo leer el archivo de la lista EDT",
          "El archivo de lista_edt existe, pero no se puede leer (permisos, archivo bloqueado...).",
          "Cerrar el archivo en otros programas y revisar sus permisos."),
    _code("CFG-034", F, "La lista EDT no tiene códigos",
          "El archivo de lista_edt está vacío o solo tiene comentarios.",
          "Escribir un código EDT por línea (por ejemplo 1.0.0, 1.1.0, 1.1.1)."),
    _code("CFG-035", F, "Extensión del archivo de origen no reconocida",
          "El archivo de --origen no es .mpp, .mpt, .mpd ni .xml.",
          "Usar un archivo .mpp (requiere MS Project) o .xml (MS Project > Guardar como > XML), o indicar --motor.",
          W_SOURCE),
    _code("CFG-036", F, "Falta el archivo de MS Project",
          "No se indicó --origen ni hay secciones [[proyectos]] en la configuración.",
          "Indicar el archivo con --origen.", W_SOURCE),
    _code("CFG-037", F, "Varios --origen sin [[proyectos]]",
          "Se indicó --origen más de una vez sin secciones [[proyectos]].",
          "Para combinar varios archivos en un libro, indicarlos en secciones [[proyectos]] de la configuración.",
          W_SOURCE),
    _code("CFG-038", F, "--planificacion o --linea-base con varios archivos",
          "Con varios archivos, la planificación y la línea base son de cada archivo.",
          "Indicarlas en la sección [[proyectos]] de cada archivo.", W_SOURCE),
    _code("CFG-039", F, "--origen no válido con varios archivos",
          "Con [[proyectos]], --origen debe ser NOMBRE=RUTA, con el nombre de uno de los archivos.",
          "Indicar --origen P1=C:\\ruta\\archivo.mpp, con un nombre de [[proyectos]].", W_SOURCE),
    _code("CFG-040", F, "Falta --salida con varios archivos",
          "Con varios archivos no hay un nombre de salida por defecto.",
          "Indicar el archivo a crear con --salida (o con --informe en validar).", W_OUT),
    _code("CFG-041", F, "Extensión del archivo de salida no válida",
          "El libro se guarda como .xlsx; las fichas, el informe de mejoras de validar y el informe de control de "
          "generar-cliente, como .md; la planificación del cliente, como .mpp.",
          "Indicar --salida (o --informe o --destino) con la extensión que corresponde.", W_OUT),
    _code("CFG-042", F, "El archivo de salida ya existe",
          "El archivo de --salida (o de --informe o --destino) existe y no se indicó --sobrescribir.",
          "Usar --sobrescribir para reemplazarlo o indicar otro nombre. Para conservar los EV/AC cargados en un "
          "libro, se genera un archivo nuevo con --migrar-desde.", W_OUT),
    _code("CFG-043", F, "--migrar-desde igual a --salida",
          "El libro anterior y el libro nuevo son el mismo archivo.",
          "Generar el libro nuevo con otro nombre.", "generar"),
    _code("CFG-044", W, "% planificado anterior al mes 1",
          "Algunas filas tienen % planificado antes del primer mes del horizonte (mes_inicio): se acumuló en la "
          "columna del mes 1.",
          "Si el horizonte debe empezar antes, indicar mes_inicio (o --mes-inicio) y volver a generar.", W_GV),
    _code("CFG-045", W, "% planificado posterior al último mes",
          "Algunas filas tienen % planificado después del último mes del horizonte: no aparece en las columnas "
          "mensuales y su PV no llega al BAC. El mensaje indica hasta qué % del BAC llega el PV de cada fila, el "
          "último mes con % planificado y cuántos meses hacen falta.",
          "Aumentar meses (o --meses) a la cantidad indicada y volver a generar.", W_GV),
    _code("CFG-046", W, "Monedas distintas en los archivos",
          "Con varios archivos, los proyectos tienen monedas distintas en MS Project; el libro muestra una de ellas.",
          "Indicar la moneda del libro en [exportacion] moneda.", W_GV),
    _code("CFG-047", W, "Código desconocido o retirado en silenciar",
          "Un elemento de [validacion] silenciar o de --silenciar no es un código del catálogo, no tiene el formato "
          "ÁREA-NNN (o ÁREA-NNN:EDT) o es un código retirado: se ignora. Si el código se retiró, el mensaje indica "
          "con qué código se informa ahora.",
          "Corregir el código según este documento (un código retirado se reemplaza por el que lo informa "
          "ahora)."),
    _code("CFG-048", W, "Código que no se puede silenciar",
          "Se pidió silenciar un error que detiene el programa, el resumen de las observaciones silenciadas o un "
          "aviso del registro de diagnóstico (LEC-022).",
          "Quitarlo de silenciar: esos mensajes siempre se muestran."),
    _code("CFG-049", I, "Observaciones silenciadas",
          "Resumen de las observaciones que no se muestran por [validacion] silenciar o --silenciar, con la cantidad "
          "de cada código.",
          "Ninguna. Para volver a verlas, quitar los códigos de silenciar."),
    _code("CFG-050", F, "loe_maximo fuera de rango",
          "[validacion] loe_maximo no es un % entre 0 y 100.",
          "Indicar un % entre 0 y 100 (por defecto 15; 0 = no se controla)."),
    _code("CFG-051", F, "No existe la carpeta de salida",
          "La carpeta del archivo indicado en --salida no existe.",
          "Crear la carpeta o indicar otra ruta en --salida.", W_OUT),
    _code("CFG-052", W, "Fin previsto posterior al último mes",
          "El fin previsto del cronograma actual (las fechas de las tareas de las filas del libro) es posterior al "
          "último mes del horizonte: el EV y el AC de los meses posteriores no se podrán cargar en el libro. El "
          "mensaje indica cuántos meses hacen falta.",
          "Aumentar meses (o --meses) a la cantidad indicada y volver a generar; si el atraso no es real, revisar "
          "el cronograma.", W_GV),
    _code("CFG-053", F, "Varias planificaciones fuera de validar",
          "Se indicaron varias planificaciones separadas por comas (p. ej. INTERNO,CLIENTE): solo validar las revisa, "
          "una por una, y con un solo archivo (sin [[proyectos]]).",
          "Indicar una sola planificación en planificacion o --planificacion; para revisar varias, usar validar "
          "--planificacion INTERNO,CLIENTE.", "generar, actualizar, validar y generar-fichas"),
    _code("CFG-054", F, "Falta el campo del cliente",
          "generar-cliente necesita el campo de MS Project que indica qué tareas se muestran al cliente.",
          "Indicarlo en [cliente] campo, por ejemplo campo = \"Etiqueta Cliente\" (ver config.ejemplo.toml).", W_CLI),
    _code("CFG-055", F, "Cálculo del precio mal configurado",
          "La sección [depuracion_precio] (precio de las EDT interna en la depuración de costes) no es válida: "
          "falta calculo_precio o la lista de componentes, la expresión no se puede leer o usa un nombre que no está "
          "en [depuracion_precio.componentes] ni en [depuracion_precio.recargos], un nombre o un texto está repetido "
          "o hay más de 8 coeficientes. El mensaje indica el motivo.",
          "Corregir la sección según config.ejemplo.toml: la expresión admite números (con punto decimal), los "
          "nombres de las dos listas, + - * / y paréntesis."),
    _code("CFG-056", W, "Fechas de estado distintas en los archivos",
          "Con varios archivos ([[proyectos]]), los proyectos tienen fechas de estado distintas, o alguno no tiene: el "
          "avance de cada archivo corresponde a otro día y actualizar se detiene (ACT-005 o ACT-004) si no se "
          "indica --fecha-corte. La consola de generar muestra siempre este aviso, antes que los demás.",
          "Registrar el avance de todos los archivos a la misma fecha (el último día del mes informado o su último "
          "día laborable) y fijar esa fecha de estado en cada uno (Proyecto > Fecha de estado).", W_GV),

    # --- LEC: lectura de MS Project -----------------------------------------------------------------------------
    _code("LEC-001", F, "No existe el archivo de origen",
          "El archivo de --origen (o de [[proyectos]]) no existe en esa ruta.",
          "Revisar la ruta y el nombre del archivo."),
    _code("LEC-002", F, "El XML no es válido",
          "El archivo .xml está dañado o incompleto y no se puede leer.",
          "Volver a guardarlo desde MS Project (Archivo > Guardar como > XML)."),
    _code("LEC-003", F, "El XML no es de MS Project",
          "El archivo .xml no tiene el elemento <Project> de MS Project.",
          "Usar un XML guardado desde MS Project (Archivo > Guardar como > XML)."),
    _code("LEC-004", F, "El XML no tiene tareas",
          "El archivo .xml de MS Project no tiene la lista de tareas.",
          "Revisar que el proyecto tenga tareas y volver a guardarlo como XML."),
    _code("LEC-005", F, "Falta pywin32 o MS Project para leer .mpp",
          "Los archivos .mpp se leen con MS Project por automatización COM, que necesita el paquete pywin32.",
          "Instalar pywin32 (pip install pywin32) o guardar el proyecto como XML y procesar ese archivo."),
    _code("LEC-006", F, "No se pudo iniciar MS Project",
          "MS Project de escritorio no está instalado o no se puede iniciar por automatización COM.",
          "Verificar la instalación de MS Project, o exportar el proyecto a XML y usar ese archivo."),
    _code("LEC-007", F, "MS Project no responde",
          "MS Project no responde por automatización COM: puede estar cerrándose o mostrando un diálogo (inicio de "
          "sesión, recuperación de archivos...).",
          "Abrir MS Project y cerrar el diálogo, o esperar a que termine de cerrarse (o cerrarlo desde el "
          "Administrador de tareas), y volver a intentar."),
    _code("LEC-008", F, "No hay proyectos abiertos en MS Project",
          "Se indicó --origen activo y MS Project no tiene ningún proyecto abierto.",
          "Abrir el proyecto en MS Project o indicar la ruta del archivo."),
    _code("LEC-009", F, "MS Project no pudo abrir el archivo",
          "MS Project no pudo abrir el .mpp (archivo dañado, de una versión más reciente, bloqueado o con un "
          "diálogo pendiente).",
          "Abrir el archivo a mano en MS Project para ver el problema, o guardarlo como XML y usar ese archivo."),
    _code("LEC-010", W, "No se encontró un campo de la configuración",
          "Un campo indicado en [campos] o en [multiempresa] no existe en el proyecto. El mensaje indica cuál y su "
          "efecto: edt, se usa el campo EDT (WBS) de MS Project; metodo, todas las filas tienen el método por "
          "defecto; peso_hito, los hitos no tienen peso; visible, las filas del libro siguen el nivel máximo; "
          "descripcion, las fichas de tareas no tienen descripción; campo_empresa, todas las tareas se tratan como "
          "compartidas; campo_reparto, las tareas compartidas se reparten con reparto_por_defecto.",
          "Revisar el nombre del campo (como se ve en MS Project o el nombre interno, p. ej. Texto1), o dejar la "
          "clave vacía si no se usa."),
    _code("LEC-011", W, "No se encontró el campo de método",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-012", W, "No se encontró el campo de peso de hitos",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-013", W, "No se encontró el campo Visible",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-014", W, "No se encontró el campo de descripción",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-015", W, "No se encontró el campo de empresa",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-016", W, "No se encontró el campo de reparto",
          "Desde la versión 2.6.0, todo campo de la configuración que no se encuentra en el proyecto se informa con "
          "LEC-010.",
          "Ninguna.", retired="2.6.0", by="LEC-010"),
    _code("LEC-017", W, "Tareas externas omitidas",
          "El proyecto tiene tareas externas (vínculos con otros proyectos): no forman parte del cálculo.",
          "Ninguna, si es lo esperado."),
    _code("LEC-018", W, "Excepción de calendario no interpretada",
          "Un calendario del XML tiene una excepción con una periodicidad que el programa no interpreta: se omite. "
          "Solo afecta el reparto de la duración de las tareas entre los meses.",
          "Si la excepción es importante (p. ej. un cierre anual), indicarla como fechas en feriados."),
    _code("LEC-019", W, "Campo de la configuración vacío",
          "Un campo indicado en [campos] o en [multiempresa] existe, pero está vacío en todas las tareas. El mensaje "
          "indica cuál y su efecto: metodo, todas las filas tienen el método por defecto (las filas sin método no "
          "se informan aparte con MET-005); peso_hito, ningún hito tiene peso y los hitos no tienen % planificado "
          "propio; campo_empresa, todas las tareas se tratan como compartidas (las tareas sin empresa no se "
          "informan aparte con EMP-003).",
          "Cargar el campo en MS Project, o revisar el nombre indicado en la configuración.", W_GV),
    _code("LEC-020", W, "Ningún hito tiene peso",
          "Desde la versión 2.6.0, todo campo de la configuración vacío en todas las tareas se informa con LEC-019.",
          "Ninguna.", W_GV, retired="2.6.0", by="LEC-019"),
    _code("LEC-021", W, "Campo de empresa vacío",
          "Desde la versión 2.6.0, todo campo de la configuración vacío en todas las tareas se informa con LEC-019.",
          "Ninguna.", W_GV, retired="2.6.0", by="LEC-019"),
    _code("LEC-022", W, "No se pudo cerrar el proyecto en MS Project",
          "Al terminar la lectura de un .mpp, MS Project no cerró el proyecto (por ejemplo, por un diálogo abierto): "
          "queda abierto, sin guardar cambios. Se muestra en el registro de diagnóstico (WARNING evm.reader_com), "
          "por eso no se puede silenciar.",
          "Cerrar el proyecto a mano en MS Project, sin guardar.", "generar, actualizar, validar y generar-fichas "
          "con archivos .mpp"),
    _code("LEC-023", W, "Campo de reparto vacío",
          "Hasta la versión 2.5, en multi-empresa, se avisaba cuando el campo de reparto existía, pero estaba vacío en "
          "todas las tareas. Desde la versión 2.6.0 no se avisa: una tarea compartida (ambas) sin reparto se reparte "
          "con reparto_por_defecto, como está previsto.",
          "Ninguna.", W_GV, retired="2.6.0"),
    _code("LEC-024", I, "Líneas base guardadas",
          "Líneas base guardadas y la que usa el % planificado: la cantidad de tareas revisadas (las de la "
          "planificación exportada) con fecha de comienzo en cada una.",
          "Ninguna, si el % planificado usa la línea base aprobada; si no, indicar la línea base con linea_base o "
          "--linea-base.", W_V),
    _code("LEC-025", W, "Calendario del proyecto no leído",
          "Con un .mpp, MS Project no devolvió el calendario del proyecto (la jornada de cada día de la semana y los "
          "días no laborables) como se esperaba: el % planificado se reparte con un calendario de lunes a viernes, "
          "de 8:00 a 12:00 y de 13:00 a 17:00, y los feriados de la configuración.",
          "Comparar con el resultado del XML del mismo proyecto (Archivo > Guardar como > XML), que lee el "
          "calendario del archivo, y usar el XML si difieren.", "generar, actualizar, validar y generar-fichas con "
          "archivos .mpp"),
    _code("LEC-026", F, "No se encontró el campo del cliente",
          "El campo de [cliente] campo no existe en el archivo de MS Project, ni con ese nombre ni como alias: no se "
          "puede saber qué tareas se muestran al cliente.",
          "Revisar el nombre del campo en [cliente] campo (el alias, por ejemplo \"Etiqueta Cliente\", o el nombre, "
          "por ejemplo \"Text15\").", W_CLI),

    # --- EDT: estructura y códigos EDT --------------------------------------------------------------------------
    _code("EDT-001", E, "Código EDT con formato no válido",
          "El código tiene segmentos vacíos (p. ej. 1..2 o 1.2.).",
          "Corregir el código en MS Project: números separados por un punto, como 1.2.3."),
    _code("EDT-002", W, "Código EDT no numérico",
          "El código no está formado por números separados por puntos, así que no se puede comprobar con el de su "
          "tarea superior.",
          "Usar códigos numéricos (1.2.3). Los códigos sin puntos, como INTERNO, se admiten como planificación."),
    _code("EDT-003", W, "Código de nivel inferior sin tarea superior con código",
          "La tarea tiene un código de nivel 2 o más, pero ninguna tarea superior tiene código EDT.",
          "Asignar a la tarea superior el código indicado en el mensaje, o revisar el código de la tarea."),
    _code("EDT-004", E, "Mismo código EDT que la tarea superior",
          "Desde la versión 2.6.0, un código EDT igual al de la tarea superior se informa con EDT-005: tampoco depende "
          "del de la tarea superior.",
          "Ninguna.", retired="2.6.0", by="EDT-005"),
    _code("EDT-005", E, "El código EDT no depende del de la tarea superior",
          "El código no empieza por el de su tarea superior (p. ej. 1.3.1 bajo 1.2.0), o es el mismo código que el "
          "de su tarea superior.",
          "Corregir el código según el patrón indicado en el mensaje (p. ej. 1.2.1 bajo 1.2.0)."),
    _code("EDT-006", W, "El código EDT salta un nivel",
          "El código tiene dos niveles más que su tarea superior: probablemente falta el código de una tarea "
          "intermedia.",
          "Asignar código EDT a la tarea intermedia, o corregir el código."),
    _code("EDT-007", W, "Código EDT fuera de orden",
          "El código es menor que el de la tarea anterior del mismo nivel.",
          "Ordenar las tareas en MS Project o corregir la numeración."),
    _code("EDT-008", I, "Numeración con saltos",
          "Entre tareas del mismo nivel faltan números (p. ej. 1.2.1 → 1.2.3).",
          "Ninguna, si es intencional (p. ej. un paquete eliminado); si no, corregir los códigos."),
    _code("EDT-009", E, "Código EDT duplicado",
          "Varias tareas tienen el mismo código EDT: el libro no puede distinguir sus filas. Los ceros finales no "
          "cuentan: 1.2, 1.2.0 y 1.2.0.0 son el mismo código.",
          "Corregir los códigos. Si el proyecto tiene varias planificaciones (p. ej. INTERNO y CLIENTE), elegir una "
          "con planificacion."),
    _code("EDT-010", (F, E), "No se encontró la planificación",
          "Ninguna tarea tiene el código EDT (o el nombre) de la planificación indicada. validar lo informa como "
          "error y revisa el proyecto completo.",
          "Revisar planificacion (o --planificacion): el mensaje lista las tareas de nivel superior. Si el archivo "
          "tiene una sola planificación, dejar planificacion vacía (planificacion = \"\") o quitar la clave: se usa "
          "todo el proyecto.",
          "generar, actualizar, generar-fichas, generar-cliente y validar"),
    _code("EDT-011", (F, E), "Código EDT en dos archivos",
          "Con varios archivos de MS Project, el mismo código EDT está en dos de ellos: el libro identifica cada "
          "fila por su código. validar lo informa como error.",
          "Revisar los códigos EDT o la planificación de cada archivo (p. ej. 1.0.0 a 8.0.0 en uno y 9.0.0 en otro).",
          "generar, generar-fichas y validar"),
    _code("EDT-012", (F, E), "Ningún código de la lista EDT existe",
          "Ningún código de lista_edt corresponde a una tarea: el libro no tendría filas. validar lo informa como "
          "error.",
          "Revisar la lista, el campo EDT ([campos] edt) y la planificación elegida.",
          "generar, generar-fichas y validar"),
    _code("EDT-013", W, "Códigos de la lista EDT sin tarea",
          "Algunos códigos de lista_edt no existen en MS Project, o su tarea está inactiva o fuera de la "
          "planificación: no tienen fila en el libro. Un resumen con la cantidad y una línea por código.",
          "Corregir la lista o el código de la tarea."),
    _code("EDT-014", W, "Código de la lista EDT sin tarea",
          "Desde la versión 2.6.0, el resumen y cada código de la lista se informan con EDT-013.",
          "Ninguna.", retired="2.6.0", by="EDT-013"),
    _code("EDT-015", I, "Elemento de la lista EDT con su tarea superior fuera de la lista",
          "El código está en la lista EDT y el de su tarea superior no: la fila queda debajo del elemento listado "
          "más cercano.",
          "Ninguna, si es lo esperado; si no, agregar la tarea superior a la lista."),
    _code("EDT-016", I, "Tareas inactivas excluidas",
          "Las tareas inactivas (y su rama) no forman parte del % planificado. Con una planificación, se cuentan "
          "solo las suyas.",
          "Ninguna. Para incluirlas, activarlas en MS Project."),
    _code("EDT-017", I, "Planificación exportada",
          "Se exportan solo las tareas que dependen de la planificación indicada (p. ej. INTERNO). Si la tarea de la "
          "planificación no tiene código EDT, se elige por su nombre.",
          "Ninguna."),
    _code("EDT-018", I, "Fila «Otras tareas»",
          "Las tareas sin código EDT (o de nivel mayor al máximo, o fuera de la lista EDT) que dependen de una tarea "
          "resumen forman su fila -D, con su propio BAC.",
          "Ingresar el BAC de la fila -D, o asignar códigos EDT a esas tareas."),
    _code("EDT-019", (W, I), "Tareas fuera de la lista EDT omitidas",
          "Algunas tareas no están en la lista EDT ni dependen de un elemento de la lista: no forman parte del libro "
          "ni del PV y el EV. Es una advertencia si alguna tiene duración o es un hito con peso.",
          "Agregar a la lista los elementos que faltan, o revisar sus códigos."),
    _code("EDT-020", (W, I), "Tareas sin código EDT omitidas",
          "Algunas tareas sin código EDT no dependen de ningún elemento EDT: no forman parte de ninguna fila. Es "
          "una advertencia si alguna tiene duración o es un hito con peso.",
          "Asignarles código EDT o moverlas debajo de un elemento EDT."),
    _code("EDT-021", W, "Ninguna tarea tiene código EDT",
          "El libro no tiene filas.",
          "Revisar el campo EDT configurado ([campos] edt) y la planificación elegida."),
    _code("EDT-022", I, "Tareas agregadas en su elemento superior",
          "Las tareas sin código EDT o de nivel mayor al máximo (o fuera de la lista EDT) se suman en la fila de su "
          "elemento EDT superior.",
          "Ninguna."),

    # --- VIS: campo Visible -------------------------------------------------------------------------------------
    _code("VIS-001", W, "Valor de Visible no reconocido",
          "El campo Visible tiene un valor distinto de Sí o No: se usa el nivel máximo.",
          "Escribir Sí, No o dejarlo vacío."),
    _code("VIS-002", W, "Visible = Sí dentro de una rama excluida",
          "La tarea tiene Visible = Sí, pero depende de una tarea con Visible = No: queda fuera del cálculo con esa "
          "rama.",
          "Quitar el No de la tarea superior o el Sí de la tarea."),
    _code("VIS-003", W, "Visible = Sí bajo una tarea que representa a sus subtareas",
          "La tarea tiene Visible = Sí, pero depende de una tarea con Visible = No cuyas hermanas también lo tienen: "
          "se suma con su rama.",
          "Revisar los valores de Visible de esa rama."),
    _code("VIS-004", W, "Visible = Sí en una tarea sin código EDT",
          "Una tarea sin código EDT no puede tener fila propia: se suma en su fila superior.",
          "Asignarle código EDT o quitar el Sí."),
    _code("VIS-005", W, "Todos los elementos de nivel superior con Visible = No",
          "Si se aplicara, el libro no tendría filas: se ignora en esos elementos.",
          "Revisar el campo Visible de los elementos de nivel superior."),
    _code("VIS-006", I, "Campo Visible sin usar con lista EDT",
          "Con lista_edt, las filas del libro las decide la lista y el campo Visible no se usa.",
          "Ninguna."),
    _code("VIS-007", I, "Tareas con Visible = No",
          "Esas tareas (con su rama) no forman parte del % planificado ni del % ganado. Un resumen con la cantidad "
          "y una línea por tarea.",
          "El BAC y el AC que se ingresan en el libro no las deben incluir."),
    _code("VIS-008", I, "Tarea con Visible = No",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con VIS-007.",
          "Ninguna.", retired="2.6.0", by="VIS-007"),
    _code("VIS-009", I, "Todas las subtareas con Visible = No",
          "La tarea representa a sus subtareas: su fila es el paquete de trabajo, con toda su rama (o se suman en su "
          "fila superior).",
          "Ninguna, si es lo esperado."),
    _code("VIS-010", I, "Visible = Sí por debajo del nivel máximo",
          "Esas tareas tienen fila propia aunque su código sea más detallado que el nivel máximo.",
          "Ninguna."),

    # --- FEC: fechas y duración ---------------------------------------------------------------------------------
    _code("FEC-001", E, "Fin anterior al comienzo",
          "La tarea termina antes de comenzar, en la línea base (fin previsto anterior al comienzo previsto) o en el "
          "programa actual. El mensaje indica cuál.",
          "Corregir las fechas de la tarea en MS Project; si el error está en la línea base, guardar de nuevo la "
          "línea base de la tarea."),
    _code("FEC-002", E, "Fin anterior al comienzo",
          "Desde la versión 2.6.0, el fin anterior al comienzo, en la línea base o en el programa actual, se informa "
          "con FEC-001.",
          "Ninguna.", retired="2.6.0", by="FEC-001"),
    _code("FEC-003", W, "Tarea sin Comienzo o sin Fin",
          "La tarea no tiene fechas en el programa actual (p. ej. una tarea programada manualmente sin fechas).",
          "Indicar sus fechas o su duración."),
    _code("FEC-004", W, "Hito con duración",
          "La tarea está marcada como hito y tiene duración en el programa actual o en la línea base que usa el % "
          "planificado (el mensaje indica dónde): se calcula como hito, sin su duración (sin Peso hito no pesa en el "
          "% planificado; con Peso hito, pesa entero en el mes de su fecha de fin). La duración en la línea base "
          "suele aparecer cuando una tarea se convierte en hito después de guardar la línea base.",
          "Si no es un hito, quitar la marca de hito (Información de la tarea > Avanzado). Si lo es y la duración "
          "está en la línea base, guardar de nuevo su línea base (Establecer línea base > Tareas seleccionadas), "
          "como parte de un cambio de línea base aprobado."),
    _code("FEC-005", W, "Duración 0 sin marca de hito",
          "La tarea dura 0 y no es un hito: no pesa en el % planificado de su fila.",
          "Si es un hito, marcarla como hito (y, si corresponde, indicar su peso)."),
    _code("FEC-006", W, "Duración estimada",
          "Algunas tareas tienen duración estimada (con «?», p. ej. «5 días?»): la duración define el % planificado. "
          "Un resumen con la cantidad y una línea por tarea.",
          "Confirmar o corregir la duración de esas tareas."),
    _code("FEC-007", W, "Duración estimada",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con FEC-006.",
          "Ninguna.", retired="2.6.0", by="FEC-006"),
    _code("FEC-008", W, "Tareas programadas manualmente",
          "MS Project no recalcula las fechas de esas tareas cuando cambian sus predecesoras o el avance: la lógica "
          "del cronograma deja de funcionar en ese punto. En un cronograma integrado no debería haber ninguna.",
          "Programarlas automáticamente. Si es intencional (p. ej. fechas impuestas por terceros), silenciar "
          "FEC-008."),
    _code("FEC-009", W, "El proyecto no tiene la línea base seleccionada",
          "El % planificado se calculó con las fechas actuales de las tareas: el PV sigue al programa actual y no "
          "muestra el atraso respecto del plan aprobado. Si el proyecto tiene otras líneas base guardadas (con un XML; "
          "con un .mpp, solo en validar), el mensaje las indica, con la cantidad de tareas de cada una. La consola de "
          "generar muestra siempre este aviso, antes que los demás.",
          "Guardar la línea base en MS Project (Proyecto > Establecer línea base) y volver a exportar; o, si otra "
          "línea base guardada es la aprobada y está completa, indicarla con linea_base (en [[proyectos]], la del "
          "archivo) o --linea-base."),
    _code("FEC-010", W, "Tareas sin fechas de línea base",
          "Algunas tareas no tienen fechas de línea base (p. ej. se agregaron después de guardarla): su % "
          "planificado se calculó con sus fechas actuales. Un resumen con la cantidad y una línea por tarea.",
          "Si corresponde, actualizar la línea base de esas tareas."),
    _code("FEC-011", W, "Tarea sin fechas de línea base",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con FEC-010.",
          "Ninguna.", retired="2.6.0", by="FEC-010"),
    _code("FEC-012", W, "Tarea sin fechas",
          "La tarea no tiene fechas de línea base ni fechas actuales: no aporta % planificado.",
          "Indicar sus fechas o su duración."),
    _code("FEC-013", W, "Línea base de tareas resumen distinta de la de su detalle",
          "La línea base de algunas tareas resumen no va del comienzo más temprano al fin más tardío de la línea "
          "base de sus tareas de detalle activas, o termina antes de comenzar. Suele pasar al guardar la línea base "
          "de tareas seleccionadas sin resumirla en sus tareas resumen. MS Project muestra en esas tareas fechas de "
          "línea base que no corresponden a su detalle; el libro no se ve afectado, porque usa las fechas de las "
          "tareas de detalle. Un resumen y una línea por tarea resumen, con las fechas de su línea base y las de sus "
          "tareas de detalle.",
          "Guardar la línea base de esas tareas con Establecer línea base > Tareas seleccionadas y Resumir líneas "
          "base: Para todas las tareas de resumen, como parte de un cambio de línea base aprobado.", W_V),
    _code("FEC-014", W, "Tarea resumen con la línea base sin resumir",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con FEC-013.",
          "Ninguna.", W_V, retired="2.6.0", by="FEC-013"),
    _code("FEC-015", W, "Línea base incompleta",
          "Tareas con solo el comienzo o solo el fin en una línea base. Cada tarea de detalle con la línea base que "
          "usa el % planificado incompleta tiene su línea: se toma como si durara 0 y no pesa en el % planificado de "
          "su fila. validar agrega un resumen de todas las tareas (también las resumen) con una línea base "
          "incompleta: con un .mpp solo se revisa la línea base que usa el % planificado; con un XML, todas.",
          "Guardar de nuevo la línea base de esas tareas (Establecer línea base > Tareas seleccionadas), o borrar la "
          "fecha suelta si esa línea base no se usa."),
    _code("FEC-016", W, "Líneas base incompletas",
          "Desde la versión 2.6.0, el resumen de las líneas base incompletas y cada tarea se informan con FEC-015.",
          "Ninguna.", W_V, retired="2.6.0", by="FEC-015"),
    _code("FEC-017", W, "Tareas con calendario propio",
          "Algunas tareas tienen un calendario de tarea distinto del calendario del proyecto: MS Project calcula sus "
          "fechas con él, pero el programa reparte su duración entre los meses con el calendario del proyecto, y "
          "su % planificado puede diferir algo del de MS Project.",
          "Ninguna, si el calendario de la tarea es necesario; si no, quitarlo (Información de la tarea > "
          "Avanzado > Calendario: Ninguno).", W_V),
    _code("FEC-018", W, "Tarea con la línea base de un día",
          "La tarea dura al menos 5 días laborables en el programa actual, y en la línea base que usa el % "
          "planificado, un día o menos (comienzo y fin el mismo día): en el % planificado de su fila pesa como una "
          "tarea de un día o menos, entera en ese mes. Suele pasar cuando la línea base se guardó antes de planificar "
          "la tarea (p. ej. cuando era un hito o una tarea de un día).",
          "Guardar de nuevo la línea base de la tarea (Establecer línea base > Tareas seleccionadas), como parte de un "
          "cambio de línea base aprobado."),

    # --- HIT: hitos y pesos -------------------------------------------------------------------------------------
    _code("HIT-001", W, "Peso en una tarea que no es un hito de detalle",
          "La tarea tiene peso de hito, pero no es un hito o es una tarea resumen: el peso no se usa (las tareas "
          "pesan por su duración).",
          "Quitar el peso, o marcar la tarea como hito."),
    _code("HIT-002", W, "«Hitos ponderados» sin hitos",
          "La fila usa el método «Hitos ponderados» y no tiene hitos: actualizar calculará el % ganado con el % "
          "físico completado (o el % completado).",
          "Agregar los hitos del paquete o cambiar el método."),
    _code("HIT-003", I, "«Hitos ponderados» con hitos sin peso",
          "Ningún hito de la fila tiene peso: los hitos se ponderan por igual.",
          "Indicar el peso de los hitos, si no deben pesar igual."),
    _code("HIT-004", I, "«Hitos ponderados» con algunos hitos sin peso",
          "Los hitos sin peso no suman % ganado al cumplirse.",
          "Indicar el peso de esos hitos, si deben sumar."),
    _code("HIT-005", W, "Fila sin % planificado",
          "Ninguna tarea de la fila tiene fechas ni es un hito con peso: su BAC no se puede repartir en el tiempo.",
          "Indicar fechas o duración a sus tareas, o el peso de sus hitos."),
    _code("HIT-006", W, "Hitos con peso negativo",
          "Algunos hitos tienen peso negativo: se toma como 0.",
          "Corregir el peso de esos hitos."),
    _code("HIT-007", W, "Los pesos de los hitos suman más de 100 %",
          "Los pesos se escalaron a 100 % y las tareas sin hito de la fila no reciben % planificado.",
          "Corregir los pesos para que sumen como máximo 100 %."),
    _code("HIT-008", W, "Pesos de hitos menores a 100 % sin tareas con duración",
          "La fila solo tiene hitos y sus pesos suman menos de 100 %: se escalaron a 100 %.",
          "Corregir los pesos para que sumen 100 %."),
    _code("HIT-009", W, "Tareas sin duración ni hitos con peso",
          "Ninguna tarea de la fila tiene duración ni es un hito con peso: todas pesan igual.",
          "Indicar la duración de las tareas o el peso de los hitos."),

    # --- MET: método de medición --------------------------------------------------------------------------------
    _code("MET-001", W, "Método no reconocido",
          "El método de la fila no es uno de los admitidos: actualizar calculará el % ganado con el % completado.",
          "Corregir el método en MS Project o en la columna C del libro."),
    _code("MET-002", I, "Subtareas con otro método",
          "Algunas subtareas de la fila tienen otro método en MS Project: se usa el de la tarea de la fila.",
          "Ninguna, si es lo esperado."),
    _code("MET-003", W, "Fila sin método con subtareas con método",
          "La tarea de la fila no tiene método y sus subtareas sí: se usa el método por defecto.",
          "Indicar el método en la tarea de la fila."),
    _code("MET-004", W, "Método de fórmula fija en un paquete largo",
          "Con 0/100 (o 50/50, 25/75...) el EV queda fijo hasta que el paquete termina: en un paquete largo, el EV no "
          "refleja el avance.",
          "Dividir el paquete o usar un método de avance (% completado, hitos ponderados). Los umbrales son "
          "[validacion] meses_0_100 y meses_formula_fija."),
    _code("MET-005", I, "Filas sin método definido",
          "Algunas filas no tienen método de medición (columna C). No se informa si el campo de método no se "
          "encontró o está vacío en todas las tareas (LEC-010 o LEC-019): todas las filas tienen el método por "
          "defecto.",
          "Indicar el método en MS Project o en la columna C del libro."),

    # --- EMP: empresas ------------------------------------------------------------------------------------------
    _code("EMP-001", I, "% de reparto en una tarea que no es compartida",
          "La tarea tiene % de reparto, pero su empresa no es la de las tareas compartidas (ambas): el reparto no se "
          "usa.",
          "Quitar el reparto o cambiar la empresa de la tarea."),
    _code("EMP-002", I, "Empresa sin tareas en la fila",
          "La empresa no participa en las tareas de algunas filas: si se le asigna BAC, su PV se reparte con el % "
          "planificado de la fila completa.",
          "Ninguna, si no se le asigna BAC en esas filas."),
    _code("EMP-003", W, "Tareas sin empresa",
          "Algunas tareas no tienen empresa (ni la heredan de sus tareas superiores): se tratan como compartidas, "
          "con el reparto por defecto. Un resumen con la cantidad y una línea por tarea (la de nivel más alto de "
          "cada rama sin empresa, con sus subtareas). No se informa si el campo de empresa no se encontró o está "
          "vacío en todas las tareas (LEC-010 o LEC-019).",
          "Indicar la empresa de esas tareas o de su tarea superior."),
    _code("EMP-004", W, "Tarea sin empresa",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con EMP-003.",
          "Ninguna.", retired="2.6.0", by="EMP-003"),
    _code("EMP-005", W, "Empresa no reconocida",
          "El valor del campo de empresa no es empresa_a, empresa_b ni ambas: la tarea se trata como compartida.",
          "Corregir el valor en MS Project."),
    _code("EMP-006", W, "% de reparto fuera de rango",
          "El % de reparto de la tarea no es mayor que 0 y como máximo 100: se usa el reparto por defecto.",
          "Corregir el % de reparto."),

    # --- AVA: avance a la fecha de estado (validar) ---------------------------------------------------------------
    _code("AVA-001", E, "Fin real sin Comienzo real",
          "La tarea tiene Fin real y no tiene Comienzo real.",
          "Registrar el Comienzo real."),
    _code("AVA-002", E, "Fin real anterior al Comienzo real",
          "Las fechas reales de la tarea están invertidas.",
          "Corregir las fechas reales."),
    _code("AVA-003", W, "% completado sin Comienzo real",
          "La tarea tiene % completado y no tiene Comienzo real.",
          "Registrar el Comienzo real (MS Project lo completa al registrar el avance)."),
    _code("AVA-004", W, "100 % completado sin Fin real",
          "La tarea está completa y no tiene Fin real.",
          "Registrar el Fin real."),
    _code("AVA-005", W, "Fin real con % completado menor a 100",
          "La tarea tiene Fin real, pero el programa la considera terminada solo con el 100 %.",
          "Completar el % completado o quitar el Fin real."),
    _code("AVA-006", W, "Avance sin fecha de estado",
          "El proyecto tiene avance cargado y no tiene fecha de estado: no se revisó el avance a la fecha de estado.",
          "Indicar la fecha de estado en MS Project (Proyecto > Información del proyecto) o con --fecha-corte; "
          "actualizar la necesita."),
    _code("AVA-007", I, "Sin fecha de estado ni avance",
          "El proyecto no tiene fecha de estado ni avance cargado: no se revisó el avance.",
          "Ninguna; la fecha se puede indicar con --fecha-corte."),
    _code("AVA-008", E, "Avance registrado en el futuro",
          "El Comienzo real o el Fin real de la tarea es posterior a la fecha de estado.",
          "Corregir la fecha real o la fecha de estado."),
    _code("AVA-009", W, "Tarea atrasada sin terminar",
          "Desde la versión 2.6.0, cada tarea con trabajo pendiente antes de la fecha de estado se informa con AVA-011,"
          " como su resumen.",
          "Ninguna.", retired="2.6.0", by="AVA-011"),
    _code("AVA-010", W, "Tarea sin comenzar con Comienzo pasado",
          "Desde la versión 2.6.0, cada tarea con trabajo pendiente antes de la fecha de estado se informa con AVA-011,"
          " como su resumen.",
          "Ninguna.", retired="2.6.0", by="AVA-011"),
    _code("AVA-011", W, "Trabajo pendiente antes de la fecha de estado",
          "Tareas sin terminar con trabajo programado antes de la fecha de estado: su Fin es anterior a la fecha de "
          "estado, o no tienen avance ni Comienzo real y su Comienzo es anterior. El cronograma no está "
          "actualizado. Un resumen con la cantidad y una línea por tarea.",
          "Registrar el avance o reprogramar el trabajo pendiente después de la fecha de estado (Proyecto > "
          "Actualizar proyecto)."),
    _code("AVA-012", I, "Tareas que debían comenzar y no comenzaron",
          "Desde la versión 2.6.0, las tareas atrasadas respecto de la línea base (las que debían comenzar o terminar) "
          "se informan juntas con AVA-013.",
          "Ninguna.", retired="2.6.0", by="AVA-013"),
    _code("AVA-013", (W, I), "Tareas atrasadas respecto de la línea base",
          "Según la línea base, las tareas que debían terminar a la fecha de estado y no terminaron o terminaron "
          "después de su fin de línea base, y las que debían comenzar y no comenzaron. Es una advertencia si el "
          "índice de ejecución de la línea base (BEI) es menor que 0,95 o si las tareas incumplidas superan el 5 % "
          "(criterios DCMA): los dos índices están en la evaluación DCMA (puntos 14 y 11). El BEI solo cuenta las "
          "tareas terminadas que tienen línea base (el mensaje indica cuántas terminadas no la tienen).",
          "Revisar el atraso de las tareas que lista el mensaje."),
    _code("AVA-014", (W, I), "Tareas incumplidas respecto de la línea base",
          "Desde la versión 2.6.0, las tareas atrasadas respecto de la línea base (las que debían comenzar o terminar) "
          "se informan juntas con AVA-013, y el % de tareas incumplidas está en la evaluación DCMA (punto 11).",
          "Ninguna.", retired="2.6.0", by="AVA-013"),
    _code("AVA-015", (W, I), "Índice de longitud del camino crítico (CPLI)",
          "Días laborables que quedan hasta el fin previsto del proyecto, más la holgura frente al fin de la línea "
          "base, divididos por los días que quedan. Menor que 1 indica que el fin previsto es posterior al de la "
          "línea base; es una advertencia si es menor que 0,95 (criterio DCMA). El mensaje indica los días que "
          "quedan y el atraso; el índice está en la evaluación DCMA (punto 13).",
          "Revisar las tareas críticas: con un CPLI bajo, cumplir el fin de la línea base exige recuperar el "
          "atraso en el trabajo que queda."),
    _code("AVA-016", W, "Avance fuera de secuencia",
          "Tareas con avance que su vínculo con la predecesora todavía no permite: la tarea tiene avance (o está "
          "terminada) y su predecesora no terminó (vínculo fin a comienzo o fin a fin) o no comenzó (comienzo a "
          "comienzo o comienzo a fin). Las fechas y la holgura del trabajo que queda pueden no ser realistas. Los "
          "vínculos con adelanto no se revisan (LOG-006). Un resumen y una línea por tarea.",
          "Corregir el avance registrado de la tarea o de su predecesora, o el vínculo, si ya no representa la "
          "secuencia real del trabajo."),
    _code("AVA-017", W, "Tarea con avance fuera de secuencia",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con AVA-016.",
          "Ninguna.", retired="2.6.0", by="AVA-016"),
    _code("AVA-018", W, "Fecha de estado que no es fin de mes",
          "La fecha de estado no es el último día del mes ni su último día laborable (según el calendario del "
          "proyecto): actualizar carga el avance en el mes de la fecha de estado, y el PV acumulado de ese mes "
          "incluye el mes completo (ACT-018); el avance y el PV no corresponden a la misma fecha.",
          "Fijar la fecha de estado en el último día del mes informado, o en su último día laborable (Proyecto > "
          "Fecha de estado), o indicar --fecha-corte."),

    # --- LOG: lógica del cronograma (validar) ---------------------------------------------------------------------
    _code("LOG-001", W, "Cronograma sin vínculos",
          "Las fechas no se recalculan cuando cambian las duraciones o el avance.",
          "Vincular las tareas (cada tarea con al menos una predecesora y una sucesora)."),
    _code("LOG-002", W, "Tareas sin predecesoras o sin sucesoras",
          "Hasta la versión 2.5, contaba en un resumen las tareas sin predecesoras o sin sucesoras, con su % de las "
          "tareas sin terminar. Desde la versión 2.6.0, cada tarea se informa con LOG-003, y el % está en el punto 1 "
          "de la evaluación DCMA.",
          "Ninguna.", retired="2.6.0", by="LOG-003"),
    _code("LOG-003", W, "Tarea sin predecesoras o sin sucesoras",
          "Tarea de detalle sin terminar a la que le faltan las predecesoras, las sucesoras o las dos (el mensaje "
          "indica cuáles): sus fechas no dependen de la lógica del cronograma o su fin no afecta a ninguna otra "
          "tarea. No les faltan predecesoras a las que comienzan con el proyecto y tienen sucesoras, ni sucesoras a "
          "las que terminan con el proyecto y tienen predecesoras; un vínculo de su tarea resumen cuenta como suyo. "
          "Cuenta para el punto 1 de la evaluación DCMA (hasta 5 % de las tareas sin terminar).",
          "Vincular la tarea: agregar la predecesora o la sucesora que le falta."),
    _code("LOG-004", W, "Tarea sin predecesoras",
          "Desde la versión 2.6.0, toda tarea sin predecesoras o sin sucesoras se informa con LOG-003.",
          "Ninguna.", retired="2.6.0", by="LOG-003"),
    _code("LOG-005", W, "Tarea sin sucesoras",
          "Desde la versión 2.6.0, toda tarea sin predecesoras o sin sucesoras se informa con LOG-003.",
          "Ninguna.", retired="2.6.0", by="LOG-003"),
    _code("LOG-006", W, "Vínculo con adelanto",
          "El vínculo tiene posición negativa (adelanto); la evaluación DCMA recomienda no usarlos. Una línea por "
          "tarea, con sus predecesoras con adelanto.",
          "Reemplazar el adelanto por tareas o vínculos sin adelanto."),
    _code("LOG-007", W, "Vínculo comienzo a fin (CF)",
          "Es un tipo de vínculo poco habitual.",
          "Usar un vínculo fin a comienzo (FC), salvo que sea necesario."),
    _code("LOG-008", (W, I), "Vínculos con posición positiva",
          "Una espera fija oculta trabajo que conviene representar como tarea. El mensaje indica los vínculos de la "
          "posición más larga a la más corta, con su valor y la predecesora. Es una advertencia si superan el 5 % de "
          "los vínculos (criterio DCMA; el % está en la evaluación DCMA, punto 3).",
          "Reemplazar las esperas por tareas, si representan trabajo."),
    _code("LOG-009", W, "Pocos vínculos fin a comienzo",
          "Menos del 90 % de los vínculos son fin a comienzo (FC) (criterio DCMA; el % está en la evaluación DCMA, "
          "punto 4). El mensaje indica cuántos vínculos son de cada otro tipo y las tareas que los tienen.",
          "Revisar los vínculos de otro tipo."),
    _code("LOG-010", W, "Tarea resumen con vínculos",
          "Los vínculos de una tarea resumen ocultan la lógica de sus tareas de detalle.",
          "Vincular las tareas de detalle."),
    _code("LOG-011", W, "Tareas con delimitaciones",
          "Las delimitaciones que fijan fechas (Debe comenzar el, No finalizar después del...) impiden que la lógica "
          "mueva esas tareas (DCMA: hasta 5 % de las tareas sin terminar, punto 5). Un resumen y una línea por "
          "tarea, con su delimitación.",
          "Usar «Lo antes posible» y, si hace falta, una fecha límite."),
    _code("LOG-012", W, "Tarea con delimitación",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con LOG-011.",
          "Ninguna.", retired="2.6.0", by="LOG-011"),
    _code("LOG-013", (W, I), "Tareas con holgura alta",
          "Tareas sin terminar con más de 44 días laborables de holgura total: suele indicar vínculos que faltan. "
          "Es una advertencia si superan el 5 % de las tareas sin terminar (criterio DCMA; el % está en la "
          "evaluación DCMA, punto 6).",
          "Revisar las sucesoras de esas tareas: la holgura alta desaparece al vincularlas con el trabajo que "
          "depende de ellas."),
    _code("LOG-014", W, "Tareas con holgura negativa",
          "Tareas sin terminar cuyo fin previsto supera una fecha límite o una delimitación. La evaluación DCMA no "
          "admite holgura negativa (punto 7). Un resumen y una línea por tarea, con su holgura total.",
          "Recuperar el atraso (cambiar la secuencia o la duración del trabajo que queda) o revisar las fechas "
          "límite y las delimitaciones de esas tareas y de sus predecesoras."),
    _code("LOG-015", W, "Tarea con holgura negativa",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con LOG-014.",
          "Ninguna.", retired="2.6.0", by="LOG-014"),
    _code("LOG-016", (W, I), "Tareas de duración alta",
          "Tareas sin terminar con más de 44 días laborables de duración en la línea base (sin los hitos ni las "
          "filas LOE): son difíciles de seguir y de medir. Es una advertencia si superan el 5 % de las tareas sin "
          "terminar (criterio DCMA; el % está en la evaluación DCMA, punto 8).",
          "Dividir esas tareas en tareas más cortas, con entregables intermedios."),
    _code("LOG-017", W, "Varias rutas críticas",
          "La opción Calcular varias rutas críticas de MS Project está activada: cada cadena de tareas tiene su "
          "propia ruta crítica y toda tarea sin sucesoras queda sin holgura, es decir, crítica. La holgura y las "
          "tareas críticas (columnas Holgura y Crítica de Indicadores EDT, paquetes críticos de actualizar) no "
          "indican la ruta crítica del proyecto.",
          "Desactivarla (Archivo > Opciones > Avanzadas) y vincular el final de cada rama con el hito de fin del "
          "proyecto."),
    _code("LOG-018", W, "Vínculos con otras planificaciones",
          "Con una planificación (p. ej. INTERNO), algunas de sus tareas están vinculadas con tareas activas de fuera "
          "de ella (otra planificación del mismo archivo), y las planificaciones no son independientes. MS Project "
          "calcula con esos vínculos las fechas, la holgura y las tareas críticas de la planificación, pero los "
          "controles de lógica (DCMA) y de avance fuera de secuencia solo cuentan los vínculos entre tareas de la "
          "planificación: una tarea cuya única sucesora es de otra planificación se informa sin sucesoras. Un "
          "resumen, que indica cuántas tareas se quedarían sin predecesoras o sin sucesoras al quitarlos, y una "
          "línea por tarea, con las tareas de otras planificaciones con que está vinculada.",
          "Si las planificaciones deben ser independientes, completar la lógica propia de cada una y quitar esos "
          "vínculos en MS Project. Si deben estar vinculadas (p. ej. una replica a la otra), silenciar LOG-018."),
    _code("LOG-019", W, "Tarea vinculada con otra planificación",
          "Desde la versión 2.6.0, el resumen y cada tarea se informan con LOG-018.",
          "Ninguna.", retired="2.6.0", by="LOG-018"),

    # --- LIB: libro ---------------------------------------------------------------------------------------------
    _code("LIB-001", I, "Cálculo del PV y el EV",
          "Recordatorio: el PV y el EV de cada fila son el BAC × el % planificado y el % ganado.",
          "Ingresar el BAC de cada fila en las celdas amarillas de la columna D.", W_G),
    _code("LIB-002", F, "No existe el libro",
          "El libro indicado (--libro o --migrar-desde) no existe en esa ruta.",
          "Revisar la ruta y el nombre del libro."),
    _code("LIB-003", F, "El libro no es .xlsx",
          "El archivo indicado no es un libro .xlsx.",
          "Indicar el libro .xlsx generado con generar."),
    _code("LIB-004", F, "No se pudo abrir el libro",
          "El archivo no se puede abrir como libro de Excel (dañado, cifrado o de otro formato).",
          "Recuperar el libro desde un respaldo."),
    _code("LIB-005", F, "El libro no fue generado por este programa",
          "Faltan las hojas ocultas _meta y _estructura.",
          "Indicar un libro generado con generar."),
    _code("LIB-006", F, "Falta un dato de la hoja _meta",
          "La hoja oculta _meta del libro no tiene un dato necesario.",
          "Recuperar el libro desde un respaldo (la hoja _meta no se debe modificar)."),
    _code("LIB-007", F, "Dato de la hoja _meta no válido",
          "Un dato de la hoja oculta _meta fue modificado.",
          "Recuperar el libro desde un respaldo (la hoja _meta no se debe modificar)."),
    _code("LIB-008", F, "Hoja _estructura no válida",
          "La hoja oculta _estructura fue modificada.",
          "Recuperar el libro desde un respaldo (la hoja _estructura no se debe modificar)."),
    _code("LIB-009", F, "Libro de una versión anterior",
          "El libro se generó con la versión 1.x, que calculaba el PV y el EV con los costos de MS Project.",
          "Usar la versión 1.x para ese libro, o generar un libro nuevo."),
    _code("LIB-010", F, "Falta una hoja del libro",
          "Una hoja que el programa lee o escribe se eliminó o se cambió de nombre.",
          "Restaurar su nombre original o recuperar el libro desde un respaldo."),
    _code("LIB-011", F, "Filas del libro movidas",
          "Una hoja mensual no tiene las filas con las que se generó el libro: se insertaron, eliminaron u ordenaron "
          "filas. actualizar ubica cada fila por su posición.",
          "Deshacer el cambio o recuperar el libro desde un respaldo. Para cambiar la EDT, se modifica MS Project y "
          "se regenera el libro con --migrar-desde."),
    _code("LIB-012", F, "Fórmula sin valor calculado",
          "Una celda de ingreso tiene una fórmula que Excel no calculó (el libro se guardó sin calcular).",
          "Abrir y guardar el libro en Excel, o reemplazar la fórmula por un número."),
    _code("LIB-013", F, "Valor no numérico",
          "Una celda de ingreso tiene un texto o un valor que no es un número.",
          "Corregir el valor de la celda indicada."),
    _code("LIB-014", F, "Número guardado como texto",
          "Una celda tiene un número guardado como texto: Excel no lo suma.",
          "Ingresarlo como número."),
    _code("LIB-015", F, "No se puede guardar el archivo",
          "El libro (o el archivo de salida) está abierto en Excel o en otro programa, o no hay permiso para "
          "escribirlo.",
          "Cerrar el archivo y volver a ejecutar el comando.", "generar, actualizar y analizar"),

    # --- MIG: --migrar-desde ------------------------------------------------------------------------------------
    _code("MIG-001", F, "Migración entre una y dos empresas",
          "El historial de EV y AC de un libro de una sola empresa no se puede repartir entre dos empresas (ni al "
          "revés).",
          "Generar el libro nuevo sin --migrar-desde y cargar los datos a mano.", "generar --migrar-desde"),
    _code("MIG-002", W, "Libro anterior de un archivo y nuevo de varios",
          "Los EV/AC se asocian por código EDT.",
          "Verificar el traslado en el libro nuevo."),
    _code("MIG-003", W, "Libro anterior de varios archivos y nuevo de uno",
          "Los EV/AC se asocian por código EDT.",
          "Verificar el traslado en el libro nuevo."),
    _code("MIG-004", W, "Archivo del libro anterior que no está en el nuevo",
          "Un archivo de [[proyectos]] del libro anterior no forma parte del nuevo: sus filas se asocian por código "
          "EDT.",
          "Verificar el traslado de esas filas."),
    _code("MIG-005", W, "Proyecto distinto",
          "El libro anterior corresponde a otro proyecto de MS Project: los EV/AC se asocian por código EDT.",
          "Verificar el traslado en el libro nuevo."),
    _code("MIG-006", W, "Nivel máximo distinto",
          "Las filas que ahora se dividen en varias no pueden repartir su historial.",
          "Cargar a mano el historial de las filas que informa MIG-014."),
    _code("MIG-007", I, "Historial sumado en el paquete",
          "El historial de una fila anterior que quedó dentro de un paquete nuevo, o de una fila -D (Otras tareas), "
          "se sumó en esa fila.",
          "Ninguna."),
    _code("MIG-008", W, "Historial asociado por código EDT",
          "La fila recibió el historial de una fila anterior con otro nombre, asociada por código EDT.",
          "Verificar que sea el mismo paquete."),
    _code("MIG-009", W, "% ganado de varias filas sin BAC",
          "La fila recibió el % ganado de varias filas anteriores sin BAC: no se pudo combinar.",
          "Revisar la hoja % EV de esa fila."),
    _code("MIG-010", I, "Métodos de medición conservados",
          "Se conservaron los métodos editados en la columna C del libro anterior.",
          "Revisarlos si el libro anterior no indica si se editaron en Excel."),
    _code("MIG-011", W, "EV/AC antes del primer mes del nuevo horizonte",
          "Se acumularon en el mes 1.",
          "Revisar mes_inicio del libro nuevo."),
    _code("MIG-012", W, "EV/AC después del último mes del nuevo horizonte",
          "Se acumularon en el último mes.",
          "Aumentar meses del libro nuevo."),
    _code("MIG-013", I, "BAC conservado",
          "Se conservó el BAC ingresado en el libro anterior.",
          "Ninguna."),
    _code("MIG-014", E, "Historial que no se pudo trasladar",
          "La tarea de la fila anterior no existe en el proyecto nuevo o se dividió en varias filas.",
          "Cargar a mano el BAC, el EV y el AC en las filas que correspondan."),
    _code("MIG-015", W, "Análisis de desviaciones sin fila en el libro nuevo",
          "El paquete ya no tiene fila: el texto del análisis se conservó, sin valores.",
          "Revisar el texto en la hoja Análisis de desviaciones."),
    _code("MIG-016", I, "Análisis de desviaciones trasladado",
          "Se trasladaron las filas de la hoja Análisis de desviaciones.",
          "Ninguna."),
    _code("MIG-017", W, "Filas sin trasladar",
          "Resumen en la consola de las filas con historial que no se pudieron trasladar (ver MIG-014).",
          "Ver la hoja Validación del libro nuevo.", "generar --migrar-desde"),
    _code("MIG-018", I, "Registro de cambios trasladado",
          "Se trasladaron las filas de la hoja Registro de cambios y la referencia de los valores informados, con "
          "la que actualizar detecta los cambios siguientes.",
          "Ninguna."),
    _code("MIG-019", E, "Valores escritos sobre fórmulas sin trasladar",
          "El libro anterior tiene valores escritos sobre celdas que deben tener fórmula (subtotales, fila TOTAL o "
          "suma de las empresas; ver ACT-024). El libro nuevo tiene la fórmula en esas celdas: los valores no se "
          "trasladan.",
          "Revisar esas celdas en el libro anterior: si son el AC o el BAC de un paquete, cargarlos en sus filas de "
          "ingreso del libro nuevo.", "generar --migrar-desde"),
    _code("MIG-020", W, "Ceros intermedios con otro criterio",
          "El libro anterior y el nuevo se generaron con distinto ignorar_ceros_intermedios: los códigos con un 0 "
          "intermedio (p. ej. 1.2.0.1) cambian de nivel. Las filas que ahora se dividen en varias no pueden repartir "
          "su historial.",
          "Cargar a mano el historial de las filas que informa MIG-014."),

    # --- DEP: depuración de costes ------------------------------------------------------------------------------
    _code("DEP-001", I, "Depuración de costes activa",
          "Recordatorio: la previsión se pega en la hoja PV-depuración y el libro de entrega se genera sin "
          "depuración.",
          "Ninguna."),
    _code("DEP-002", W, "Filas con código EDT no numérico",
          "Esas filas no pueden recibir la previsión de la hoja PV-depuración, que se asigna por código numérico.",
          "Usar códigos EDT numéricos.", "generar (hoja Validación) y validar, con depuracion_costes"),
    _code("DEP-003", W, "Códigos EDT repetidos en el libro",
          "La previsión de un código repetido se asigna a la primera fila con ese código.",
          "Corregir los códigos repetidos (ver EDT-009).", "generar (hoja Validación) y validar, con "
          "depuracion_costes"),
    _code("DEP-004", I, "Previsión no trasladada: libro sin depuración",
          "El libro nuevo no tiene depuración de costes: la hoja PV-depuración no se trasladó.",
          "Ninguna (es lo habitual en el libro de entrega).", "generar --migrar-desde"),
    _code("DEP-005", W, "Previsión no trasladada: otra empresa",
          "El libro anterior tiene la previsión de una empresa y el nuevo depura otra.",
          "Pegar la previsión de nuevo."),
    _code("DEP-006", I, "Previsión trasladada",
          "Se conservó el contenido de la hoja PV-depuración: el plan de costes y los recargos de la tabla Recargos. "
          "El plan de una hoja de la versión 2.5 o anterior, que empezaba en la celda A8, pasa a la celda A12.",
          "Ninguna."),
    _code("DEP-007", I, "Previsión pasada al formato con EDT interna",
          "La hoja PV-depuración del libro anterior no tenía las columnas de EDT interna: los códigos numéricos "
          "quedaron como EDT cliente y los demás, como EDT interna, y se agregó la columna EAC.",
          "Revisar la hoja PV-depuración."),
    _code("DEP-008", (W, I), "Meses de la previsión desplazados",
          "Cambió el mes 1 del horizonte: los meses de la hoja PV-depuración se desplazaron para conservar su fecha. "
          "Es una advertencia si quedaron importes fuera de las columnas.",
          "Revisar los importes que quedaron fuera, si los hay."),
    _code("DEP-009", I, "Previsión pasada al formato con la columna EAC",
          "La hoja PV-depuración del libro anterior (versión 2.4) no tenía la columna EAC: se agregó con la suma de "
          "los meses de cada línea, y los meses empiezan en la columna G.",
          "Ninguna."),
    _code("DEP-010", (I, W), "Reparto del plan de costes conservado",
          "Hasta la versión 2.5, --migrar-desde conservaba en el libro nuevo el % planificado de las filas procesadas "
          "con procesar-costes. Desde la versión 2.6.0, procesar-costes no modifica el % planificado y --migrar-desde "
          "lo vuelve a calcular con MS Project (ver DEP-020).",
          "Ninguna.", W_MIG, retired="2.6.0"),
    _code("DEP-011", W, "Reparto del plan de costes no combinado",
          "Hasta la versión 2.5, se informaba la fila nueva que reunía varias filas procesadas con procesar-costes. "
          "Desde la versión 2.6.0, el % planificado de todas las filas se calcula con MS Project (ver DEP-020).",
          "Ninguna.", W_MIG, retired="2.6.0"),
    _code("DEP-012", F, "Libro sin depuración de costes",
          "procesar-costes necesita un libro generado con depuracion_costes = true, con el plan de costes pegado en la "
          "hoja PV-depuración.",
          "Generar el libro con depuracion_costes = true (y --migrar-desde para conservar el BAC).", W_PROC),
    _code("DEP-013", F, "Plan de costes sin líneas leídas",
          "Ninguna línea de EDT cliente de la hoja PV-depuración tiene el código de una fila de ingreso del libro: la "
          "hoja está vacía o sus códigos no son los del libro.",
          "Escribir el plan de costes desde la celda A12 y revisar la columna 'Asignado a'.", W_PROC),
    _code("DEP-014", W, "Líneas de EDT cliente sin leer",
          "Líneas de EDT cliente con importes que no se leen: su código no es el de una fila del libro, o es el de un "
          "resumen del libro y ninguna línea de un código inferior se lee. Su importe no llega a ninguna fila "
          "«(coste)».",
          "Corregir los códigos en el plan de costes, agregar las filas al libro (p. ej. con lista_edt) o detallar "
          "el resumen con las EDT cliente de sus filas.", W_PROC),
    _code("DEP-015", W, "Plan de costes después del horizonte",
          "El plan de costes tiene importes en meses posteriores al horizonte del libro: forman parte del EAC de la "
          "fila «(coste)», pero no de sus meses ni del gráfico.",
          "Ampliar el horizonte del libro (meses) o revisar el plan de costes.", W_PROC),
    _code("DEP-016", I, "Filas sin línea en el plan de costes",
          "Filas de ingreso de la empresa sin línea de EDT cliente en el plan de costes: su fila «(coste)» queda "
          "vacía.",
          "Ninguna, si la empresa no tiene costes en esas filas.", W_PROC),
    _code("DEP-017", W, "Filas sin importes en el plan de costes",
          "La línea de EDT cliente de la fila no tiene importes en los meses (p. ej. faltan los cálculos del plan): "
          "su fila «(coste)» queda en 0.",
          "Completar los cálculos del plan (o ejecutar procesar-costes --calculos) y revisar los importes.", W_PROC),
    _code("DEP-018", W, "Importes del plan de costes que no se leen",
          "Importes guardados como texto, o celdas con fórmula sin valor calculado (el libro no se guardó en Excel "
          "después de escribirlas).",
          "Escribir números; abrir y guardar el libro en Excel, o ejecutar procesar-costes --calculos.", W_PROC),
    _code("DEP-019", W, "EDT cliente repetida en el plan de costes",
          "Dos o más líneas de EDT cliente tienen el mismo código: se suman sus importes.",
          "Revisar el plan de costes: cada EDT cliente debería tener una sola línea.", W_PROC),
    _code("DEP-020", I, "% planificado de MS Project recuperado",
          "Filas que en el libro anterior tenían en la hoja % PV el reparto del plan de costes, escrito por "
          "procesar-costes en una versión anterior a la 2.6.0: en el libro nuevo, su % planificado vuelve a ser el de "
          "MS Project. El plan de costes se compara con el PV en las filas «(coste)».",
          "Ninguna.", W_MIG),
    _code("DEP-021", W, "Filas con el reparto del plan de costes",
          "El libro tiene filas cuyo % planificado es el reparto del plan de costes, escrito por procesar-costes en "
          "una versión anterior a la 2.6.0: no es el de MS Project. actualizar no compara esas filas con la línea "
          "base (ACT-047).",
          "Generar un libro nuevo con --migrar-desde: vuelve a calcular su % planificado con MS Project y conserva el "
          "BAC, el % ganado y el AC.", f"{W_ACT} y {W_PROC}"),
    _code("DEP-022", W, "Componentes de coste fuera del precio",
          "Con [depuracion_precio], líneas de componente de coste (solo la columna D) con importes cuyo texto no "
          "está en [depuracion_precio.componentes]: no entran en el precio de su EDT interna. Los textos se comparan "
          "sin distinguir mayúsculas, tildes ni espacios de más.",
          "Corregir el texto de la línea o agregarlo a [depuracion_precio.componentes] (y a calculo_precio).", W_PROC),
    _code("DEP-023", W, "Tabla Recargos incompleta",
          "Con [depuracion_precio], a la tabla Recargos de la hoja PV-depuración le falta la fila de un coeficiente "
          "que usa calculo_precio, la columna de un año con importes o el valor de un coeficiente en un año con "
          "importes (o la tabla entera, que procesar-costes --calculos agrega vacía). En el precio, los coeficientes "
          "que faltan se toman como 0.",
          "Completar la tabla Recargos (desde la celda D1): un % por coeficiente y año.", W_PROC),
    _code("DEP-024", F, "Hoja PV-depuración de una versión anterior",
          "La hoja PV-depuración tiene el formato de la versión 2.5 o anterior, con el plan desde la celda A8. Desde "
          "la versión 2.6.0, el plan empieza en la celda A12, con la tabla Recargos arriba.",
          "Generar un libro nuevo con --migrar-desde: pasa el plan a la celda A12 y conserva el BAC, el % ganado y el "
          "AC.", W_PROC),
    _code("DEP-025", W, "Recargos no trasladados",
          "La tabla Recargos del libro anterior tiene coeficientes cuyo texto no está en [depuracion_precio.recargos] "
          "de la configuración del libro nuevo: sus valores no se trasladaron.",
          "Si el coeficiente cambió de texto, escribir sus valores en la tabla Recargos del libro nuevo.", W_MIG),
    _code("DEP-026", W, "Tabla Recargos sin cálculo del precio",
          "La hoja PV-depuración tiene la tabla Recargos, pero la configuración de procesar-costes no tiene "
          "[depuracion_precio]: con --calculos, las EDT interna suman sus componentes (a coste), en lugar de su "
          "precio.",
          "Ejecutar procesar-costes con --config y la configuración de la depuración, que tiene [depuracion_precio].",
          W_PROC),

    # --- ACT: actualizar y analizar -----------------------------------------------------------------------------
    _code("ACT-001", F, "El libro es de un solo archivo",
          "El libro se generó con un solo archivo de MS Project y la configuración tiene [[proyectos]].",
          "Indicar ese archivo con --origen, con una configuración sin [[proyectos]]."),
    _code("ACT-002", F, "El libro es de varios archivos",
          "El libro se generó con varios archivos de MS Project y se indicó uno solo.",
          "Indicar los archivos en secciones [[proyectos]], con los mismos nombres que al generar el libro."),
    _code("ACT-003", F, "Los archivos no coinciden con los del libro",
          "Los nombres de [[proyectos]] no son los del libro.",
          "Indicar todos los archivos del libro, con los mismos nombres. Para agregar o quitar un archivo, generar "
          "un libro nuevo con --migrar-desde."),
    _code("ACT-004", F, "Sin fecha de estado",
          "El proyecto (o algún archivo) no tiene fecha de estado y no se indicó --fecha-corte.",
          "Indicar --fecha-corte AAAA-MM-DD o definir Proyecto > Fecha de estado en MS Project."),
    _code("ACT-005", F, "Fechas de estado distintas",
          "Con varios archivos, las fechas de estado no coinciden.",
          "Indicar --fecha-corte o definir la misma fecha de estado en todos los archivos."),
    _code("ACT-006", F, "Fecha de corte fuera del horizonte",
          "La fecha de corte no está entre el primer y el último mes del libro.",
          "Revisar la fecha, o regenerar el libro con un horizonte mayor (--migrar-desde).", W_AN),
    _code("ACT-007", F, "Valores en meses posteriores a la fecha de corte",
          "El libro tiene % ganado cargado después del mes de la fecha de corte.",
          "Borrar esos valores, o usar --forzar si se quiere corregir un mes anterior (ver ACT-021)."),
    _code("ACT-008", F, "El libro no tiene fecha de corte",
          "analizar necesita la fecha de corte de la hoja Resumen, que escribe actualizar.",
          "Ejecutar actualizar antes de analizar.", "analizar"),
    _code("ACT-009", W, "Planificación distinta de la del libro",
          "El libro se generó para otra planificación: se usa la del libro.",
          "Revisar planificacion en la configuración."),
    _code("ACT-010", W, "Proyecto distinto del libro",
          "El libro se generó para otro proyecto de MS Project: las filas se asocian por código EDT.",
          "Verificar los resultados."),
    _code("ACT-011", W, "La tarea de la fila ya no existe",
          "La fila se omite.",
          "Si cambió la EDT, regenerar el libro con --migrar-desde."),
    _code("ACT-012", W, "La tarea de la fila está inactiva",
          "La fila se omite y se conservan los valores ya cargados.",
          "Si quedó fuera del alcance, regenerar el libro con --migrar-desde."),
    _code("ACT-013", W, "No se encontró la tarea de la fila",
          "Con un proyecto distinto del libro, ninguna tarea tiene el código EDT de la fila: se omite.",
          "Revisar el código EDT de la tarea."),
    _code("ACT-014", W, "El campo Visible cambió",
          "El campo Visible de algunas tareas no coincide con las filas del libro: el libro conserva sus filas.",
          "Para aplicar el cambio, regenerar el libro con --migrar-desde."),
    _code("ACT-015", W, "Tareas nuevas sin fila en el libro",
          "Se agregaron tareas en MS Project después de generar el libro, con un código EDT (o en la lista EDT) que "
          "les daría fila propia, debajo de una tarea resumen del libro: su avance no se incluye en el % ganado. Las "
          "que quedan dentro de una fila de ingreso se informan con ACT-051.",
          "Para incluirlas, regenerar el libro con --migrar-desde."),
    _code("ACT-016", W, "lista_edt en un libro generado sin lista",
          "La configuración indica lista_edt, pero el libro se generó sin lista EDT: se mantienen sus filas.",
          "Para aplicar la lista, regenerar el libro con --migrar-desde."),
    _code("ACT-017", W, "La lista EDT cambió",
          "La lista EDT de la configuración cambió después de generar el libro: se usa la lista con la que se "
          "generó.",
          "Para aplicar los cambios, regenerar el libro con --migrar-desde."),
    _code("ACT-018", W, "Fecha de corte que no es fin de mes",
          "La fecha de corte no es el último día del mes ni su último día laborable (según el calendario del "
          "proyecto): el PV acumulado incluye el mes completo.",
          "Usar el último día del mes, o su último día laborable, como fecha de corte."),
    _code("ACT-019", W, "Fecha de estado distinta de la fecha de corte",
          "El avance registrado en MS Project debe corresponder a la fecha de corte.",
          "Usar la fecha de estado de MS Project como fecha de corte, o actualizar el avance a esa fecha."),
    _code("ACT-020", W, "Fecha de corte anterior a la última cargada",
          "Se carga un mes anterior al último cargado en el libro.",
          "Verificar que sea intencional (p. ej. una corrección)."),
    _code("ACT-021", W, "--forzar con meses posteriores cargados",
          "Los valores de los meses posteriores se calcularon con el acumulado anterior: ya no coinciden con MS "
          "Project.",
          "Revisar esos meses o volver a cargarlos con actualizar, en orden."),
    _code("ACT-022", W, "Ninguna fila tiene BAC",
          "El PV, el EV y los indicadores quedan en 0 hasta ingresar el BAC.",
          "Ingresar el BAC de cada fila en la columna D de la hoja PV."),
    _code("ACT-023", W, "Filas con % ganado sin BAC",
          "El EV de esas filas no se puede valorar.",
          "Ingresar su BAC en la columna D de la hoja PV."),
    _code("ACT-024", W, "Celdas sin la fórmula del libro",
          "Subtotales, la fila TOTAL o la suma de las empresas tienen un valor en lugar de su fórmula: los totales "
          "del libro no coinciden con los que muestra la consola.",
          "Restaurar las fórmulas (p. ej. copiándolas de una celda vecina de la misma fila) o recuperar el libro "
          "desde un respaldo."),
    _code("ACT-025", W, "Celdas de ingreso convertidas en número",
          "Celdas de ingreso con un número guardado como texto o con una fórmula se reemplazaron por su valor.",
          "Ninguna; se recomienda ingresar números.", "actualizar, analizar y procesar-costes"),
    _code("ACT-026", I, "El AC se carga a mano",
          "Recordatorio: los indicadores de costo usan el AC cargado en las hojas AC.",
          "Cargar el AC del mes."),
    _code("ACT-027", I, "% ganado del mes negativo",
          "El % ganado acumulado bajó respecto del mes anterior: el mes tiene un valor negativo (corrección).",
          "Verificar el avance registrado en MS Project.", W_ROW),
    _code("ACT-028", I, "La empresa no participa en las tareas de la fila",
          "Se usó el avance de la fila completa para el % ganado de esa empresa.",
          "Ninguna.", W_ROW),
    _code("ACT-029", I, "Hitos con pesos negativos ponderados por igual",
          "Método «Hitos ponderados»: ningún hito tiene un peso válido (pesos negativos): se ponderaron por igual.",
          "Corregir el peso de los hitos.", W_ROW),
    _code("ACT-030", I, "Hitos sin peso ponderados por igual",
          "Método «Hitos ponderados»: ningún hito tiene peso: se ponderaron por igual.",
          "Indicar el peso de los hitos, si no deben pesar igual.", W_ROW),
    _code("ACT-031", I, "Hitos con peso negativo tomados con peso 0",
          "Método «Hitos ponderados»: algunos hitos tienen peso negativo.",
          "Corregir el peso de esos hitos.", W_ROW),
    _code("ACT-032", I, "Sin hitos: % físico completado",
          "Método «Hitos ponderados» sin hitos: se usó el % físico completado.",
          "Agregar los hitos del paquete o cambiar el método.", W_ROW),
    _code("ACT-033", I, "Sin hitos ni % físico: % completado",
          "Método «Hitos ponderados» sin hitos ni % físico completado: se usó el % completado.",
          "Agregar los hitos del paquete o cambiar el método.", W_ROW),
    _code("ACT-034", I, "Método no definido: % completado",
          "La fila no tiene método de medición: se usó el % completado. Las filas con la misma observación se "
          "muestran en una sola línea, con sus códigos EDT.",
          "Indicar el método en la columna C del libro o en MS Project.", W_ROW),
    _code("ACT-035", I, "Método no reconocido: % completado",
          "El método de la fila no es uno de los admitidos: se usó el % completado (ver MET-001). Las filas con el "
          "mismo método se muestran en una sola línea, con sus códigos EDT.",
          "Corregir el método en la columna C del libro o en MS Project.", W_ROW),
    _code("ACT-036", (W, I), "Cambio de BAC",
          "El BAC de la fila cambió desde la ejecución anterior: el cambio se agregó a la hoja Registro de cambios. "
          "Es una advertencia si la fila tiene meses ya informados: su PV y su EV se recalcularon con el BAC nuevo "
          "(cambio retroactivo).",
          "Indicar el motivo y el número de cambio aprobado en la hoja Registro de cambios.", W_AN),
    _code("ACT-037", W, "Valores de meses ya informados modificados",
          "El % ganado o el AC de meses anteriores al mes de corte cambió desde la ejecución anterior, o se cargó el "
          "AC de un mes anterior al último informado: el cambio se agregó a la hoja Registro de cambios.",
          "Si es una corrección, indicar el motivo en la hoja Registro de cambios; si no, restaurar el valor "
          "anterior.", W_AN),
    _code("ACT-038", I, "Referencia del registro de cambios guardada",
          "Primera ejecución con el registro de cambios: se guardaron el BAC y los valores de los meses informados "
          "de cada fila. Los cambios se registran a partir de la ejecución siguiente.",
          "Ninguna.", W_AN),
    _code("ACT-039", W, "Avance sin costo",
          "La fila tiene EV acumulado y no tiene AC acumulado: el costo real no se cargó o se imputó a otro "
          "paquete, y el CPI no se puede calcular.",
          "Cargar el AC del paquete en la hoja AC o revisar el avance informado.", W_AN),
    _code("ACT-040", W, "Costo sin avance",
          "La fila tiene AC acumulado y no tiene EV acumulado: costo sin trabajo medido. No se revisa en los "
          "paquetes con método 0/100 o «Hitos ponderados», en los que el EV llega al terminar o al cumplirse cada "
          "hito.",
          "Revisar el avance en MS Project, el método de medición o la imputación del AC.", W_AN),
    _code("ACT-041", W, "Costo después de terminar el paquete",
          "La fila tiene AC en meses posteriores al mes en que su % ganado llegó al 100 %.",
          "Revisar si es un costo tardío del paquete (y explicarlo) o un costo de otro paquete.", W_AN),
    _code("ACT-042", I, "AC del mes de corte sin cargar",
          "Los controles del EV y el AC se revisan hasta el último mes con AC cargado (o no se revisan, si ningún "
          "mes tiene AC).",
          "Cargar el AC del mes en la hoja AC y ejecutar analizar.", W_AN),
    _code("ACT-043", W, "Análisis del mes anterior incompleto",
          "La fila del último análisis anterior al mes de corte tiene texto, pero no tiene Causa o no tiene Acción "
          "correctiva.",
          "Completar la causa y la acción correctiva en la hoja Análisis de desviaciones.", W_AN),
    _code("ACT-044", W, "Desviaciones del mes anterior sin analizar",
          "Paquetes que estaban fuera de umbral en el último análisis anterior al mes de corte y cuya fila no tiene "
          "ningún texto: se quitan de la hoja. Los que siguen fuera de umbral vuelven a listarse en el mes de corte.",
          "Analizar las desviaciones de cada mes antes del cierre siguiente.", W_AN),
    _code("ACT-045", W, "Acción vencida con el paquete fuera de umbral",
          "El paquete sigue fuera de umbral en el mes de corte y la Fecha compromiso de su acción correctiva es "
          "anterior a la fecha de corte: la acción no se cumplió o no fue eficaz.",
          "Revisar la acción correctiva y su fecha compromiso en la hoja Análisis de desviaciones.", W_AN),
    _code("ACT-046", W, "Meses sin informe",
          "Entre el primer mes con % ganado y el mes de corte hay meses sin % ganado cargado: su avance quedó en el "
          "mes cargado siguiente. La norma EIA-748-E pide la información al menos cada mes.",
          "Ejecutar actualizar al cierre de cada mes, con la fecha de estado de ese mes, aunque el informe sea "
          "bimestral. Si la frecuencia está acordada con el cliente, silenciar ACT-046.", W_AN),
    _code("ACT-047", W, "Línea base de MS Project distinta de la del libro",
          "El % planificado que resulta hoy de MS Project no coincide con el de las hojas % PV, que calculó generar: "
          "se guardó de nuevo la línea base, o cambiaron las tareas o el calendario de la fila. El libro sigue con la "
          "línea base anterior. No se comparan las filas con tareas sin fechas de línea base.",
          "Si el cambio de línea base está aprobado, generar un libro nuevo con --migrar-desde (conserva el BAC, el % "
          "ganado y el AC); si no, restaurar la línea base en MS Project.", W_ACT),
    _code("ACT-048", W, "EAC poco realista",
          "El TCPI(BAC) supera al CPI en más de 0,10, o el AC ya alcanzó el BAC con trabajo pendiente: terminar "
          "dentro del BAC exige una eficiencia de costo muy superior a la lograda. Se calcula con el proyecto (cada "
          "empresa en multi-empresa) en el último mes con AC cargado, desde el 20 % completado: antes, el CPI "
          "acumulado todavía no es representativo.",
          "Revisar el EAC y, si el desvío se confirma, evaluar una línea base por encima del objetivo (OTB) con el "
          "cliente.", W_AN),
    _code("ACT-049", W, "Mucho BAC en filas LOE",
          "El % del BAC total en filas con método LOE supera [validacion] loe_maximo: su EV es igual a su PV y acerca "
          "el SPI a 1 aunque el trabajo medible se atrase. La norma EIA-748-E pide mantener el LOE al mínimo.",
          "Medir con hitos o % completado el trabajo que lo permita, o ajustar loe_maximo si el contrato lo admite.",
          W_AN),
    _code("ACT-050", W, "Ceros intermedios con otro criterio que el libro",
          "ignorar_ceros_intermedios de la configuración no coincide con el del libro, y algunas tareas con un 0 "
          "intermedio en el código (p. ej. 1.2.0.1) tendrían otra fila: el libro conserva sus filas.",
          "Para aplicar la configuración, regenerar el libro con --migrar-desde; si no, indicar en la configuración "
          "el valor con que se generó el libro."),
    _code("ACT-051", W, "Tareas nuevas que se suman en una fila",
          "Se agregaron tareas en MS Project después de generar el libro, con un código EDT (o en la lista EDT) que "
          "les daría fila propia, dentro de una fila de ingreso: su avance se suma en el % ganado de esa fila, pero "
          "su duración no forma parte del % planificado del libro (ver ACT-047).",
          "Para darles fila propia, con su % planificado, regenerar el libro con --migrar-desde."),

    # --- FIC: fichas de tareas ----------------------------------------------------------------------------------
    _code("FIC-001", W, "No hay fichas",
          "Ninguna tarea tiene código EDT.",
          "Revisar el campo EDT configurado ([campos] edt) y la planificación elegida."),
    _code("FIC-002", I, "Vínculos con tareas sin ficha",
          "Los vínculos con tareas de otra planificación, inactivas, con Visible = No, fuera de la lista EDT o de "
          "otro proyecto no se muestran.",
          "Ninguna."),
    _code("FIC-003", W, "Fichas sin línea base",
          "Se usaron las fechas y la duración del programa actual (se indica en la ficha).",
          "Guardar la línea base de esas tareas."),
    _code("FIC-004", W, "El archivo tiene errores",
          "El cronograma tiene errores (por ejemplo, códigos EDT duplicados) que pueden afectar las fichas.",
          "Revisarlo con validar."),

    # --- CLI: planificación para el cliente (generar-cliente) ---------------------------------------------------
    _code("CLI-001", F, "La planificación no es de primer nivel",
          "generar-cliente parte de una tarea de primer nivel del esquema (la planificación, por ejemplo INTERNO) y "
          "sus subtareas; la planificación indicada está debajo de otra tarea.",
          "Indicar en planificacion una tarea de primer nivel, o dejarla vacía para usar todo el proyecto."),
    _code("CLI-002", F, "Ninguna tarea se muestra al cliente",
          "Ninguna tarea de la planificación tiene Sí en el campo de [cliente] campo (en ella y en todas sus tareas "
          "superiores).",
          "Revisar el campo configurado en [cliente] campo y marcar con Sí las tareas que debe ver el cliente, "
          "empezando por las de primer nivel de la planificación."),
    _code("CLI-003", W, "Valor no reconocido en el campo del cliente",
          "El campo de [cliente] campo tiene un valor que no es Sí ni No (por ejemplo, un error de escritura): la "
          "tarea no se muestra.",
          "Corregir el valor (Sí o No). El informe de control indica las tareas.", W_CLI_V),
    _code("CLI-004", W, "Tarea con Sí debajo de una tarea sin Sí",
          "La tarea tiene Sí, pero una tarea superior no: no se muestra, con sus subtareas. Una tarea solo se muestra "
          "si ella y todas sus tareas superiores tienen Sí.",
          "Si se debe mostrar, poner Sí también en sus tareas superiores; si no, quitarle el Sí.", W_CLI_V),
    _code("CLI-005", I, "Resúmenes que pasan a ser tareas",
          "No se muestra ninguna subtarea del resumen: en el archivo del cliente pasa a ser una tarea con sus fechas, "
          "su duración, su avance y sus líneas base, y recibe los vínculos de sus subtareas.",
          "Ninguna, si el cliente no debe ver ese detalle. El informe de control los lista."),
    _code("CLI-006", I, "Vínculos nuevos en la planificación del cliente",
          "Vínculos que reemplazan a los que pasan por tareas ocultas (puentes) o que pasan de una subtarea a su "
          "resumen convertido. Si un vínculo retrasaría a su sucesora respecto de la planificación interna, su "
          "posición se ajusta a las fechas actuales.",
          "Ninguna. El informe de control indica cada vínculo, su tipo y su posición."),
    _code("CLI-007", W, "Vínculos que no se crean",
          "El vínculo puente uniría un resumen con una de sus subtareas, o cerraría un ciclo con los vínculos que ya "
          "tienen las tareas: MS Project no lo admite.",
          "Revisar en el informe de control la lógica de esas tareas en la planificación interna."),
    _code("CLI-008", I, "Resúmenes que cambian de fechas",
          "Una subtarea que no se muestra fija el comienzo o el fin del resumen: en el archivo del cliente, el resumen "
          "va de la primera a la última de las subtareas que se muestran.",
          "Ninguna, si es lo esperado. El informe de control indica las fechas."),
    _code("CLI-009", F, "Varios archivos en generar-cliente",
          "La configuración tiene secciones [[proyectos]]: generar-cliente trabaja con un solo archivo de MS Project.",
          "Usar una configuración sin [[proyectos]] e indicar el archivo con --origen."),
    _code("CLI-010", F, "La generación necesita el archivo .mpp",
          "El archivo del cliente se genera con MS Project a partir de una copia del .mpp de la planificación "
          "interna. Con un XML o con 'activo' solo se puede simular.",
          "Indicar en --origen el archivo .mpp, o usar --simular para ver los cambios sin generar el archivo."),
    _code("CLI-011", F, "El destino es el archivo de origen",
          "--destino indica el mismo archivo que --origen: la planificación interna nunca se modifica.",
          "Indicar otro archivo en --destino (por defecto, <origen>_cliente.mpp)."),
    _code("CLI-012", W, "El origen está abierto en MS Project",
          "El archivo de la planificación interna está abierto en MS Project: se copia la versión guardada en el "
          "disco, sin los cambios que no se guardaron.",
          "Guardar el archivo en MS Project y volver a generar el archivo del cliente si había cambios sin guardar."),
    _code("CLI-013", I, "Fechas fijadas en el archivo del cliente",
          "Sin las tareas ocultas, MS Project adelantaba estas tareas: su comienzo se fijó con la delimitación «No "
          "comenzar antes del» en la fecha de la planificación interna ([cliente] fijar_fechas).",
          "Ninguna. El informe de control lista las tareas."),
    _code("CLI-014", W, "Fechas distintas de la planificación interna",
          "Después de recalcular, estas tareas del archivo del cliente no tienen las fechas de la planificación "
          "interna y no se pudieron fijar: están comenzadas o terminadas, tienen otra delimitación, se atrasan, o "
          "fijar_fechas = false.",
          "Revisar en el informe de control el motivo de cada tarea y, si hace falta, la lógica de la planificación "
          "interna."),
    _code("CLI-015", W, "Vínculos que MS Project no aceptó",
          "MS Project rechazó un vínculo nuevo (por ejemplo, porque cerraría un ciclo con un resumen).",
          "Revisar en el informe de control el error de cada vínculo. Las fechas de sus tareas se comparan igual."),
    _code("CLI-016", I, "Vínculos creados por MS Project al borrar tareas",
          "Al borrar las tareas ocultas, MS Project vinculó las tareas de alrededor (vinculación automática). Esos "
          "vínculos no son de la planificación interna y se quitaron.",
          "Ninguna."),
    _code("CLI-017", W, "No se pudo quitar el nombre del campo del cliente",
          "Los valores del campo se vaciaron en todas las tareas, pero MS Project no permitió quitar su nombre "
          "(alias): el cliente puede ver el campo, vacío, en la lista de campos.",
          "Quitar el nombre a mano en el archivo del cliente (Proyecto > Campos personalizados) o revisar el "
          "registro de diagnóstico."),
    _code("CLI-018", F, "Error de MS Project al generar el archivo del cliente",
          "MS Project devolvió un error en uno de los pasos (borrar tareas, vincular, recalcular, guardar...) o el "
          "resultado no es el esperado (por ejemplo, quedan tareas que no se muestran). No se genera el archivo del "
          "cliente; la planificación interna no se modifica.",
          "Volver a ejecutar con --registro RUTA y revisar el paso y el error en el registro de diagnóstico."),
    _code("CLI-019", W, "Resúmenes convertidos con datos sin restaurar",
          "Al pasar de resumen a tarea, MS Project cambió la duración, el avance o una línea base del resumen y no "
          "aceptó el valor original.",
          "Revisar esas tareas en el archivo del cliente (el informe de control indica el dato)."),
    _code("CLI-020", W, "Resúmenes con otro % completado",
          "MS Project calcula el % completado de un resumen con la duración de sus tareas de detalle. En el archivo del "
          "cliente no están las tareas ocultas, y su avance no cuenta; un resumen convertido pesa por su duración, no "
          "por la de sus subtareas: el % del resumen cambia.",
          "Revisar la lista del informe de control. Si el cliente debe ver el avance de esos resúmenes, se debe "
          "mostrar el detalle que lo explica (marcar con Sí las tareas que lo forman)."),
    _code("CLI-021", W, "Tareas que no conservan su % completado",
          "Una tarea de detalle o un resumen convertido tiene en el archivo del cliente otro % completado que en la "
          "planificación interna (por ejemplo, porque MS Project no aceptó el avance al convertir el resumen).",
          "Revisar esas tareas en el archivo del cliente y el registro de diagnóstico (--registro)."),
    _code("CLI-022", I, "Tareas que quedan como «Otras tareas»",
          "Con [cliente] actividad_interna = true, las tareas que no se muestran quedan en el archivo del cliente como "
          "«Otras tareas»: sin nombre, código EDT, notas, hipervínculo ni campos de texto (salvo los que usa el libro "
          "EVM), con sus fechas, su duración, su avance, sus líneas base y sus vínculos. Así el % completado de cada "
          "tarea, de los resúmenes y del proyecto, y el libro EVM, son los de la planificación interna. Las tareas "
          "inactivas que no se muestran se quitan: no cuentan en ningún cálculo.",
          "Ninguna. Con filtrar_vista = true, el archivo se abre con un filtro que oculta las «Otras tareas»."),
    _code("CLI-023", W, "% completado del proyecto distinto",
          "El % completado del proyecto, calculado como MS Project (con la duración de las tareas de detalle), no es "
          "el mismo en el archivo del cliente: con actividad_interna = false se borran las tareas que no se muestran, "
          "y su avance deja de contar.",
          "Usar actividad_interna = true (por defecto) para que coincida, o mostrar las tareas que explican el avance "
          "(Sí en el campo del cliente). Si ocurre con actividad_interna = true, revisar el registro de diagnóstico."),
    _code("CLI-024", W, "Filas del libro EVM con otras tareas o valores",
          "Una fila de ingreso del libro EVM del cliente no tiene las mismas tareas de detalle, el mismo % planificado "
          "de cada mes o el mismo % ganado que en el libro de la planificación interna (o solo está en uno de los dos "
          "libros). Ocurre al borrar las tareas que no se muestran (actividad_interna = false). Con «Otras tareas», "
          "ocurre si no se muestra ningún elemento EDT de un resumen (su fila reúne las de esos elementos) o si no se "
          "muestra un elemento EDT que no depende de ninguno que se muestre: sin su código, sus tareas no forman parte "
          "de ninguna fila del libro del cliente.",
          "Revisar la lista del informe de control, que indica en qué fila del otro libro están las tareas de cada "
          "fila. Para que coincidan, usar actividad_interna = true (por defecto) y mostrar esos elementos EDT (Sí en "
          "el campo del cliente) o ubicarlos debajo de un elemento que se muestre."),
    _code("CLI-025", W, "Filas -D en el libro EVM del cliente",
          "El libro EVM del cliente tiene filas -D («Otras tareas de …»): tareas sin código EDT dentro de un elemento "
          "EDT que tiene filas propias, por ejemplo las de un elemento que no se muestra, cuyo código se quitó. Sus "
          "valores son los mismos, pero el cliente ve esas filas en su libro.",
          "Para que no aparezcan, mostrar ese elemento EDT (Sí en el campo del cliente) o reorganizar la EDT para que "
          "las tareas internas no queden junto a elementos EDT que se muestran. El informe de control indica la fila "
          "interna de cada fila -D."),
    _code("CLI-026", W, "Tareas con datos distintos en el archivo del cliente",
          "Con actividad_interna = true, cada tarea del archivo del cliente debe tener las fechas, el avance, la "
          "duración y la línea base de la planificación interna; estas tareas no los tienen después de guardar el "
          "archivo.",
          "Revisar la lista del informe de control y el registro de diagnóstico (--registro)."),
    _code("CLI-027", W, "Datos de las «Otras tareas» que no se quitaron",
          "MS Project no dejó vaciar un campo de las «Otras tareas» (los campos calculados con una fórmula se informan "
          "aparte, con CLI-030), o el campo EDT configurado es el de MS Project (EDT o número de esquema) y su código "
          "no se puede quitar: el cliente puede ver esos datos.",
          "Revisar el campo en el archivo del cliente. Para quitar los códigos, usar un campo EDT personalizado "
          "([campos] edt, por ejemplo un campo de texto)."),
    _code("CLI-028", W, "No se pudo crear el filtro de la vista",
          "Con [cliente] filtrar_vista = true, MS Project no aceptó crear o aplicar el filtro que oculta las «Otras "
          "tareas»: el archivo del cliente se abre con todas las tareas.",
          "Revisar el registro de diagnóstico (--registro). El % completado y el libro EVM no cambian."),
    _code("CLI-029", W, "No se pudo comparar el libro EVM",
          "El libro EVM de la planificación interna o el del cliente no se pudo calcular para compararlos (por ejemplo, "
          "por un error de configuración del libro).",
          "Revisar el error indicado: es el mismo que daría generar con esa configuración."),
    _code("CLI-030", I, "Campos calculados en las «Otras tareas»",
          "Un campo personalizado de texto tiene una fórmula: MS Project calcula su valor en cada tarea y no deja "
          "vaciarlo. En las «Otras tareas», el valor se calcula con los datos que se conservan (fechas, avance, líneas "
          "base...).",
          "Nada, si la fórmula solo usa esos datos. Si usa un campo con información interna, el cliente puede ver el "
          "resultado: revisar la fórmula en MS Project."),
    _code("CLI-031", W, "Tareas sin valor debajo de una tarea que se muestra",
          "La tarea no tiene valor en el campo del cliente y su tarea superior se muestra: no se muestra (con "
          "actividad_interna, queda como «Otras tareas»), con sus subtareas. Lo habitual es una tarea nueva que no se "
          "marcó. Se indica la tarea de más arriba de cada rama; las inactivas no se incluyen.",
          "Si el cliente la debe ver, poner Sí. Si no, poner No: así queda decidido y no se vuelve a avisar.",
          W_CLI_V),
    _code("CLI-032", I, "Calendarios y macros quitados del archivo del cliente",
          "El archivo del cliente es una copia del .mpp. Se quitaron los calendarios que no usa ninguna tarea ni "
          "recurso, ni el proyecto (salvo el de MS Project para las líneas base de Project 98), y los módulos de "
          "macros. La lista está en el informe de control.",
          "Ninguna."),
    _code("CLI-033", W, "Calendarios o macros que no se pudieron quitar",
          "MS Project no dejó quitar un calendario sin uso, o el archivo del cliente guardado todavía tiene módulos "
          "de macros (se comprueba leyendo el archivo). Las macros se quitan con el proyecto VBA o, si MS Project no "
          "da acceso a él (Centro de confianza), con el Organizador.",
          "Quitarlos en el archivo del cliente con el Organizador de MS Project (Archivo > Información > "
          "Organizador), o permitir en el Centro de confianza de MS Project el acceso al modelo de objetos de "
          "proyectos de VBA y volver a generar el archivo."),

    # --- SIS: errores del sistema -------------------------------------------------------------------------------
    _code("SIS-001", F, "Error inesperado",
          "Un error no previsto del programa (código de salida 3).",
          "Volver a ejecutar con -v para ver el detalle técnico e informar el problema."),
    _code("SIS-002", F, "No se pudo leer o escribir un archivo",
          "Error del sistema de archivos: disco lleno, unidad de red no disponible, ruta no válida...",
          "Revisar el espacio en disco, la ruta y los permisos."),
    _code("SIS-003", F, "Error sin código específico",
          "Un error previsto (archivo, permisos o datos) sin un código propio en este catálogo.",
          "Leer el mensaje; si no es claro, volver a ejecutar con -v e informar el problema."),
]

del F, E, W, I

CATALOG: dict[str, Code] = {c.id: c for c in _CODES}
# El resumen de las observaciones silenciadas y los avisos del registro de diagnóstico (no son observaciones)
UNSILENCEABLE = {"CFG-049", "LEC-022"}
SUMMARY_ID = "CFG-049"


# ----------------------------------------------------------------------------
# Silenciar ([validacion] silenciar y --silenciar)
# ----------------------------------------------------------------------------
def _norm_code(code: str) -> str:
    """Código EDT para comparar: sin espacios, en mayúsculas y sin los ceros finales (1.2.0 = 1.2)."""
    text = (code or "").strip().upper()
    head, dash, tail = text.partition("-")
    parts = head.split(".")
    if parts and all(p.isdigit() for p in parts):
        while len(parts) > 1 and parts[-1].strip("0") == "":
            parts.pop()
        head = ".".join(str(int(p)) for p in parts)
    return head + dash + tail


@dataclass
class Silence:
    """Códigos silenciados: {código: None (todas sus observaciones) o los códigos EDT silenciados}. Cuenta las
    observaciones que quita, para el resumen (CFG-049)."""
    rules: dict[str, set[str] | None] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)       # avisos de los códigos no válidos (CFG-047, CFG-048)
    hidden: Counter = field(default_factory=Counter)

    def matches(self, issue: Issue) -> bool:
        if issue.id not in self.rules:
            return False
        codes = self.rules[issue.id]
        return codes is None or _norm_code(issue.code) in codes

    def filter(self, issues: list[Issue]) -> list[Issue]:
        """Las observaciones sin las silenciadas."""
        out = []
        for i in issues:
            if self.matches(i):
                self.hidden[i.id] += 1
            else:
                out.append(i)
        return out

    def summary(self) -> Issue | None:
        """Resumen de las observaciones silenciadas hasta ahora (None si no hay)."""
        if not self.hidden:
            return None
        total = sum(self.hidden.values())
        what = "observación silenciada" if total == 1 else "observaciones silenciadas"
        detail = ", ".join(f"{k} ({n})" for k, n in sorted(self.hidden.items()))
        return Issue(SUMMARY_ID, message=f"{total} {what} ([validacion] silenciar o --silenciar): {detail}.")

    def apply(self, issues: list[Issue]) -> list[Issue]:
        """Las observaciones sin las silenciadas, con los avisos de silenciar al principio y el resumen al final.
        Se llama una vez por comando (después de filtrar con filter las que se muestran aparte)."""
        out = self.filter(self.issues + list(issues))
        summary = self.summary()
        return out + ([summary] if summary else [])


def parse_silence(entries: list[str]) -> Silence:
    """Reglas de silenciar: 'FEC-003' (todas sus observaciones) o 'FEC-003:1.2.3' (solo las de ese código EDT). Los
    códigos desconocidos o retirados se informan con CFG-047, y los que no se pueden silenciar, con CFG-048."""
    out = Silence()
    for raw in entries:
        text = str(raw).strip()
        if not text:
            continue
        head, sep, wbs = text.partition(":")
        head = head.strip().upper()
        entry = CATALOG.get(head) if ID_PATTERN.fullmatch(head) else None
        if entry is None or (sep and not wbs.strip()):
            out.issues.append(Issue("CFG-047", message=f"'{text}' no es un código del catálogo (ÁREA-NNN o "
                                                       "ÁREA-NNN:EDT, ver docs/codigos.md): se ignora."))
            continue
        if entry.retired:
            now = f" (desde entonces se informa con {entry.by})" if entry.by else ""
            out.issues.append(Issue("CFG-047", message=f"'{text}': {entry.id} es un código retirado en la versión "
                                                       f"{entry.retired}{now}: se ignora."))
            continue
        if not entry.silenceable:
            why = ("detiene el programa" if SEV_FATAL in entry.severities
                   else "es el resumen de las observaciones silenciadas" if entry.id == SUMMARY_ID
                   else "se muestra en el registro de diagnóstico")
            out.issues.append(Issue("CFG-048", message=f"{entry.id} no se puede silenciar: {why}."))
            continue
        if not sep:
            out.rules[entry.id] = None
        elif out.rules.get(entry.id, set()) is not None:
            out.rules.setdefault(entry.id, set()).add(_norm_code(wbs))
    return out
