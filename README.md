# Exportador EVM desde MS Project

Programa en Python que lee un cronograma de MS Project y genera un libro Excel para
gestionar el **Valor Ganado (EVM)**:

- **PV mensual** de cada elemento de la EDT (hasta nivel 3 del código EDT: `2.0.0`, `2.1.0`,
  `2.1.1`) durante 120 meses (10 años).
- Selección de la **planificación** a exportar cuando el proyecto tiene varias (por ejemplo
  `INTERNO` y `CLIENTE`).
- Hojas para cargar **EV** (valor ganado) y **AC** (costo real) mes a mes.
- **Cuadro de Mando** (primera hoja) con el mes de reporte seleccionable: BAC, PV, EV y AC acumulados,
  desviaciones (CV, SV), índices (CPI, SPI), pronósticos (ETC, EAC, VAC, TCPI) y tres gráficos nativos
  de Excel: evolución EVM (curva S), evolución de CPI y SPI con zonas de alarma, y CPI frente a SPI por
  elemento EDT.
- **Indicadores** (SV, CV, SPI, CPI, EAC, VAC, TCPI y Cronograma Ganado) a la fecha de corte,
  calculados con fórmulas de Excel, para el proyecto y para cada elemento de la EDT. Todo se actualiza
  solo al cargar datos o cambiar la fecha de corte.
- Un comando que **calcula EV y AC desde MS Project** a la fecha de corte, según el método de
  medición de cada fila, y los carga en el mes correspondiente.
- Opcional: **multi-empresa**, con el PV, el EV y el AC de dos empresas en hojas separadas y el total
  en las hojas `PV`, `EV` y `AC` (sección 9).

El formato de la hoja `PV` es el siguiente:

| Columna | Contenido |
|---|---|
| A | Código EDT (`1.0.0`, `1.1.0`, `1.2.3`, ...) |
| B | Actividad (nombre de la tarea) |
| C | Método de medición (lista desplegable) |
| D | BAC: presupuesto (costo de línea base) |
| E | Fecha inicio Plan (PV): inicio de la línea base |
| F | Fecha fin Plan (PV): fin de la línea base |
| G ... DV | PV de cada mes, del mes 1 al mes 120 |

Los montos se expresan por defecto en **euros (EUR)**.

---

## Contenido

1. [Requisitos e instalación](#1-requisitos-e-instalación)
2. [Uso rápido](#2-uso-rápido)
3. [Ciclo mensual recomendado](#3-ciclo-mensual-recomendado)
4. [Comandos y opciones](#4-comandos-y-opciones)
5. [Archivo de configuración](#5-archivo-de-configuración)
6. [El libro Excel](#6-el-libro-excel)
7. [Cómo se calcula el PV](#7-cómo-se-calcula-el-pv)
8. [Cómo se calculan EV y AC](#8-cómo-se-calculan-ev-y-ac)
9. [Multi-empresa (opcional)](#9-multi-empresa-opcional)
10. [Preparación del proyecto en MS Project](#10-preparación-del-proyecto-en-ms-project)
11. [Validaciones](#11-validaciones)
12. [Limitaciones conocidas](#12-limitaciones-conocidas)
13. [Solución de problemas](#13-solución-de-problemas)
14. [Pruebas y archivo de ejemplo](#14-pruebas-y-archivo-de-ejemplo)
15. [Flujo de trabajo del repositorio](#15-flujo-de-trabajo-del-repositorio)
16. [Licencia](#16-licencia)

---

## 1. Requisitos e instalación

- Windows con **Python 3.11 o posterior** (probado con Python 3.13 y 3.14.8).
- Para leer archivos **`.mpp`**: **MS Project de escritorio** instalado en la misma PC: Standard o
  Professional 2016 o posterior, incluida la versión de suscripción que Windows muestra como
  *Microsoft Project Online Desktop Client*. El programa lo controla por automatización COM.
- Sin MS Project en la PC, se debe guardar el proyecto como **XML** desde MS Project
  (*Archivo > Guardar como > Tipo: Formato XML*) y usar ese archivo.

```powershell
git clone https://github.com/MarceloSalazar/evm-automation.git
cd evm-automation
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Si PowerShell no permite ejecutar `Activate.ps1` ("la ejecución de scripts está deshabilitada"), se
debe habilitar una vez para el usuario con `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.
Otra opción es no activar el entorno y usar `.venv\Scripts\python.exe` en lugar de `python` en todos
los comandos. Si hay varias versiones de Python instaladas, el entorno se puede crear con
`py -3.14 -m venv .venv`.

Dependencias: `openpyxl` (libro Excel) y `pywin32` 311 o posterior (COM con MS Project; la versión 311
es la primera compatible con Python 3.14).

---

## 2. Uso rápido

```powershell
# 1) Crear el libro con el PV mensual (desde el .mpp o desde el XML)
python evm_export.py generar --origen "C:\Programa\Programa.mpp" --config config.toml --salida EVM_Programa.xlsx

# 2) Cada mes: calcular EV y AC a la fecha de corte y cargarlos en el libro
python evm_export.py actualizar --origen "C:\Programa\Programa.mpp" --config config.toml --libro EVM_Programa.xlsx --fecha-corte 2026-10-31
```

Los indicadores y los gráficos se consultan en las hojas `Cuadro de Mando`, `Resumen` e
`Indicadores EDT` del libro.

Para probar el programa sin MS Project se incluye un proyecto de ejemplo en XML, con las
planificaciones `INTERNO` y `CLIENTE` (ver sección 14):

```powershell
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --planificacion INTERNO --salida ejemplo.xlsx
```

---

## 3. Ciclo mensual recomendado

**Una vez, al inicio (o al aprobar una nueva línea base):**

1. Se deben cargar los costos y los métodos de medición en MS Project y guardar la línea base
   (sección 10).
2. Se ejecuta `generar` y se revisa la hoja **Validación**; lo que corresponda se corrige en MS Project.
3. Se revisa o completa el **método de medición** de cada fila (columna C de la hoja `PV`) si no se
   definió en MS Project.

**Cada mes (al cierre):**

1. Se actualiza en MS Project el avance de las tareas (% completado, % físico completado, fechas
   reales, hitos cumplidos) y, si se registra allí, el costo real.
2. Se define la **fecha de estado** (*Proyecto > Fecha de estado*) en el **último día del mes**.
3. Se ejecuta `actualizar`. El programa:
   - calcula el **EV acumulado** de cada fila según su método, y el **AC acumulado** (CRTR de
     MS Project a la fecha de estado);
   - escribe en el mes de la fecha de corte la diferencia con lo cargado en meses anteriores
     (valor **del mes**, no acumulado);
   - actualiza la fecha de corte de la hoja `Resumen` y guarda una **copia de respaldo** del libro.
4. Si el AC proviene del ERP o de contabilidad, se ejecuta `actualizar --solo ev` y el AC del mes se
   carga a mano en la hoja `AC` (celdas amarillas).
5. Al abrir el libro, `Cuadro de Mando`, `Resumen` e `Indicadores EDT` se recalculan solos. El
   `Cuadro de Mando` muestra el mes de corte; para ver un mes anterior se elige en su desplegable.

> Aunque el informe sea bimestral, se recomienda registrar el estado **todos los meses**: así la curva
> S y los métodos 0/100 y 50/50 se miden con el período mensual.

---

## 4. Comandos y opciones

La ayuda se obtiene con `python evm_export.py --help` y `python evm_export.py <comando> --help`.
Para ver mensajes de diagnóstico se agrega `-v` **antes** del comando:
`python evm_export.py -v generar ...`

### `generar`: crea el libro con el PV

| Opción | Descripción |
|---|---|
| `--origen` | Archivo `.mpp` (requiere MS Project), `.xml` de MS Project, o `activo` para usar el proyecto abierto en MS Project. **Obligatoria.** |
| `--salida` | Libro `.xlsx` a crear. Por defecto `<origen>_EVM.xlsx`. |
| `--config` | Archivo de configuración TOML (ver `config.ejemplo.toml`). |
| `--mes-inicio AAAA-MM` | Primer mes del horizonte (mes 1). Por defecto, el mes en que comienza la línea base. |
| `--meses N` | Cantidad de meses (columnas). Por defecto 120. |
| `--planificacion CODIGO` | Código EDT de la planificación a exportar (`INTERNO`, `CLIENTE`...). Reemplaza el valor del archivo de configuración; `""` exporta todas. |
| `--linea-base N` | 0 = Línea base, 1..10 = Línea base N. Por defecto, la que el proyecto usa para el valor acumulado. |
| `--nivel-max N` | Nivel máximo del código EDT. Por defecto 3. |
| `--motor auto\|com\|xml` | Forma de lectura. `auto` decide por la extensión. |
| `--migrar-desde libro.xlsx` | Copia al libro nuevo los EV/AC ya cargados, los métodos editados en la columna C y la fecha de corte de un libro anterior (útil al cambiar la línea base o la EDT). Ver sección 6. |
| `--sobrescribir` | Reemplaza el libro de salida si ya existe. |
| `--visible` | Muestra MS Project mientras se lee (solo COM). |

### `actualizar`: carga EV y AC del mes

| Opción | Descripción |
|---|---|
| `--origen` | Proyecto con el avance a la fecha de corte (`.mpp`, `.xml` o `activo`). **Obligatoria.** |
| `--libro` | Libro creado con `generar`. **Obligatoria.** |
| `--fecha-corte AAAA-MM-DD` | Fecha de corte. Por defecto, la fecha de estado del proyecto; si el proyecto no tiene fecha de estado, es obligatoria. Con `.mpp`, el programa fija esa fecha de estado en la copia abierta (no guarda el archivo). |
| `--solo ev\|ac` | Actualiza solo EV o solo AC. En multi-empresa, `--solo ac` requiere `ac_desde_project = true` (sección 9). |
| `--simular` | Muestra el EV y AC calculados para cada fila y los indicadores resultantes, sin modificar el libro. |
| `--forzar` | Permite actualizar aunque haya valores en meses posteriores. |
| `--sin-respaldo` | No crea la copia `<libro>.respaldo-AAAAMMDD-HHMMSS.xlsx` (si ya existe una con ese nombre, se agrega `-2`, `-3`...). |
| `--config`, `--motor`, `--linea-base`, `--nivel-max`, `--visible` | Igual que en `generar`. |
| `--planificacion` | No se aplica: `actualizar` usa siempre la planificación con la que se generó el libro (y lo advierte si la configuración indica otra). |

Repetir `actualizar` para el mismo mes **reemplaza** el valor de ese mes; no lo duplica.

`actualizar` también controla la consistencia con MS Project y lo informa en pantalla:

- Filas cuya tarea ya no existe (se omiten) y tareas con costo que no tienen fila en el libro.
- Diferencia entre el AC total cargado y el costo real (CRTR) del proyecto en MS Project.
- Celdas de ingreso con números guardados como texto o con fórmulas: se reemplazan por su valor
  numérico, porque Excel no suma texto y la biblioteca no puede recalcular fórmulas.

Al terminar, muestra en pantalla los indicadores a la fecha de corte, con los mismos cálculos que la
hoja `Resumen`.

---

## 5. Archivo de configuración

Se debe copiar `config.ejemplo.toml` como `config.toml` y ajustarlo. Todas las claves son opcionales.

| Sección / clave | Por defecto | Uso |
|---|---|---|
| `[campos] edt` | `"WBS"` | Campo con el código EDT: `WBS` (campo EDT de MS Project), `OutlineNumber` (número de esquema), un campo de texto (`Text1`...) o su alias. Si un campo personalizado tiene el alias `EDT`, `edt = "EDT"` usa ese campo. El archivo de ejemplo indica el campo personalizado `"Item EDT"`. |
| `[campos] metodo` | vacío | Campo personalizado con el método de medición, p. ej. `"Text10"` o su alias `"Método EVM"`. |
| `[campos] peso_hito` | vacío | Campo numérico con el peso de cada hito, p. ej. `"Number1"` o `"Peso hito"`. |
| `[horizonte] mes_inicio` | vacío | Primer mes (`"AAAA-MM"`). Vacío = inicio de la línea base. El archivo de ejemplo indica `"2026-01"`. |
| `[horizonte] meses` | `120` | Cantidad de meses. |
| `[exportacion] planificacion` | sin filtro | Código EDT de alto nivel de la planificación a exportar, por ejemplo `"INTERNO"` o `"CLIENTE"`; el archivo de ejemplo indica `"INTERNO"`. Si se omite, se exportan todas las tareas. |
| `[exportacion] nivel_maximo` | `3` | Nivel máximo del código EDT exportado; los niveles inferiores se suman en su elemento de este nivel. |
| `[exportacion] linea_base` | la del proyecto | `0` = Línea base, `1`..`10` = Línea base N. |
| `[exportacion] rellenar_edt` | `true` | `1` → `1.0.0`, `1.2` → `1.2.0`. |
| `[exportacion] incluir_inactivas` | `false` | Incluir tareas inactivas. |
| `[exportacion] distribucion` | `"habiles"` | Para tareas sin distribución mensual en MS Project (con `.mpp` y con `.xml`): reparto lineal por tiempo laborable (`habiles`) o por días corridos (`calendario`). |
| `[exportacion] feriados` | `[]` | Feriados adicionales. Con XML se agregan al calendario del proyecto y afectan todos los repartos por tiempo laborable; con `.mpp` solo el reparto lineal (MS Project usa sus propios calendarios). |
| `[exportacion] metodo_por_defecto` | `"Por definir"` | Método para las filas sin método. |
| `[exportacion] moneda` | `"EUR"` | Moneda que se muestra en el libro. Con `""` se usa el código de moneda del proyecto en MS Project. |
| `[exportacion] decimales` | `0` | Decimales de los montos en el libro. |
| `[metodos] lista` | ver ejemplo | Opciones de la lista desplegable de la columna C. |
| `[semaforo] rojo` / `amarillo` | `0.90` / `0.95` | Umbrales de SPI y CPI: rojo si es menor que `rojo`, amarillo si es menor que `amarillo`, verde en otro caso. |
| `[multiempresa] activar` | `false` | Activa el modo multi-empresa (sección 9). |
| `[multiempresa] campo_empresa` | `"Empresa"` | Campo de texto de MS Project con la empresa de cada tarea (por ejemplo `"Text3"`, o su alias). |
| `[multiempresa] empresa_a` / `empresa_b` | `"EmpresaA"` / `"EmpresaB"` | Valores de la columna de empresa y sufijo de las hojas de cada empresa (`PV-EmpresaA`...). Máximo 15 caracteres, sin `[ ] : * ? / \ '`. |
| `[multiempresa] ambas` | `"Ambas"` | Valor que indica una tarea de las dos empresas. |
| `[multiempresa] campo_reparto` | `"% Reparto"` | Campo numérico con el % de la empresa A en las tareas `Ambas`. |
| `[multiempresa] reparto_por_defecto` | `50` | % de la empresa A cuando el campo de reparto está vacío, es 0 o está fuera de rango, y en las tareas sin empresa. |
| `[multiempresa] ac_desde_project` | `false` | `false`: el AC de cada empresa se carga a mano. `true`: `actualizar` también lo carga desde MS Project. |

Los nombres de campo aceptan el nombre interno en inglés (`Text10`, `Number1`), el nombre en
español (`Texto10`, `Número1`) o el alias definido en MS Project.

---

## 6. El libro Excel

| Hoja | Contenido |
|---|---|
| **Cuadro de Mando** | Tablero del mes de reporte (desplegable): valores acumulados, desviaciones, índices con semáforo, pronósticos y tres gráficos. Ver sección 6.1. |
| **Resumen** | Fecha de corte (celda amarilla `B5`, con nombre `FechaCorte`), parámetros (incluida la planificación exportada, en `B14`) e indicadores del proyecto: BAC, PV, EV, AC, SV, CV, SPI, CPI, EAC (3 fórmulas), ETC, VAC, TCPI, % planificado/completado/gastado y Cronograma Ganado (ES, SV(t), SPI(t), duración y término estimados). SPI, CPI y SPI(t) con semáforo. |
| **PV** | Formato A–F + 120 meses. Fila 3: número de mes. Fila 4: fecha de cada mes. Al final: columna **Total horizonte** (en rojo si no coincide con el BAC) y fila **TOTAL**. |
| **EV** | Misma estructura. Celdas amarillas = ingreso del **valor del mes** (no acumulado). |
| **AC** | Misma estructura. Celdas amarillas = costo real **del mes**. |
| **PV-EmpresaA**, **EV-EmpresaA**, **AC-EmpresaA** y las de EmpresaB | Solo en modo multi-empresa (sección 9): valores de cada empresa, con la misma estructura. Las hojas `PV`, `EV` y `AC` son la suma de las dos empresas. |
| **Evolución mensual** | Tabla mensual calculada con fórmulas: PV, EV y AC del mes y acumulados, % de avance, SPI y CPI acumulados (los datos de la curva S). EV y AC acumulados se muestran desde el primer mes con datos hasta la fecha de corte. |
| **Indicadores EDT** | Para cada fila de la EDT, a la fecha de corte: BAC, PV, EV, AC, SV, CV, SPI, CPI, EAC, VAC y porcentajes, con semáforo. |
| **Validación** | Advertencias de la exportación (ver sección 11). |
| **Listas** | Métodos de medición del desplegable y su regla de cálculo. Se pueden agregar métodos al final de la columna A, sin dejar filas vacías. |
| `_graficos`, `_estructura`, `_meta` | Hojas ocultas: datos de los gráficos del Cuadro de Mando y datos que usan `actualizar` y `--migrar-desde`. **No se deben modificar.** |

Reglas importantes:

- Las **filas de nivel 1 y 2** con subelementos son **subtotales** (fórmulas `SUM` de sus hijos).
  Los valores se ingresan solo en las filas de nivel 3 (o en filas sin hijos).
- Las filas se pueden **contraer por nivel** con los botones de esquema (1, 2, 3) a la izquierda.
- Los gráficos no requieren ninguna acción: al cargar valores en las hojas `EV` o `AC`, al cambiar la
  fecha de corte en `Resumen!B5` o el mes de reporte del `Cuadro de Mando`, Excel recalcula las tablas
  y los gráficos.
- No se deben insertar ni eliminar filas o columnas en `PV`, `EV`, `AC` ni en las hojas de cada
  empresa: `actualizar` ubica cada fila por su posición. Para cambiar la EDT, se modifica MS Project y
  se vuelve a generar el libro con `--migrar-desde`.
- `--migrar-desde` traslada el historial de cada fila por el UniqueID de la tarea en MS Project (o por
  código EDT si el proyecto cambió de nombre), sin duplicarlo. Si una tarea quedó dentro de un paquete
  nuevo, su historial se suma a ese paquete. Si una fila se dividió en varias o la tarea se eliminó, su
  historial no se traslada y se informa como **Error** en `Validación` para cargarlo a mano.
- Al migrar, el método de la columna C se conserva solo si fue **editado en Excel**. Si el método cambió
  en MS Project, se usa el nuevo.
- Las fórmulas del libro se calculan al abrirlo en Excel. En la vista protegida de Excel o en vistas
  previas de archivos los indicadores pueden aparecer vacíos hasta habilitar la edición.
- `actualizar` vuelve a crear los gráficos del `Cuadro de Mando` (la biblioteca que escribe el libro
  no conserva gráficos). Por eso no se deben agregar gráficos ni imágenes propias en este libro; los
  gráficos adicionales se deben hacer en otro libro que lea estos datos.
- El PV acumulado a la fecha de corte incluye **el mes completo** de esa fecha. Como fecha de corte se
  debe usar el **último día del mes**.
- Los libros generados con versiones anteriores a la 1.3.0 conservan su hoja `Curva S` y `actualizar`
  sigue funcionando con ellos. Para obtener el `Cuadro de Mando`, se debe generar un libro nuevo con
  `--migrar-desde` (se conservan los EV y AC cargados).

### 6.1 Cuadro de Mando

Es la primera hoja del libro y resume el estado del proyecto en el **mes de reporte**:

- **Mes de reporte** (celda amarilla `E5`): se elige en el desplegable, que ofrece los meses hasta la
  fecha de corte de `Resumen`. `actualizar` lo fija en el mes de corte. Todos los valores son
  acumulados hasta el final del mes elegido.
- **Método EAC** (celda amarilla `R5`): `BAC / CPI` (por defecto), `AC + (BAC - EV)` o
  `AC + (BAC - EV) / (CPI × SPI)`. Define el EAC y, con él, ETC, VAC y TCPI (EAC).
- **Valores acumulados:** BAC, PV, EV y AC, con el % del BAC.
- **Desviaciones e índices:** CV (y CV % del EV), SV (y SV % del PV), CPI y SPI con semáforo según los
  umbrales de `[semaforo]`.
- **Pronósticos:** ETC = EAC − AC, EAC, VAC = BAC − EAC, TCPI (BAC) = (BAC − EV) / (BAC − AC) y
  TCPI (EAC) = (BAC − EV) / (EAC − AC). El TCPI tiene semáforo: verde hasta 1 / `amarillo`, amarillo
  hasta 1 / `rojo` y rojo por encima (o si es negativo). Con el método `BAC / CPI`, TCPI (EAC) es igual
  al CPI.
- Mientras no hay AC cargado (por ejemplo, en multi-empresa antes de cargar el AC a mano), los
  indicadores de costo (CV, CPI, ETC, EAC, VAC y TCPI) quedan vacíos.

Gráficos (nativos de Excel, calculados con fórmulas):

| Gráfico | Contenido |
|---|---|
| Evolución EVM | PV, EV y AC acumulados y BAC. El PV abarca hasta el fin del plan (o hasta la fecha de corte, si es posterior); EV y AC, hasta el mes de reporte, marcado con una línea vertical. |
| Evolución de CPI y SPI | CPI y SPI acumulados mes a mes hasta la fecha de corte, sobre las zonas de alarma roja (< `rojo`), amarilla (`rojo` a `amarillo`) y verde (≥ `amarillo`), y la línea objetivo 1,00. |
| CPI vs SPI por elemento EDT | Un punto por cada elemento EDT de nivel 1 (rombo) y 2 (círculo) en el mes de reporte. Eje X: CPI; eje Y: SPI; el centro es (1, 1). Arriba a la derecha: adelantado y con ahorro; abajo a la izquierda: atrasado y con sobrecosto. Junto al gráfico, una tabla con el CPI, el SPI y el % completado de cada elemento. |

Los gráficos de índices usan una escala fija de 0,5 a 1,5, igual en todos los informes para poder
compararlos. Los valores fuera de escala se dibujan en el límite y su valor exacto está en la tabla.
La hoja se imprime en vertical, en una página de ancho.

---

## 7. Cómo se calcula el PV

**Origen de los montos mensuales:**

| Motor | Origen del PV mensual |
|---|---|
| `.mpp` (COM) | MS Project entrega el **costo de línea base por mes** de cada tarea (`TimeScaleData`). Es la misma distribución que muestra la vista *Uso de tareas*: respeta calendarios, perfiles de trabajo, tablas de tasas y acumulación de costos fijos. |
| `.xml` | El XML trae el costo de línea base escalonado en tramos. Los tramos que cruzan de un mes a otro se reparten según el **calendario del proyecto** incluido en el XML. |
| Tareas sin datos mensuales | Si MS Project no tiene la distribución de una tarea (p. ej. un costo de línea base escrito a mano), el BAC se reparte en forma lineal entre el inicio y el fin de la línea base, por tiempo laborable. La hoja `Validación` informa cuántas tareas se repartieron así. |

**Planificación:** si se indica `planificacion` (por ejemplo `INTERNO`), solo se exportan las tareas
que dependen de la tarea con ese código EDT. Esa tarea contenedora no tiene fila propia. Si los
códigos de sus tareas incluyen el prefijo de la planificación (por ejemplo `INTERNO.1.1`), el prefijo
se quita.

**Niveles por código EDT:** el nivel de cada tarea se toma de su **código EDT**, no de su posición en
el esquema de MS Project: `2.0.0` es nivel 1, `2.1.0` nivel 2 y `2.1.1` nivel 3 (los ceros finales no
cuentan). Tienen fila propia las tareas con código EDT hasta el nivel máximo (3 por defecto). El PV de
una fila de nivel 3 es la suma del PV de todas las tareas que cuelgan de ella. Las filas de nivel 1 y 2
suman a sus hijos en Excel.

**Tareas sin código EDT:** no tienen fila propia. Si dependen de un elemento EDT, su costo se suma en
ese elemento: dentro de un paquete de nivel 3, en el propio paquete; directamente bajo un elemento de
nivel 1 o 2, en la fila `<código>-D Otros costos de «tarea»`. Si no dependen de ningún elemento EDT
(por ejemplo, un hito suelto bajo la planificación), se omiten y se informa en `Validación` con su
costo.

**Costos asignados a tareas resumen:** si una tarea resumen tiene costo propio (por ejemplo, un
costo fijo o un recurso asignado directamente a la tarea resumen), ese costo también se registra en la
fila `<código>-D Otros costos de «tarea»`, para no perder ese presupuesto, y se informa en `Validación`.

**Horizonte:** el mes 1 es, por defecto, el mes de inicio de la línea base. Si se elige un mes de
inicio posterior, el PV de los meses anteriores se **acumula en el mes 1** para que la curva acumulada
sea correcta. El PV posterior al mes 120 no aparece en las columnas y se informa en `Validación`; en
ese caso se debe aumentar `meses`.

**Sin línea base:** si el proyecto no tiene guardada la línea base elegida, se usa el programa actual
(costo y fechas actuales) y se advierte en `Validación` y en el título de las hojas.

---

## 8. Cómo se calculan EV y AC

`actualizar` calcula el **EV acumulado** de cada fila con su método (columna C de la hoja `PV`,
que se puede modificar) y escribe la diferencia con los meses anteriores en el mes de corte.

| Método (columna C) | EV acumulado a la fecha de corte |
|---|---|
| `0/100` | 0 hasta terminar; BAC al terminar (% completado = 100). |
| `50/50`, `25/75`, `20/80` (cualquier `X/Y`) | X % del BAC al comenzar (fecha de comienzo real o avance > 0); 100 % al terminar. |
| `Hitos ponderados` | BAC × Σ pesos de los hitos cumplidos / Σ pesos. Los hitos son subtareas hito de la fila, con su peso en el campo `peso_hito`. Sin pesos, todos pesan igual. Sin hitos, usa % físico completado si alguna subtarea lo tiene registrado; si no, % completado. |
| `% Completado` | BAC × % completado de las subtareas, **ponderado por su costo de línea base**. |
| `% Físico completado` | Igual, con el campo *% físico completado* de MS Project. |
| `Unidades completadas` | BAC × *% físico completado* (allí se registra el cociente unidades hechas / unidades totales). |
| `Avance lineal (LOE)` | EV = PV acumulado de la fila (nivel de esfuerzo: no genera variación de plazo). |
| `CPTR de MS Project` | EV = valor acumulado (CPTR/BCWP) calculado por MS Project a la fecha de estado, limitado al BAC de la fila. Requiere que la línea base del informe sea la misma que MS Project usa para el valor acumulado (se advierte si no lo es). |
| `Por definir` u otro texto | Se usa `% Completado` y se informa. |

Las fórmulas fijas se reconocen también escritas con `%` (`50%/50%`, `0% / 100%`).

**AC**: es el **CRTR (ACWP)** de MS Project a la fecha de estado. Para filas de nivel 3 con
subtareas es la suma de ellas. El costo real cargado directamente en una tarea resumen de nivel 1 o 2,
o en tareas sin código EDT que cuelgan de ella, se registra en su fila `-D Otros costos`. Si el costo
real proviene del ERP, se usa `--solo ev` y el AC se carga a mano.

Notas:

- El *% completado* de MS Project mide **% de la duración**, no del trabajo realizado. Para medir el
  avance físico real se recomienda *% físico completado* o hitos ponderados.
- `% Completado` aplica el porcentaje sobre el BAC. MS Project, en cambio, calcula su CPTR siguiendo
  la curva de la línea base. Si la línea base no es lineal, ambos valores pueden diferir. Para obtener
  exactamente el valor de MS Project se debe usar el método `CPTR de MS Project`.
- Si en el primer `actualizar` el programa ya lleva meses de ejecución, todo el EV y AC acumulado
  queda en el mes de corte. Si se dispone del historial mensual, se debe cargar antes a mano en los
  meses anteriores; `actualizar` respeta esos valores y solo escribe la diferencia.

**Indicadores de la hoja `Resumen`** (todos acumulados a la fecha de corte):

| Indicador | Fórmula |
|---|---|
| SV, CV | EV − PV, EV − AC |
| SPI, CPI | EV / PV, EV / AC |
| EAC | BAC / CPI; AC + (BAC − EV); AC + (BAC − EV) / (CPI × SPI) |
| ETC, VAC | EAC − AC, BAC − EAC (con EAC = BAC / CPI) |
| TCPI | (BAC − EV) / (BAC − AC) |
| ES | Meses del plan en que el PV acumulado alcanza el EV actual (interpolación lineal). Si el EV coincide con meses sin PV planificado, se toma el momento de ese tramo más cercano a AT. |
| SV(t), SPI(t) | ES − AT, ES / AT (AT = meses transcurridos hasta la fecha de corte) |

El SPI tiende a 1 al final de un proyecto atrasado; SPI(t) y SV(t) siguen mostrando el atraso real.

---

## 9. Multi-empresa (opcional)

Cuando el proyecto se ejecuta entre dos empresas, el libro puede separar el PV, el EV y el AC de cada
una. Este modo está desactivado por defecto y se activa con `activar = true` en la sección
`[multiempresa]` del archivo de configuración.

**Preparación en MS Project:**

1. Se crea un campo personalizado de tarea de tipo *Texto* (por ejemplo *Texto3*) con el nombre
   `Empresa` y, en *Atributos personalizados > Búsqueda*, los valores `EmpresaA`, `EmpresaB` y
   `Ambas` (los valores se configuran con `empresa_a`, `empresa_b` y `ambas`).
2. Se crea un campo de tipo *Número* (por ejemplo *Número2*) con el nombre `% Reparto`: porcentaje
   de la empresa A en las tareas `Ambas` (por ejemplo, 80 = 80 % para la empresa A y 20 % para la
   empresa B).
3. Se completa la columna `Empresa`. Una tarea sin empresa hereda la de su tarea superior (con su
   `% Reparto`): basta con completarla en las tareas resumen y en las subtareas que sean de otra
   empresa.

**Parte de la empresa A en cada tarea** (la empresa B recibe el resto):

| Columna `Empresa` | Parte de la empresa A |
|---|---|
| Empresa A | 100 % |
| Empresa B | 0 % |
| `Ambas` | El valor de `% Reparto`. Si está vacío o es 0, el reparto por defecto (`reparto_por_defecto`, 50 %). Si es negativo o mayor que 100, el reparto por defecto, y se informa en `Validación`. |
| Vacía | La de la tarea superior. Si ninguna tarea superior tiene empresa, se trata como `Ambas` con el reparto por defecto y se informa en `Validación`. |
| Otro valor | Se trata como `Ambas` con el reparto por defecto y se informa en `Validación`. |

`% Reparto` solo se usa en las tareas `Ambas`; en las demás se ignora. En MS Project un campo
numérico vacío vale 0, por eso 0 equivale a vacío. Para asignar el 100 % a una empresa se indica su
nombre en `Empresa`.

**Hojas del libro:** además de `PV`, `EV` y `AC`, se crean `PV-EmpresaA`, `PV-EmpresaB`,
`EV-EmpresaA`, `EV-EmpresaB`, `AC-EmpresaA` y `AC-EmpresaB` (con los nombres configurados), con el
mismo formato. En las hojas `PV`, `EV` y `AC` cada celda es la **suma** de las dos empresas (fórmula).

- **PV y BAC:** se reparten tarea por tarea según la parte de cada empresa. Una fila de nivel 3 puede
  incluir tareas de las dos empresas.
- **Cuadro de Mando, Resumen e Indicadores EDT** usan los totales.
- **EV:** `actualizar` escribe el EV del mes en `EV-EmpresaA` y `EV-EmpresaB`. El EV de cada empresa
  se calcula con su BAC en la fila y su propio avance:

| Método (columna C) | EV acumulado de cada empresa |
|---|---|
| `% Completado`, `% Físico completado`, `Unidades completadas` | BAC de la empresa × avance de sus tareas, ponderado por el costo de línea base que le corresponde en cada tarea. |
| `CPTR de MS Project` | CPTR de las tareas de la fila, repartido según la parte de cada empresa en cada tarea. |
| `Avance lineal (LOE)` | PV acumulado de la empresa. |
| `0/100`, `X/Y`, `Hitos ponderados` | BAC de la empresa × avance del paquete completo. |

- **AC:** se carga a mano en `AC-EmpresaA` y `AC-EmpresaB` (celdas amarillas). `actualizar` no lo
  modifica y no admite `--solo ac`. Con `ac_desde_project = true`, `actualizar` también carga el AC
  de cada empresa: el CRTR de cada tarea repartido según su parte.
- `actualizar` usa los nombres de las empresas y los campos con los que se generó el libro (guardados
  en la hoja `_meta`), aunque la configuración no tenga la sección `[multiempresa]`.
- `--migrar-desde` traslada el historial de cada empresa entre libros multi-empresa. No se admite
  entre un libro de una sola empresa y uno multi-empresa, porque el historial no se puede repartir
  entre las empresas.

---

## 10. Preparación del proyecto en MS Project

Los nombres de menú corresponden a MS Project en español 2016 a 2024 y a la versión de suscripción
(*Project Online Desktop Client*); pueden variar levemente según la versión.

### 10.1 Presupuesto (BAC) de cada tarea

Los costos se deben asignar **solo a las tareas de detalle**, nunca a las tareas resumen:

- **Recursos de trabajo** con tarifa: *Vista > Hoja de recursos*, columna *Tasa estándar*.
- **Recursos de costo** (montos fijos por partida): en la hoja de recursos, *Tipo = Costo*; el recurso
  se asigna a la tarea y el monto se escribe en *Información de la tarea > Recursos > Costo*.
- **Costo fijo** de la tarea: tabla *Costo* (*Vista > Tablas > Costo*), columnas *Costo fijo* y
  *Acumulación de costos fijos* (*Inicio*, *Prorrateado* o *Fin*). La acumulación define cómo se
  distribuye el PV en el tiempo.

Luego se debe **guardar la línea base**: *Proyecto > Establecer línea base > Establecer línea base*,
para *Todo el proyecto*. El *Costo previsto* (costo de línea base) de cada tarea es su **BAC**, y su
distribución en el tiempo es el **PV** que exporta este programa.

Al incorporar tareas nuevas aprobadas, se debe guardar la línea base para *Tareas seleccionadas* con
la opción *Resumir líneas base: en todas las tareas de resumen*. Se recomienda guardar una copia de
cada línea base aprobada en *Línea base 1..10* para conservar el historial de cambios.

### 10.2 Método de medición

1. Clic derecho en el encabezado de una columna > **Campos personalizados**.
2. Se elige *Tarea*, tipo *Texto*, se selecciona *Texto10* y con **Cambiar nombre** se le asigna el
   nombre `Método EVM`.
3. En *Atributos personalizados* se elige **Búsqueda** y se escriben los valores de la lista: `0/100`,
   `50/50`, `25/75`, `20/80`, `Hitos ponderados`, `% Completado`, `% Físico completado`,
   `Unidades completadas`, `Avance lineal (LOE)`, `CPTR de MS Project`.
4. Se inserta la columna `Método EVM` y se completa en las tareas de **nivel 3** (paquetes de trabajo).
5. En `config.toml` se indica `metodo = "Método EVM"`.

Guía práctica para elegir el método según la duración del paquete, medida en meses:

| Duración | Método sugerido |
|---|---|
| Empieza y termina en el mismo mes | `0/100` |
| Hasta 2 meses | `50/50` (o `25/75`, `20/80` si el costo se concentra al final) |
| 3 meses o más, con entregables intermedios | `Hitos ponderados` (al menos un hito por mes) |
| Trabajo medible en cantidades (m³, planos, equipos) | `% Físico completado` o `Unidades completadas` |
| Gestión, soporte, supervisión | `Avance lineal (LOE)`, limitado a lo indispensable |

### 10.3 Hitos ponderados

1. Dentro de cada paquete de nivel 3 se crean **subtareas hito** (duración 0) para cada entregable.
2. Se crea un campo personalizado *Número1* con el nombre `Peso hito` y se escribe el peso de cada
   hito (por ejemplo 30, 30 y 40). En `config.toml` se indica `peso_hito = "Peso hito"`.
3. Cada hito se marca al 100 % cuando se cumple. El EV del paquete es BAC × pesos cumplidos / total.

### 10.4 Método del valor acumulado propio de MS Project (opcional)

MS Project calcula su propio valor ganado (CPTR) con el campo **Método del valor acumulado** de cada
tarea (*Información de la tarea > Opciones avanzadas*): *% completado* o *% físico completado*.
El valor por defecto del proyecto se define en *Archivo > Opciones > Avanzadas > Opciones de valor
acumulado para este proyecto*, donde también se elige la línea base usada para el valor acumulado.
Si se usa *% físico completado* y se carga según la regla de cada método (0 / 50 / 100, pesos de
hitos, unidades), se puede elegir el método `CPTR de MS Project` en la columna C. La tabla *Valor
acumulado* (*Vista > Tablas > Más tablas*) sirve para comparar.

### 10.5 Fecha de estado y costo real

- Antes de cada `actualizar` se debe definir **Proyecto > Fecha de estado** = último día del mes
  informado. Sin fecha de estado, MS Project usa la fecha de hoy.
- MS Project calcula el costo real con el trabajo real y las tarifas. Para escribir costos reales del
  ERP en MS Project se debe desmarcar *Archivo > Opciones > Programación > Los costos reales siempre
  se calculan por Project*. También se puede cargar el AC del ERP directamente en la hoja `AC`.

### 10.6 Exportación a XML (si no se usa COM)

*Archivo > Guardar como > Tipo: Formato XML (\*.xml)*. El XML se debe guardar **después** de fijar
la fecha de estado: el CPTR y el CRTR del archivo se calculan a esa fecha.

---

## 11. Validaciones

La hoja **Validación** (y la consola) informa, entre otros:

- Planificación exportada y cantidad de tareas que dependen de ella.
- Tareas sin código EDT omitidas (con su costo) y tareas sin código EDT o de nivel mayor al máximo que
  se sumaron en su elemento EDT superior.
- Proyecto sin línea base: se usó el programa actual.
- Tareas con costo pero sin línea base (agregadas después de guardarla): no aportan PV.
- Tareas resumen con costo propio (se agrega una fila de otros costos).
- Línea base de una tarea resumen distinta de la suma de sus subtareas (línea base no acumulada).
- Distribución mensual de MS Project que no suma el costo de línea base (costo de línea base editado
  a mano): se ajusta en forma proporcional.
- PV anterior al mes 1 (acumulado en el mes 1) o posterior al mes 120 (fuera de las columnas).
- BAC total distinto del costo de línea base del proyecto o de la planificación.
- Códigos EDT duplicados (por ejemplo, si no se eligió una planificación y ambas usan los mismos
  códigos), filas sin método de medición, tareas inactivas excluidas.
- Multi-empresa: BAC de cada empresa, tareas sin empresa, valores de empresa no reconocidos y
  `% Reparto` fuera de rango.

---

## 12. Limitaciones conocidas

- El modo `.mpp` requiere MS Project de escritorio en Windows. Sin él, se debe usar el XML.
- El programa no guarda cambios en el archivo de MS Project: lo abre en modo solo lectura.
- **Proyectos maestros con subproyectos insertados:** el pool de recursos se abre en solo lectura; se
  debe verificar en `Validación` que el BAC total coincida con el del proyecto maestro.
- Motor XML: los tramos que cruzan de mes se reparten con el calendario del proyecto. No se usan los
  calendarios de recursos ni de tareas, y se ignoran las excepciones recurrentes relativas (por
  ejemplo, "primer lunes del mes" o "último viernes de marzo"). Con `.mpp` no hay esta limitación.
- Si la curva PV no alcanza el BAC dentro del horizonte, la duración planificada y el Cronograma
  Ganado quedan vacíos; en ese caso se debe aumentar `meses`.
- `% Completado` aplica el porcentaje sobre el BAC (ver sección 8).
- No hay en MS Project un método nativo 0/100, 50/50 o de hitos: el programa los calcula a partir
  del avance registrado.
- El libro está pensado para Excel. LibreOffice abre las fórmulas correctamente, pero dibuja como
  cero los meses sin datos de EV y AC en los gráficos.

---

## 13. Solución de problemas

| Mensaje o síntoma | Solución |
|---|---|
| `No se pudo iniciar MS Project por automatización COM` | Se debe verificar que MS Project de escritorio esté instalado. Alternativa: usar el XML. |
| `Para leer archivos .mpp se necesita pywin32` | Se debe instalar con `python -m pip install pywin32`. |
| Error de atributos de `win32com` o `gen_py` | Se debe borrar la carpeta `%TEMP%\gen_py` y volver a ejecutar. |
| MS Project muestra un diálogo y el programa no avanza | Se recomienda ejecutar con `--visible` para ver el diálogo (inicio de sesión, archivo protegido con contraseña, etc.). |
| Con *Project Online Desktop Client*, MS Project pide elegir una cuenta o un perfil al iniciarse | Se debe abrir MS Project manualmente, elegir la cuenta (por ejemplo *Equipo* para trabajar con archivos locales) y luego ejecutar el programa, que usa esa instancia abierta y no la cierra. |
| El proyecto está en Project Online, SharePoint u OneDrive (se abre desde una dirección web) | Se abre en MS Project y se usa `--origen activo`, o se guarda una copia local `.mpp` y se usa esa ruta. |
| El `.mpp` se descargó de Internet o del correo y MS Project no lo abre | Windows puede haberlo bloqueado: clic derecho en el archivo > *Propiedades* > *Desbloquear*. |
| `No se pudo guardar ...` | Si el libro está abierto en Excel, se debe cerrar y repetir el comando. |
| `Hay valores cargados en meses posteriores a la fecha de corte` | Se deben borrar esos meses en `EV`/`AC` o usar `--forzar`. |
| `El proyecto no tiene fecha de estado` | Se debe definir *Proyecto > Fecha de estado* en MS Project o indicar `--fecha-corte`. |
| `contiene el número ... guardado como texto` | Se debe escribir ese valor como número en Excel, o ejecutar `actualizar`, que lo convierte. |
| `La celda ... tiene una fórmula sin valor calculado` | Se debe abrir y guardar el libro en Excel, o reemplazar la fórmula de esa celda de ingreso por un número. |
| La curva no muestra EV/AC | Se debe completar la fecha de corte en `Resumen!B5`. |
| `No se encontró la planificación ...` | El mensaje lista las tareas de nivel superior con su código EDT; se debe indicar en `planificacion` uno de esos códigos. |
| El libro solo muestra los elementos de primer nivel (`1.0.0`, `2.0.0`...) | Se debe revisar que el campo indicado en `[campos] edt` contenga los códigos EDT de todas las tareas. El nivel se toma del código EDT. |
| `En modo multi-empresa el AC se carga a mano ...` | El AC se carga en las hojas de cada empresa. Para cargarlo desde MS Project se debe indicar `ac_desde_project = true` en `[multiempresa]`. |
| `--migrar-desde no se admite entre un libro de una sola empresa y uno multi-empresa` | Se debe generar el libro nuevo sin `--migrar-desde` y cargar a mano el historial que se necesite. |
| `No se encontró el campo de empresa ...` | Se debe revisar `campo_empresa` en `[multiempresa]`: nombre del campo (`Text3`) o su alias en MS Project. |
| Las tildes se ven mal en la consola | Solo afecta a la consola, no al libro. Se recomienda usar Windows Terminal o ejecutar `chcp 65001`. |

---

## 14. Pruebas y archivo de ejemplo

`ejemplos/programa_ejemplo.xml` es un programa ficticio con dos planificaciones, `INTERNO` y `CLIENTE`
(esta con costos un 15 % mayores). Cada una tiene 4 elementos de nivel EDT 1, con códigos del tipo
`1.0.0`, `1.1.0`, `1.1.1` y `1.1.1.1`, costos, línea base mensualizada, métodos de medición y pesos de
hitos. Los hitos de los paquetes y la reunión de arranque no tienen código EDT. Los códigos EDT están
también en el campo personalizado `Item EDT`, el que indica `config.ejemplo.toml`. Los campos `Empresa`
y `% Reparto` permiten probar el modo multi-empresa: gestión compartida 60/40, ingeniería de
EmpresaA (salvo los estudios de terreno, de EmpresaB), construcción de EmpresaB y puesta en marcha
compartida al 50 %. Para generar versiones con avance simulado a una fecha:

```powershell
python ejemplos\generar_ejemplo.py --fecha-estado 2027-06-30 --salida avance_2027-06.xml
```

Ciclo completo con el ejemplo (en `config.toml` se deben definir `metodo = "Método EVM"` y
`peso_hito = "Peso hito"`; para probar el modo multi-empresa, también `activar = true` en
`[multiempresa]`):

```powershell
copy config.ejemplo.toml config.toml
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --config config.toml --salida ejemplo.xlsx
python ejemplos\generar_ejemplo.py --fecha-estado 2027-06-30 --salida avance.xml
python evm_export.py actualizar --origen avance.xml --config config.toml --libro ejemplo.xlsx
```

Pruebas automáticas:

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest                          # todas las pruebas
python -m pytest -m "not libreoffice"     # sin las que recalculan con LibreOffice (más rápido)
```

| Archivo de pruebas | Contenido |
|---|---|
| `test_robustez_calculo.py` | Estructura EDT y PV, métodos de medición, multi-empresa, reparto en el tiempo, indicadores y configuración. Incluye reglas que se verifican sobre árboles EDT aleatorios. |
| `test_robustez_lectores.py` | Lectura del XML (campos, líneas base, calendarios y excepciones) y del `.mpp` con objetos que simulan MS Project. |
| `test_robustez_libro.py` | `actualizar`, `--migrar-desde`, libros modificados a mano, errores de la línea de comandos y del archivo de configuración. |
| `test_proyectos_aleatorios.py` | Proyectos aleatorios de punta a punta (`generar`, varias fechas de corte, migración y recálculo). |
| Resto | Casos del programa de ejemplo, Cuadro de Mando, multi-empresa y planificación. |

- **LibreOffice:** si está instalado, las pruebas marcadas `libreoffice` recalculan las fórmulas del
  libro y las comparan con el cálculo en Python. Se usa un perfil propio de LibreOffice, por lo que
  funcionan aunque LibreOffice esté abierto. Si no está instalado, se omiten.
- **Proyectos aleatorios:** la cantidad de proyectos (y de árboles EDT aleatorios) se aumenta con la
  variable de entorno `EVM_SEMILLAS`, por ejemplo `$env:EVM_SEMILLAS = 50` antes de `python -m pytest`.
  Cada caso indica su semilla, de modo que un fallo se puede repetir.
- **Problemas conocidos:** las pruebas marcadas `xfail` documentan problemas detectados que aún no se
  corrigieron; aparecen como `xfailed` y no hacen fallar la ejecución. Cuando se corrige el problema,
  la prueba aparece como fallida (`XPASS(strict)`) para que se quite la marca.
- **Cobertura (opcional):** `python -m pip install coverage`, luego
  `python -m coverage run --branch --source=evm -m pytest` y `python -m coverage report -m`.

El lector `.mpp` se prueba con objetos que simulan MS Project. Con MS Project disponible, se recomienda
una verificación con un proyecto real: comparar el BAC total y el PV de algunos meses con la vista
*Uso de tareas* (campo *Costo previsto* en la escala temporal mensual).

### Estructura del código

| Archivo | Función |
|---|---|
| `evm_export.py` | Punto de entrada. |
| `evm/cli.py` | Comandos y opciones. |
| `evm/config.py` | Configuración TOML. |
| `evm/model.py` | Datos comunes: tareas y proyecto leídos de MS Project. |
| `evm/reader_com.py` | Lectura de `.mpp` con MS Project (COM). |
| `evm/reader_xml.py`, `evm/workcalendar.py` | Lectura del XML de MS Project y de su calendario. |
| `evm/fields.py` | Identificación de campos personalizados (Texto1..30, Número1..20 y alias). |
| `evm/wbs.py` | Estructura EDT hasta el nivel máximo y PV por fila. |
| `evm/companies.py` | Multi-empresa: parte de cada empresa por tarea y reparto del EV y el AC. |
| `evm/timeline.py` | Horizonte mensual y repartos en el tiempo. |
| `evm/workbook.py` | Escritura del libro Excel y de sus indicadores. |
| `evm/dashboard.py` | Hoja `Cuadro de Mando` y sus tres gráficos. |
| `evm/methods.py` | Métodos de medición del EV. |
| `evm/update.py` | Comando `actualizar` y lectura del libro. |
| `evm/migrate.py` | Traslado de EV, AC y métodos al regenerar el libro (`--migrar-desde`). |
| `evm/indicators.py` | Indicadores que se muestran en pantalla (mismos cálculos que la hoja `Resumen`). |

---

## 15. Flujo de trabajo del repositorio

- La rama `main` contiene solo **versiones estables**. No se hacen commits directos en `main`.
- Todo cambio se hace en una rama propia y se incorpora mediante un **Pull Request**:

```powershell
git switch main
git pull
git switch -c mejora/descripcion-breve
# ... cambios ...
python -m pytest
git add -A
git commit -m "Descripción del cambio"
git push -u origin mejora/descripcion-breve
gh pr create --base main --fill
```

- Antes de abrir el Pull Request, todas las pruebas deben pasar (`python -m pytest`).
- No se deben subir al repositorio libros generados, configuraciones personales (`config.toml`) ni
  archivos de proyectos reales: el repositorio es público.
- Todos los textos (documentación, mensajes y notas del libro) se escriben en forma impersonal, por
  ejemplo "se debe habilitar" en lugar de "habilítelo".

---

## 16. Licencia

Este programa se distribuye bajo la **licencia MIT** (ver el archivo [LICENSE](LICENSE)): se puede
usar, copiar, modificar y distribuir libremente, también con fines comerciales, siempre que se
conserve el aviso de copyright. Se entrega "tal cual", sin garantía de ningún tipo y sin
responsabilidad de los autores.
