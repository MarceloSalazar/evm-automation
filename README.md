# Exportador EVM desde MS Project

Programa en Python que lee un cronograma de MS Project y genera un libro Excel para
gestionar el **Valor Ganado (EVM)**:

- **PV mensual** de cada elemento de la EDT (hasta nivel 3) durante 120 meses (10 años).
- Hojas para cargar **EV** (valor ganado) y **AC** (costo real) mes a mes.
- **Curva S** e **indicadores** (SV, CV, SPI, CPI, EAC, VAC, TCPI y Cronograma Ganado) a la fecha de
  corte, calculados con fórmulas de Excel, para el proyecto y para cada elemento de la EDT. La curva S
  es un gráfico nativo de Excel que se actualiza solo al cargar datos o cambiar la fecha de corte.
- Un comando que **calcula EV y AC desde MS Project** a la fecha de corte, según el método de
  medición de cada fila, y los carga en el mes correspondiente.

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
9. [Preparación del proyecto en MS Project](#9-preparación-del-proyecto-en-ms-project)
10. [Validaciones](#10-validaciones)
11. [Limitaciones conocidas](#11-limitaciones-conocidas)
12. [Solución de problemas](#12-solución-de-problemas)
13. [Pruebas y archivo de ejemplo](#13-pruebas-y-archivo-de-ejemplo)
14. [Flujo de trabajo del repositorio](#14-flujo-de-trabajo-del-repositorio)
15. [Licencia](#15-licencia)

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

La curva S y los indicadores se consultan en las hojas `Curva S`, `Resumen` e `Indicadores EDT` del
libro.

Para probar el programa sin MS Project se incluye un proyecto de ejemplo en XML (ver sección 13):

```powershell
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --salida ejemplo.xlsx
```

---

## 3. Ciclo mensual recomendado

**Una vez, al inicio (o al aprobar una nueva línea base):**

1. Se deben cargar los costos y los métodos de medición en MS Project y guardar la línea base
   (sección 9).
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
5. Al abrir el libro, `Resumen`, `Curva S` e `Indicadores EDT` se recalculan solos.

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
| `--linea-base N` | 0 = Línea base, 1..10 = Línea base N. Por defecto, la que el proyecto usa para el valor acumulado. |
| `--nivel-max N` | Nivel máximo de la EDT. Por defecto 3. |
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
| `--solo ev\|ac` | Actualiza solo EV o solo AC. |
| `--simular` | Muestra el EV y AC calculados para cada fila y los indicadores resultantes, sin modificar el libro. |
| `--forzar` | Permite actualizar aunque haya valores en meses posteriores. |
| `--sin-respaldo` | No crea la copia `<libro>.respaldo-AAAAMMDD-HHMMSS.xlsx` (si ya existe una con ese nombre, se agrega `-2`, `-3`...). |
| `--config`, `--motor`, `--linea-base`, `--nivel-max`, `--visible` | Igual que en `generar`. |

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
| `[campos] edt` | `"WBS"` | Campo con el código EDT: `WBS` (EDT), `OutlineNumber` (número de esquema), un campo de texto (`Text1`...) o su alias. |
| `[campos] metodo` | vacío | Campo personalizado con el método de medición, p. ej. `"Text10"` o su alias `"Método EVM"`. |
| `[campos] peso_hito` | vacío | Campo numérico con el peso de cada hito, p. ej. `"Number1"` o `"Peso hito"`. |
| `[horizonte] mes_inicio` | vacío | Primer mes (`"AAAA-MM"`). Vacío = inicio de la línea base. |
| `[horizonte] meses` | `120` | Cantidad de meses. |
| `[exportacion] nivel_maximo` | `3` | Nivel de la EDT exportado; los niveles inferiores se suman en su tarea de este nivel. |
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

Los nombres de campo aceptan el nombre interno en inglés (`Text10`, `Number1`), el nombre en
español (`Texto10`, `Número1`) o el alias definido en MS Project.

---

## 6. El libro Excel

| Hoja | Contenido |
|---|---|
| **Resumen** | Fecha de corte (celda amarilla `B5`, con nombre `FechaCorte`), parámetros e indicadores del proyecto: BAC, PV, EV, AC, SV, CV, SPI, CPI, EAC (3 fórmulas), ETC, VAC, TCPI, % planificado/completado/gastado y Cronograma Ganado (ES, SV(t), SPI(t), duración y término estimados). SPI, CPI y SPI(t) con semáforo. |
| **PV** | Formato A–F + 120 meses. Fila 3: número de mes. Fila 4: fecha de cada mes. Al final: columna **Total horizonte** (en rojo si no coincide con el BAC) y fila **TOTAL**. |
| **EV** | Misma estructura. Celdas amarillas = ingreso del **valor del mes** (no acumulado). |
| **AC** | Misma estructura. Celdas amarillas = costo real **del mes**. |
| **Curva S** | Tabla mensual calculada con fórmulas: PV, EV y AC del mes y acumulados, % de avance, SPI y CPI acumulados. Gráfico de líneas nativo de Excel con PV, EV, AC acumulados y BAC sobre todo el horizonte. EV y AC se dibujan desde el primer mes con datos hasta la fecha de corte. |
| **Indicadores EDT** | Para cada fila de la EDT, a la fecha de corte: BAC, PV, EV, AC, SV, CV, SPI, CPI, EAC, VAC y porcentajes, con semáforo. |
| **Validación** | Advertencias de la exportación (ver sección 10). |
| **Listas** | Métodos de medición del desplegable y su regla de cálculo. Se pueden agregar métodos al final de la columna A, sin dejar filas vacías. |
| `_estructura`, `_meta` | Hojas ocultas que usan `actualizar` y `--migrar-desde`. **No se deben modificar.** |

Reglas importantes:

- Las **filas de nivel 1 y 2** con subelementos son **subtotales** (fórmulas `SUM` de sus hijos).
  Los valores se ingresan solo en las filas de nivel 3 (o en filas sin hijos).
- Las filas se pueden **contraer por nivel** con los botones de esquema (1, 2, 3) a la izquierda.
- La curva S no requiere ninguna acción: al cargar valores en las hojas `EV` o `AC`, o al cambiar la
  fecha de corte en `Resumen!B5`, Excel recalcula la tabla y el gráfico.
- No se deben insertar ni eliminar filas o columnas en `PV`, `EV` y `AC`: `actualizar` ubica cada fila
  por su posición. Para cambiar la EDT, se modifica MS Project y se vuelve a generar el libro con
  `--migrar-desde`.
- `--migrar-desde` traslada el historial de cada fila por el UniqueID de la tarea en MS Project (o por
  código EDT si el proyecto cambió de nombre), sin duplicarlo. Si una tarea quedó dentro de un paquete
  nuevo, su historial se suma a ese paquete. Si una fila se dividió en varias o la tarea se eliminó, su
  historial no se traslada y se informa como **Error** en `Validación` para cargarlo a mano.
- Al migrar, el método de la columna C se conserva solo si fue **editado en Excel**. Si el método cambió
  en MS Project, se usa el nuevo.
- Las fórmulas del libro se calculan al abrirlo en Excel. En la vista protegida de Excel o en vistas
  previas de archivos los indicadores pueden aparecer vacíos hasta habilitar la edición.
- `actualizar` vuelve a crear el gráfico de la hoja `Curva S` (la biblioteca que escribe el libro no
  conserva gráficos). Por eso no se deben agregar gráficos ni imágenes propias en este libro; los
  gráficos adicionales se deben hacer en otro libro que lea estos datos.
- El PV acumulado a la fecha de corte incluye **el mes completo** de esa fecha. Como fecha de corte se
  debe usar el **último día del mes**.

---

## 7. Cómo se calcula el PV

**Origen de los montos mensuales:**

| Motor | Origen del PV mensual |
|---|---|
| `.mpp` (COM) | MS Project entrega el **costo de línea base por mes** de cada tarea (`TimeScaleData`). Es la misma distribución que muestra la vista *Uso de tareas*: respeta calendarios, perfiles de trabajo, tablas de tasas y acumulación de costos fijos. |
| `.xml` | El XML trae el costo de línea base escalonado en tramos. Los tramos que cruzan de un mes a otro se reparten según el **calendario del proyecto** incluido en el XML. |
| Tareas sin datos mensuales | Si MS Project no tiene la distribución de una tarea (p. ej. un costo de línea base escrito a mano), el BAC se reparte en forma lineal entre el inicio y el fin de la línea base, por tiempo laborable. La hoja `Validación` informa cuántas tareas se repartieron así. |

**Agregación hasta el nivel 3:** el PV de una fila de nivel 3 es la suma del PV de todas sus
subtareas (niveles 4, 5, ...). Las filas de nivel 1 y 2 suman a sus hijos en Excel.

**Costos asignados a tareas resumen:** si una tarea resumen tiene costo propio (por ejemplo, un
costo fijo o un recurso asignado directamente a la tarea resumen), se agrega una fila
`<código>-D Costos directos de «tarea»` para no perder ese presupuesto, y se informa en `Validación`.

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
subtareas es la suma de ellas. El costo real cargado directamente en una tarea resumen de nivel 1 o 2
se registra en su fila `-D Costos directos`. Si el costo real proviene del ERP, se usa `--solo ev` y
el AC se carga a mano.

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

## 9. Preparación del proyecto en MS Project

Los nombres de menú corresponden a MS Project en español 2016 a 2024 y a la versión de suscripción
(*Project Online Desktop Client*); pueden variar levemente según la versión.

### 9.1 Presupuesto (BAC) de cada tarea

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

### 9.2 Método de medición

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

### 9.3 Hitos ponderados

1. Dentro de cada paquete de nivel 3 se crean **subtareas hito** (duración 0) para cada entregable.
2. Se crea un campo personalizado *Número1* con el nombre `Peso hito` y se escribe el peso de cada
   hito (por ejemplo 30, 30 y 40). En `config.toml` se indica `peso_hito = "Peso hito"`.
3. Cada hito se marca al 100 % cuando se cumple. El EV del paquete es BAC × pesos cumplidos / total.

### 9.4 Método del valor acumulado propio de MS Project (opcional)

MS Project calcula su propio valor ganado (CPTR) con el campo **Método del valor acumulado** de cada
tarea (*Información de la tarea > Opciones avanzadas*): *% completado* o *% físico completado*.
El valor por defecto del proyecto se define en *Archivo > Opciones > Avanzadas > Opciones de valor
acumulado para este proyecto*, donde también se elige la línea base usada para el valor acumulado.
Si se usa *% físico completado* y se carga según la regla de cada método (0 / 50 / 100, pesos de
hitos, unidades), se puede elegir el método `CPTR de MS Project` en la columna C. La tabla *Valor
acumulado* (*Vista > Tablas > Más tablas*) sirve para comparar.

### 9.5 Fecha de estado y costo real

- Antes de cada `actualizar` se debe definir **Proyecto > Fecha de estado** = último día del mes
  informado. Sin fecha de estado, MS Project usa la fecha de hoy.
- MS Project calcula el costo real con el trabajo real y las tarifas. Para escribir costos reales del
  ERP en MS Project se debe desmarcar *Archivo > Opciones > Programación > Los costos reales siempre
  se calculan por Project*. También se puede cargar el AC del ERP directamente en la hoja `AC`.

### 9.6 Exportación a XML (si no se usa COM)

*Archivo > Guardar como > Tipo: Formato XML (\*.xml)*. El XML se debe guardar **después** de fijar
la fecha de estado: el CPTR y el CRTR del archivo se calculan a esa fecha.

---

## 10. Validaciones

La hoja **Validación** (y la consola) informa, entre otros:

- Proyecto sin línea base: se usó el programa actual.
- Tareas con costo pero sin línea base (agregadas después de guardarla): no aportan PV.
- Tareas resumen con costo propio (se agrega una fila de costos directos).
- Línea base de una tarea resumen distinta de la suma de sus subtareas (línea base no acumulada).
- Distribución mensual de MS Project que no suma el costo de línea base (costo de línea base editado
  a mano): se ajusta en forma proporcional.
- PV anterior al mes 1 (acumulado en el mes 1) o posterior al mes 120 (fuera de las columnas).
- BAC total distinto del costo de línea base del proyecto.
- Códigos EDT duplicados o vacíos, filas sin método de medición, tareas inactivas excluidas.

---

## 11. Limitaciones conocidas

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
  cero los meses sin datos de EV y AC en el gráfico.

---

## 12. Solución de problemas

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
| Las tildes se ven mal en la consola | Solo afecta a la consola, no al libro. Se recomienda usar Windows Terminal o ejecutar `chcp 65001`. |

---

## 13. Pruebas y archivo de ejemplo

`ejemplos/programa_ejemplo.xml` es un programa ficticio de 4 elementos de nivel 1, hasta nivel 4,
con costos, línea base mensualizada, métodos de medición y pesos de hitos. Para generar versiones con
avance simulado a una fecha:

```powershell
python ejemplos\generar_ejemplo.py --fecha-estado 2027-06-30 --salida avance_2027-06.xml
```

Ciclo completo con el ejemplo (en `config.toml` se deben definir `metodo = "Método EVM"` y
`peso_hito = "Peso hito"`):

```powershell
copy config.ejemplo.toml config.toml
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --config config.toml --salida ejemplo.xlsx
python ejemplos\generar_ejemplo.py --fecha-estado 2027-06-30 --salida avance.xml
python evm_export.py actualizar --origen avance.xml --config config.toml --libro ejemplo.xlsx
```

Pruebas automáticas:

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

Si LibreOffice está instalado, una de las pruebas recalcula todas las fórmulas del libro y las compara
con el cálculo en Python. El lector `.mpp` se prueba con objetos que simulan MS Project. Con MS Project
disponible, se recomienda una primera verificación con un proyecto real: comparar el BAC total y el PV
de algunos meses con la vista *Uso de tareas* (campo *Costo previsto* en la escala temporal mensual).

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
| `evm/timeline.py` | Horizonte mensual y repartos en el tiempo. |
| `evm/workbook.py` | Escritura del libro Excel, incluidas la curva S y los indicadores. |
| `evm/methods.py` | Métodos de medición del EV. |
| `evm/update.py` | Comando `actualizar` y lectura del libro. |
| `evm/migrate.py` | Traslado de EV, AC y métodos al regenerar el libro (`--migrar-desde`). |
| `evm/indicators.py` | Indicadores que se muestran en pantalla (mismos cálculos que la hoja `Resumen`). |

---

## 14. Flujo de trabajo del repositorio

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

## 15. Licencia

Este programa se distribuye bajo la **licencia MIT** (ver el archivo [LICENSE](LICENSE)): se puede
usar, copiar, modificar y distribuir libremente, también con fines comerciales, siempre que se
conserve el aviso de copyright. Se entrega "tal cual", sin garantía de ningún tipo y sin
responsabilidad de los autores.
