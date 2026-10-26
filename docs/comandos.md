# Comandos y opciones

[Inicio](../README.md) › [Documentación](README.md) › Comandos y opciones

El programa tiene cinco comandos: `validar` (revisa el archivo de MS Project sin crear ningún
archivo), `generar` (crea el libro con el % planificado de cada mes), `actualizar` (carga el % ganado
de un mes), `analizar` (actualiza la hoja `Análisis de desviaciones` sin leer MS Project) y
`generar-fichas` (crea un documento con la ficha de cada tarea para el cliente). El BAC y el AC se
ingresan a mano en el libro ([Cálculo del PV y el EV](calculo-pv-ev.md)). La ayuda se obtiene con:

```powershell
python evm_export.py --help
python evm_export.py validar --help
python evm_export.py generar --help
python evm_export.py actualizar --help
python evm_export.py analizar --help
python evm_export.py generar-fichas --help
```

Con el [instalador de Windows](../README.md#opción-1-instalador-para-windows-sin-python), el programa se
ejecuta con `evm_export` en lugar de `python evm_export.py` (por ejemplo, `evm_export generar --help`).

Para ver mensajes de diagnóstico se agrega `-v` **antes** del comando:
`python evm_export.py -v generar ...`

Todos los comandos muestran en la primera línea la versión del programa (`Exportador EVM 2.3.0`). En
una consola, mientras trabajan, muestran también una barra de progreso con la tarea en curso, el % de esa
tarea y del total, y el tiempo restante estimado:

```text
[██████░░░░░░░░░░░░░░]  31 %  Lectura de MS Project: 1.234 de 3.000 tareas (41 %)  ·  ETA 0:42
```

La barra se borra al terminar. Con la salida redirigida a un archivo (por ejemplo, en un registro
automático) no se muestra.

---

## `validar`: revisa el archivo de MS Project

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --config config.toml
```

Lee el archivo de MS Project con la configuración, igual que `generar`, y muestra en pantalla las
observaciones agrupadas por tema. **No crea ni modifica ningún archivo.** Revisa lo mismo que la hoja
`Validación` de `generar`:

- los códigos EDT: que el código de cada tarea dependa del de su tarea superior (bajo `1.2.0` van
  `1.2.1`, `1.2.2`...; `1.14.2` es un error), su formato, los niveles salteados, el orden y los
  duplicados;
- las fechas y la duración: tareas sin línea base, fechas invertidas, hitos con duración, tareas con
  duración 0 que no son hitos, duraciones estimadas y tareas programadas manualmente;
- los pesos de los hitos, el campo `Visible`, el método de medición (también frente a la duración del
  paquete) y las empresas.

Además, revisa dos aspectos del cronograma que no afectan al libro:

- **Avance a la fecha de estado:** *Comienzo real* y *Fin real* coherentes con el *% completado*,
  avance posterior a la fecha de estado, trabajo pendiente programado antes de la fecha de estado
  (cronograma sin actualizar) e índice de ejecución de la línea base (BEI).
- **Lógica del cronograma:** tareas sin predecesoras o sin sucesoras, adelantos, posiciones, vínculos
  comienzo a fin, vínculos en tareas resumen y delimitaciones que fijan fechas, con los criterios de la
  evaluación DCMA de 14 puntos.

El detalle de cada control está en [Validaciones](validaciones-y-problemas.md#1-validaciones). Se
recomienda ejecutar `validar` antes de `generar` y antes de cada `actualizar`.

| Opción | Descripción |
|---|---|
| `--origen`, `--config`, `--motor`, `--planificacion`, `--linea-base`, `--nivel-max`, `--visible`, `--silenciar` | Igual que en `generar`. Con varios archivos (`[[proyectos]]`) se revisan todos. Los errores silenciados no cuentan para el código de salida. |
| `--fecha-corte AAAA-MM-DD` | Fecha de corte (fecha de estado) con la que se revisa el avance. Por defecto, la fecha de estado de cada archivo. Sin fecha de estado, solo se revisa la coherencia de las fechas reales. |
| `--todas` | Muestra todas las observaciones. Por defecto, hasta 20 por tema. |

Si la planificación indicada no existe, o si dos archivos de `[[proyectos]]` tienen el mismo código
EDT (casos en los que `generar` se detiene), `validar` lo informa como error y revisa lo que puede. El
programa termina con código 1 si encuentra errores y con 0 si no los hay (aunque haya advertencias),
para usarlo en procesos automáticos. Si no puede leer la configuración o el archivo de MS Project,
termina con código 2 ([Códigos de salida](#códigos-de-salida)).

Ejemplo de salida con el avance simulado del proyecto de ejemplo a junio de 2027 (`config.toml` del
[tutorial](ejemplo-paso-a-paso.md#2-configuración); se omiten algunas líneas):

```text
Exportador EVM 2.3.0
Leyendo avance.xml ...
Validación: Programa de ejemplo EVM
  Planificación: INTERNO  |  Línea base  |  Fecha de estado: 30/06/2027
  Tareas revisadas: 51  |  Filas EDT: 36 (24 de ingreso)  |  Nivel máximo: 3

Lectura y configuración: sin observaciones

Estructura y códigos EDT: 3 informativas
  [i] EDT-017 INTERNO Planificación interna: Planificación 'INTERNO': se exportan las 51 tareas que dependen de 'Planificación interna'.
  ...

Fechas y duración: sin observaciones

Método de medición: 6 advertencias, 1 informativa
  [!] MET-004 1.2.1 Licitación de ingeniería: Método 0/100 en un paquete de 3,6 meses (línea base), más de 1: el EV queda en 0 hasta que termina. Se recomienda dividir el paquete o usar un método de avance (% completado, hitos ponderados). Umbral: [validacion] meses_0_100.
  ...

Avance a la fecha de estado: 1 advertencia
  [!] AVA-013 Índice de ejecución de la línea base (BEI) = 0,33: 4 tareas terminadas y 12 que debían terminar a la fecha de estado (DCMA: al menos 0,95). Sin terminar: 1.2.1, 1.2.2, 2.1.1, 2.1.2.1, 2.1.2.2, 2.1.2.3, 2.1.3, 2.1.4.

Lógica del cronograma: 1 advertencia
  [!] LOG-001 El cronograma no tiene vínculos entre tareas: las fechas no se recalculan cuando cambian las duraciones o el avance. Se recomienda vincular las tareas (cada tarea con al menos una predecesora y una sucesora).

Resultado: 8 advertencias, 4 informativas.
No se creó ni se modificó ningún archivo.
```

Cada línea indica la severidad (`[X]` error, `[!]` advertencia, `[i]` informativa), el código de la
observación (por ejemplo `MET-004`, explicado en [Códigos de errores y advertencias](codigos.md)), el
código EDT y el nombre de la tarea (las tareas sin código EDT, con su Id de MS Project) y la observación.
Las observaciones que ya se revisaron se pueden ocultar con `--silenciar` o con `[validacion] silenciar`
([Silenciar observaciones](codigos.md#2-silenciar-observaciones)).

---

## `generar`: crea el libro con el PV

```powershell
python evm_export.py generar --origen "C:\Programa\Programa.mpp" --config config.toml --salida EVM_Programa.xlsx
```

| Opción | Descripción |
|---|---|
| `--origen` | Archivo `.mpp` (requiere MS Project), `.xml` de MS Project, o `activo` para usar el proyecto abierto en MS Project. **Obligatoria**, salvo con varios archivos (`[[proyectos]]` en la configuración): en ese caso se omite, o se indica `NOMBRE=RUTA` para reemplazar la ruta de un archivo ([Varios archivos](varios-archivos.md)). `NOMBRE=RUTA` se puede repetir; una ruta relativa se toma desde la carpeta actual (no desde la del archivo de configuración). |
| `--salida` | Libro `.xlsx` a crear. Por defecto `<origen>_EVM.xlsx` (con `--origen activo`, `EVM.xlsx` en la carpeta actual); obligatoria con varios archivos. |
| `--config` | Archivo de configuración TOML ([Configuración](configuracion.md)). |
| `--mes-inicio AAAA-MM` | Primer mes del horizonte (mes 1). Por defecto, `mes_inicio` de la configuración o, si no se indica, el mes en que comienza la línea base. |
| `--meses N` | Cantidad de meses (columnas). Por defecto, `meses` de la configuración (120). |
| `--planificacion CODIGO` | Código EDT de la planificación a exportar (`INTERNO`, `CLIENTE`...). Reemplaza el valor del archivo de configuración; `""` exporta todas. Con varios archivos no se admite: la planificación se indica en cada `[[proyectos]]`. |
| `--linea-base N` | 0 = Línea base, 1..10 = Línea base N. Por defecto, la que el proyecto usa para el valor acumulado. Con varios archivos no se admite: se indica en cada `[[proyectos]]`. |
| `--nivel-max N` | Nivel máximo del código EDT. Por defecto, `nivel_maximo` de la configuración (3). Con `lista_edt`, las filas las decide la lista y el nivel solo se usa para completar los códigos con ceros ([Lista EDT](calculo-pv-ev.md#34-lista-edt-filas-definidas-en-un-archivo)). |
| `--motor auto\|com\|xml` | Forma de lectura. `auto` decide por la extensión. |
| `--migrar-desde libro.xlsx` | Copia al libro nuevo el BAC ingresado, el % ganado y el AC ya cargados, los métodos editados en la columna C, el análisis de desviaciones y la fecha de corte de un libro anterior ([El libro Excel](libro-excel.md#4-cambios-de-línea-base-o-de-edt---migrar-desde)). Con depuración de costes, también la previsión pegada en `PV-depuración` ([Depuración de costes](depuracion-costes.md#8-ciclo-de-ajuste)). |
| `--sobrescribir` | Reemplaza el libro de salida si ya existe. |
| `--visible` | Muestra la ventana de MS Project mientras se lee (solo COM). No tiene relación con el campo `Visible`. |
| `--silenciar CODIGOS` | Códigos de las observaciones que no se muestran, separados por comas (`LOG-004,FEC-007`, o `MET-004:1.2.1` para una sola fila); se puede repetir y se suma a `[validacion] silenciar` ([Silenciar observaciones](codigos.md#2-silenciar-observaciones)). |

Ejemplo de salida con el proyecto de ejemplo (se omiten algunas líneas):

```text
Exportador EVM 2.3.0
Leyendo ejemplos/programa_ejemplo.xml ...
Proyecto: Programa de ejemplo EVM
  Planificación: INTERNO
  Línea base  |  PV: % planificado de MS Project (duración de la línea base de las tareas y peso de los hitos) × BAC ingresado en el libro
  Filas EDT: 36 (24 de ingreso)  |  Nivel máximo: 3
  Horizonte: 01/2026 a 12/2035 (120 meses)
  BAC: se debe ingresar el BAC de cada fila en la columna D de PV (celdas amarillas).
  [i] LIB-001 El PV y el EV de cada fila son el BAC × el % planificado y el % ganado (hojas % PV y % EV), que se calculan con la duración de la línea base de las tareas y el peso de los hitos. Se debe ingresar el BAC de cada fila en las celdas amarillas de la columna D de la hoja PV.
  [i] EDT-017 INTERNO Planificación interna: Planificación 'INTERNO': se exportan las 51 tareas que dependen de 'Planificación interna'.
  [i] EDT-020 Se omitió 1 tarea sin código EDT que no depende de ningún elemento EDT: no forma parte de ninguna fila del libro.
  [i] EDT-022 14 tareas sin código EDT o de nivel EDT mayor a 3 se agregaron en su elemento EDT superior.
  [i] MET-005 1 fila sin método de medición definido (columna C).
  [!] MET-004 1.2.1 Licitación de ingeniería: Método 0/100 en un paquete de 3,6 meses (línea base), más de 1: el EV queda en 0 hasta que termina. Se recomienda dividir el paquete o usar un método de avance (% completado, hitos ponderados). Umbral: [validacion] meses_0_100.
  ...
Libro generado: ...\EVM_Ejemplo.xlsx
```

Después de `generar`, se ingresa el BAC de cada fila en las celdas amarillas de la columna D de la hoja
`PV` (en multi-empresa, de `PV-EmpresaA` y `PV-EmpresaB`).

---

## `actualizar`: carga el % ganado del mes

```powershell
python evm_export.py actualizar --origen "C:\Programa\Programa.mpp" --config config.toml --libro EVM_Programa.xlsx --fecha-corte 2026-10-31
```

| Opción | Descripción |
|---|---|
| `--origen` | Proyecto con el avance a la fecha de corte (`.mpp`, `.xml` o `activo`). **Obligatoria**, salvo con varios archivos: se leen todos los de `[[proyectos]]`, y `NOMBRE=RUTA` reemplaza la ruta de un archivo (relativa a la carpeta actual; se puede repetir). |
| `--libro` | Libro creado con `generar`. **Obligatoria.** |
| `--fecha-corte AAAA-MM-DD` | Fecha de corte. Por defecto, la fecha de estado del proyecto; si el proyecto no tiene fecha de estado, es obligatoria. Con varios archivos, sin esta opción todos deben tener la misma fecha de estado. Con `.mpp`, el programa fija esa fecha de estado en la copia abierta (no guarda el archivo). |
| `--simular` | Muestra el EV calculado para cada fila y los indicadores resultantes, sin modificar el libro. |
| `--forzar` | Permite actualizar aunque haya valores en meses posteriores. Esos meses no se recalculan: se avisa cuáles quedaron con un acumulado que ya no coincide con MS Project, para revisarlos o volver a cargarlos en orden. |
| `--sin-respaldo` | No crea la copia `<libro>.respaldo-AAAAMMDD-HHMMSS.xlsx` (si ya existe una con ese nombre, se agrega `-2`, `-3`...). |
| `--config`, `--motor`, `--visible`, `--silenciar` | Igual que en `generar`. |
| `--linea-base N` | Igual que en `generar`. Debe ser la misma línea base con la que se generó el libro: sus duraciones definen el peso de cada tarea en el % ganado. |
| `--planificacion`, `--nivel-max` | No se aplican: `actualizar` usa siempre la planificación, el nivel máximo y la lista EDT con los que se generó el libro (si la configuración indica otra planificación u otra lista, lo advierte). |

- Escribe el % ganado del mes en la hoja `% EV` (el EV es el BAC × ese %). No modifica el AC, que se
  carga a mano en la hoja `AC` (en multi-empresa, en `AC-EmpresaA` y `AC-EmpresaB`), antes o después de
  `actualizar`.
- Repetir `actualizar` para el mismo mes **reemplaza** el valor de ese mes; no lo duplica.
- `actualizar` controla la consistencia con MS Project (tareas eliminadas o inactivas, tareas con código
  EDT sin fila en el libro, filas con avance y sin BAC, celdas con texto o fórmulas, subtotales sin
  fórmula) y lo informa en pantalla ([Cálculo del PV y el EV](calculo-pv-ev.md#14-controles-y-validaciones)).
  Las tareas con código EDT agregadas en MS Project después de generar el libro no tienen fila: su avance
  no se incluye en el % ganado hasta regenerar el libro con `--migrar-desde`.
- Antes de leer MS Project verifica el libro (hojas, filas y hojas ocultas) y se detiene sin modificarlo
  si encuentra un problema; también avisa si el libro está abierto en Excel. El libro se guarda primero
  en un archivo temporal, de modo que un error al guardar no daña el libro anterior.
- Agrega a la hoja `Análisis de desviaciones` los paquetes con SPI o CPI fuera del rango verde del
  semáforo en el mes de corte, y conserva lo ya escrito ([Análisis de desviaciones](analisis-desviaciones.md)).
- Escribe el plazo según el cronograma de MS Project: columnas Q a U de `Indicadores EDT` (fin previsto,
  desvío de fin, holgura y crítica de cada elemento) y la sección de plazo del `Resumen`
  ([El libro Excel](libro-excel.md#indicadores-edt)).
- Al terminar, muestra los paquetes fuera de umbral, los indicadores a la fecha de corte (con los mismos
  cálculos que la hoja `Resumen`, incluida la tendencia) y el plazo del proyecto.

Ejemplo de salida:

```text
Exportador EVM 2.3.0
Leyendo ...\avance_2027-06.xml ...
Fecha de corte: 30/06/2027 (mes 18 del horizonte)
Observaciones por fila:
  [i] ACT-034 4.2.2 Cierre administrativo: método no definido: se usó % completado
  [i] ACT-026 El AC se carga a mano en la hoja 'AC'; los indicadores de costo usan los valores cargados.
Paquetes fuera de umbral (SPI o CPI fuera de 0,95 a 1,05): 13 (hoja 'Análisis de desviaciones')
  1.1.1      Oficina de gestión del programa (PMO)     SPI  1,00  CPI  1,11
  1.1.2      Control de proyectos y EVM                SPI  1,00  CPI  1,11
  1.1.3      Gestión de riesgos                        SPI  1,00  CPI  1,09
  1.2.1      Licitación de ingeniería                  SPI  0,00  CPI  0,00
  1.2.2      Licitación de obras                       SPI  0,50  CPI  0,54
  1.2.3      Administración de contratos               SPI  1,00  CPI  1,12
  2.1.1      Estudios de terreno                       SPI  0,50  CPI  0,54
  2.1.2      Diseño conceptual                         SPI  1,00  CPI  1,09
  2.1.3      Ingeniería básica civil                   SPI  0,85  CPI  0,93
  2.1.4      Ingeniería básica eléctrica               SPI  0,85  CPI  0,93
  2.2.1      Detalle civil                             SPI  0,84  CPI  0,93
  2.2.2      Detalle mecánico                          SPI  0,86  CPI  0,93
  2.2.3      Detalle eléctrico y control               SPI  0,84  CPI  0,93

BAC (EUR)              44.210.000
PV acumulado              5.895.918   % planificado 13,3 %
EV acumulado              4.946.380   % completado  11,2 %
AC acumulado              5.393.844   % gastado     12,2 %
SV = EV - PV               -949.538   SPI = 0,84
CV = EV - AC               -447.464   CPI = 0,92
EAC = BAC/CPI            48.209.366   VAC = -3.999.366
TCPI                           1,01
ES (meses)                    15,98   AT = 18  SV(t) = -2,02  SPI(t) = 0,89
Últimos 3 meses                CPI = 0,95  SPI = 0,88
Mes anterior                   CPI = 0,91  SPI = 0,84  EAC = 48.494.504  (variación del EAC: -285.138)

Plazo según MS Project: fin de línea base 18/10/2030  |  fin previsto 09/12/2030 (desvío +36 días laborables)
  Paquetes críticos sin terminar: 5 (5 con fin atrasado)
Libro actualizado: ...\EVM_Ejemplo.xlsx
Los indicadores y los gráficos se calculan en el libro (hojas 'Cuadro de Mando', 'Resumen' e 'Indicadores EDT').
```

---

## `analizar`: actualiza el análisis de desviaciones

```powershell
python evm_export.py analizar --libro EVM_Programa.xlsx
```

Vuelve a calcular la lista de paquetes fuera de umbral de la hoja `Análisis de desviaciones` para el mes
de la fecha de corte del libro, con los valores que tiene el libro y **sin leer MS Project**. Se usa
después de cargar a mano el AC del mes (por ejemplo, del ERP), para incluir los paquetes con CPI fuera
de umbral ([Análisis de desviaciones](analisis-desviaciones.md#3-cierre-mensual)).

| Opción | Descripción |
|---|---|
| `--libro` | Libro creado con `generar`, con al menos un `actualizar`. **Obligatoria.** |
| `--sin-respaldo` | No crea la copia de respaldo del libro. |
| `--config` | Solo se usa `[validacion] silenciar`: el libro guarda los umbrales del semáforo al generarse. |
| `--silenciar CODIGOS` | Igual que en `generar`. |

---

## `generar-fichas`: fichas de tareas para el cliente

```powershell
python evm_export.py generar-fichas --origen "C:\Programa\Programa.mpp" --config config.toml --salida Fichas_Programa.md
```

Crea un documento Markdown (`.md`) con una ficha por cada elemento EDT, para entregar al cliente. **No
crea ni modifica el libro.** Las fichas corresponden a las filas del libro (planificación,
`nivel_maximo`, campo `Visible`, tareas inactivas y varios archivos), sin las filas «-D». Cada ficha
tiene como título el código EDT y el nombre, con el nivel de su fila (`##` para `1.0.0`, `###` para
`1.1.0`...), y estos datos:

| Dato | Origen |
|---|---|
| Inicio y fin planificados | Fechas de la línea base, las mismas del % planificado (`linea_base`). Si una tarea no tiene línea base, se usan sus fechas actuales y la ficha lo indica en una observación. Si el proyecto no tiene línea base, todas las fichas usan el programa actual y la cabecera del documento lo indica. |
| Duración | Duración de la línea base en días laborables: la duración ÷ las horas por día del proyecto (*Archivo > Opciones > Programación*, 8 por defecto). Si el archivo no tiene la duración, se usa el tiempo laborable del calendario entre las dos fechas. Los hitos muestran `0 días (hito)`. |
| Predecesoras y sucesoras | `código - nombre` de los elementos con ficha vinculados, en el orden del libro. |
| Descripción (opcional) | Las notas de la tarea o un campo de texto, si se indica `[campos] descripcion` ([Configuración](configuracion.md)). Por defecto no se lee. |

**Vínculos.** En MS Project, los vínculos suelen estar entre tareas de detalle, por debajo del nivel
máximo. Cada vínculo se traslada a los elementos con ficha que contienen sus dos tareas. Por ejemplo, si
una tarea de `1.2.3` tiene como predecesora una tarea de `1.2.1`, la ficha de `1.2.3` muestra `1.2.1`
como predecesora y la de `1.2.1` muestra `1.2.3` como sucesora.

- Los vínculos entre tareas del mismo elemento no se muestran. Si hay varios vínculos entre los mismos
  dos elementos, se muestran una sola vez.
- La ficha de un elemento resumen (por ejemplo `1.2.0`, con fichas debajo) muestra solo los vínculos
  de la propia tarea resumen y de sus tareas sin ficha propia (por ejemplo, las tareas sin código EDT
  que dependen directamente de ella).
- Los vínculos con tareas sin ficha (de otra planificación, inactivas, con `Visible = No` o de otro
  proyecto) no se muestran, para que el documento no incluya tareas que no se entregan. La consola
  indica cuántos se omitieron.
- No se indica el tipo de vínculo (fin a comienzo, comienzo a comienzo...) ni la posición o el adelanto.

| Opción | Descripción |
|---|---|
| `--salida` | Archivo `.md` a crear. Por defecto `<origen>_fichas.md` (con `--origen activo`, `Fichas.md` en la carpeta actual); obligatoria con varios archivos. |
| `--sobrescribir` | Reemplaza el archivo de salida si ya existe. |
| `--origen`, `--config`, `--motor`, `--planificacion`, `--linea-base`, `--nivel-max`, `--visible`, `--silenciar` | Igual que en `generar`. Con varios archivos (`[[proyectos]]`) se crea un solo documento con las fichas de todos, en el orden del libro. |

Ejemplo de salida con el proyecto de ejemplo, que no tiene vínculos entre tareas:

```text
Exportador EVM 2.3.0
Leyendo ejemplos/programa_ejemplo.xml ...
Proyecto: Programa de ejemplo EVM
  Planificación: INTERNO
  Línea base
  Fichas: 36  |  Nivel máximo: 3  |  Vínculos entre fichas: 0
Fichas generadas: ...\ejemplos\programa_ejemplo_fichas.md
```

Fragmento de un documento con tareas ilustrativas. Los estudios de terreno tienen dos tareas de detalle,
*Topografía* y *Geotecnia*, que en MS Project son las predecesoras de *Diseño conceptual*:

```markdown
# Fichas de tareas: Programa de ejemplo EVM

- Planificación: INTERNO
- Fechas planificadas: Línea base
- Duración: en días laborables
- Nivel de desglose EDT: 3
- Fichas: 5
- Fecha de emisión: 24/10/2026

...

#### 2.1.1 Estudios de terreno

- Inicio planificado: 05/01/2026
- Fin planificado: 13/02/2026
- Duración: 30 días
- Predecesoras: ninguna
- Sucesoras:
    - 2.1.2 - Diseño conceptual

#### 2.1.2 Diseño conceptual

- Inicio planificado: 16/02/2026
- Fin planificado: 27/03/2026
- Duración: 30 días
- Predecesoras:
    - 2.1.1 - Estudios de terreno
- Sucesoras:
    - 2.1.3 - Ingeniería básica civil
```

El documento se abre con cualquier editor o visor de Markdown (por ejemplo, Visual Studio Code). Para
enviarlo en Word o PDF se puede convertir, por ejemplo, con [Pandoc](https://pandoc.org/):
`pandoc Fichas_Programa.md -o Fichas_Programa.docx`. Los títulos de las fichas pasan a ser títulos de
Word, con los que se puede insertar un índice.

---

## Ejemplos frecuentes

```powershell
# Proyecto abierto en MS Project (por ejemplo, desde Project Online o SharePoint)
python evm_export.py generar --origen activo --config config.toml --salida EVM_Programa.xlsx

# Exportar la planificación del cliente en lugar de la indicada en la configuración
python evm_export.py generar --origen Programa.mpp --config config.toml --planificacion CLIENTE --salida EVM_Cliente.xlsx

# Revisar el archivo de MS Project antes de generar o actualizar el libro
python evm_export.py validar --origen Programa.mpp --config config.toml

# Revisar el cálculo antes de modificar el libro
python evm_export.py actualizar --origen Programa.mpp --config config.toml --libro EVM_Programa.xlsx --simular

# Cierre del mes: cargar el % ganado y, después de cargar a mano el AC del mes, analizar las desviaciones
python evm_export.py actualizar --origen Programa.mpp --config config.toml --libro EVM_Programa.xlsx
python evm_export.py analizar --libro EVM_Programa.xlsx

# Fichas de tareas de la planificación del cliente (documento Markdown)
python evm_export.py generar-fichas --origen Programa.mpp --config config.toml --planificacion CLIENTE --salida Fichas_Cliente.md

# Nueva línea base: libro nuevo conservando el historial
python evm_export.py generar --origen Programa.mpp --config config.toml --salida EVM_Programa_LB2.xlsx --migrar-desde EVM_Programa.xlsx

# Varios archivos de MS Project ([[proyectos]] en la configuración): sin --origen
python evm_export.py generar --config config.toml --salida EVM_Programa.xlsx
python evm_export.py actualizar --config config.toml --libro EVM_Programa.xlsx
```

Si no se activó el entorno virtual, se usa `.venv\Scripts\python.exe` en lugar de `python`.

---

## Códigos de salida

Todos los comandos terminan con un código que permite usarlos en procesos automáticos:

| Código | Significado |
|---|---|
| 0 | El comando terminó correctamente. En `validar`: no hay errores, aunque puede haber advertencias. |
| 1 | `validar` encontró errores (`[X]`) en el archivo de MS Project: se deben corregir antes de generar o actualizar el libro. Ningún otro comando termina con 1. |
| 2 | El comando no se pudo ejecutar por un problema previsto, explicado en el mensaje `ERROR CFG-001: ...` (con su [código](codigos.md)): configuración no válida, archivo inexistente o sin permiso, libro abierto en Excel o con un formato que no se reconoce, MS Project que no responde, error de lectura o escritura (disco lleno, unidad de red no disponible) u opción mal escrita. |
| 3 | Error inesperado del programa (`ERROR SIS-001: error inesperado: ...`). Se debe ejecutar de nuevo con `-v` y, si se repite, informarlo en un [ticket](https://github.com/MarceloSalazar/evm-automation/issues/new/choose) con ese detalle. |
| 130 | Ejecución cancelada con Ctrl+C. |

Por ejemplo, en PowerShell, para actualizar el libro solo si `validar` no encuentra errores:

```powershell
python evm_export.py validar --origen Programa.mpp --config config.toml
if ($LASTEXITCODE -eq 0) {
    python evm_export.py actualizar --origen Programa.mpp --config config.toml --libro EVM_Programa.xlsx
}
```
