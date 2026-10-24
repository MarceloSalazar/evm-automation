# Archivo de configuración

[Inicio](../README.md) › [Documentación](README.md) › Configuración

La configuración se guarda en un archivo TOML. Se debe copiar `config.ejemplo.toml` como `config.toml`
y ajustarlo a cada proyecto; se indica con `--config config.toml`. **Todas las claves son opcionales**:
si se omite una clave, se usa el valor por defecto.

```powershell
copy config.ejemplo.toml config.toml
```

> `config.toml` no se sube al repositorio (está en `.gitignore`): puede contener nombres de campos o de
> empresas del proyecto real.

---

## Claves

| Sección / clave | Por defecto | Uso |
|---|---|---|
| `[campos] edt` | `"WBS"` | Campo con el código EDT: `WBS` (campo EDT de MS Project), `OutlineNumber` (número de esquema), un campo de texto (`Text1`...) o su alias. Si un campo personalizado tiene el alias `EDT`, `edt = "EDT"` usa ese campo. El archivo de ejemplo indica el campo personalizado `"Item EDT"`. |
| `[campos] metodo` | vacío | Campo personalizado con el método de medición, p. ej. `"Text10"` o su alias `"Método EVM"`. |
| `[campos] peso_hito` | vacío | Campo numérico con el peso de cada hito, p. ej. `"Number1"` o `"Peso hito"`: el % del BAC de la fila que se imputa en el mes del hito ([Cálculo del PV y el EV](calculo-pv-ev.md#4--planificado-peso-de-cada-tarea)); con el método `Hitos ponderados`, también el peso de cada hito en el % ganado. Vacío = no se lee: los hitos no tienen valor propio y, con `Hitos ponderados`, pesan igual. |
| `[campos] visible` | `"Visible"` | Campo de texto con la visibilidad de cada tarea: `Sí` = fila propia, también por debajo de `nivel_maximo`; `No` = la tarea y su rama quedan fuera del PV, el EV y el AC (si todas las subtareas de una tarea tienen `No`, la tarea las representa); vacío = según `nivel_maximo` ([Campo Visible](calculo-pv-ev.md#33-campo-visible-más-o-menos-detalle-en-cada-rama)). Si se omite la clave, se lee el campo con alias `Visible` cuando existe; `""` = no se lee. |
| `[campos] descripcion` | vacío | Descripción de cada ficha del comando `generar-fichas` ([Comandos](comandos.md#generar-fichas-fichas-de-tareas-para-el-cliente)): `"Notas"` = las notas de la tarea en MS Project, o un campo de texto (`"Text20"`...) o su alias. Si un campo personalizado tiene el alias `Notas`, se usa ese campo. Vacío = no se lee: las fichas no tienen descripción. |
| `[horizonte] mes_inicio` | vacío | Primer mes (`"AAAA-MM"`). Vacío = inicio de la línea base. El archivo de ejemplo indica `"2026-01"`. |
| `[horizonte] meses` | `120` | Cantidad de meses. |
| `[exportacion] planificacion` | sin filtro | Código EDT de alto nivel de la planificación a exportar, por ejemplo `"INTERNO"` o `"CLIENTE"`; el archivo de ejemplo indica `"INTERNO"`. Si se omite, se exportan todas las tareas. |
| `[exportacion] nivel_maximo` | `3` | Nivel máximo del código EDT exportado; los niveles inferiores se suman en su elemento de este nivel. |
| `[exportacion] linea_base` | la del proyecto | `0` = Línea base, `1`..`10` = Línea base N: sus fechas definen el % planificado ([Cálculo del PV y el EV](calculo-pv-ev.md#5-horizonte-de-meses-y-proyecto-sin-línea-base)). |
| `[exportacion] rellenar_edt` | `true` | `1` → `1.0.0`, `1.2` → `1.2.0`. |
| `[exportacion] incluir_inactivas` | `false` | Incluir tareas inactivas. |
| `[exportacion] distribucion` | `"habiles"` | Reparto en el tiempo de la duración de cada tarea: por tiempo laborable (`habiles`) o por días corridos (`calendario`). |
| `[exportacion] feriados` | `[]` | Feriados adicionales para repartir la duración de las tareas. Con XML se agregan al calendario del proyecto; con `.mpp` se usan con un calendario de lunes a viernes. |
| `[exportacion] metodo_por_defecto` | `"Por definir"` | Método para las filas sin método. |
| `[exportacion] moneda` | `"EUR"` | Moneda que se muestra en el libro. Con `""` se usa el código de moneda del proyecto en MS Project. |
| `[exportacion] decimales` | `0` | Decimales de los montos en el libro (0 a 4). |
| `[exportacion] titulo` | vacío | Título del informe en el `Cuadro de Mando` y en `Resumen`. Vacío = nombre del proyecto en MS Project (con varios archivos, los nombres unidos con `+`). |
| `[exportacion] depuracion_costes` | `false` | Agrega la hoja `PV-depuración`, donde se pega la previsión mensual del plan de costes de una empresa, las filas «(coste)» en su hoja PV y la hoja `Depuración costes` con el gráfico comparativo ([Depuración de costes](depuracion-costes.md)). El libro de entrega se genera sin depuración. |
| `[exportacion] depuracion_empresa` | `empresa_a` | En multi-empresa, la empresa que se depura: el valor de `empresa_a` o de `empresa_b`. En un libro de una sola empresa se depura la hoja `PV`. |
| `[exportacion] depuracion_alineamiento` / `depuracion_umbral` | `false` / `10` | Indicador de alineamiento en las filas «(coste)»: diferencia de BAC y diferencia máxima acumulada en % del BAC, y resaltado de los meses en que la diferencia acumulada supera el umbral (% del BAC de la fila, mayor que 0 y hasta 100). |
| `[metodos] lista` | la del programa | Opciones de la lista desplegable de la columna C. Sin la clave: `Por definir`, `0/100`, `50/50`, `25/75`, `20/80`, `Hitos ponderados`, `% Completado`, `% Físico completado`, `Unidades completadas` y `Avance lineal (LOE)` ([Métodos de medición](metodos-de-medicion.md)). |
| `[semaforo] rojo_inferior` / `amarillo_inferior` | `0.90` / `0.95` | Umbrales inferiores, comunes para el SPI y el CPI: rojo si el índice es menor que `rojo_inferior`; amarillo si es menor que `amarillo_inferior`. Los nombres anteriores, `rojo` y `amarillo`, siguen siendo válidos. |
| `[semaforo] amarillo_superior` / `rojo_superior` | `1.05` / `1.10` | Umbrales superiores: amarillo si el índice es mayor que `amarillo_superior`; rojo si es mayor que `rojo_superior`. Verde entre `amarillo_inferior` y `amarillo_superior`, inclusive. Se debe cumplir 0 < `rojo_inferior` ≤ `amarillo_inferior` < `amarillo_superior` ≤ `rojo_superior`. Los umbrales se guardan en el libro al generarlo y se muestran en el `Cuadro de Mando`; los paquetes fuera del rango verde se listan en [Análisis de desviaciones](analisis-desviaciones.md). |
| `[validacion] meses_0_100` / `meses_formula_fija` | `1` / `2` | Meses de duración de la línea base a partir de los cuales se advierte que un paquete con `0/100` (o con `50/50`, `25/75`, `20/80`...) es largo para ese método: el EV queda en 0 (o en el % inicial) hasta que el paquete termina. Se informa en la hoja `Validación` y en `validar` ([Validaciones](validaciones-y-problemas.md#11-datos-del-libro-generar-y-validar)). `0` = no se controla. |
| `[multiempresa] activar` | `false` | Activa multi-empresa: el BAC, el PV, el EV y el AC de dos empresas por separado ([Multi-empresa](multiempresa.md)). |
| `[multiempresa] campo_empresa` | `"Empresa"` | Campo de texto de MS Project con la empresa de cada tarea (por ejemplo `"Text3"`, o su alias). |
| `[multiempresa] empresa_a` / `empresa_b` | `"EmpresaA"` / `"EmpresaB"` | Valores de la columna de empresa y sufijo de las hojas de cada empresa (`PV-EmpresaA`...). Máximo 15 caracteres, sin `[ ] : * ? / \ '`. |
| `[multiempresa] ambas` | `"Ambas"` | Valor que indica una tarea de las dos empresas. |
| `[multiempresa] campo_reparto` | `"% Reparto"` | Campo numérico con el % de la empresa A en las tareas `Ambas`. |
| `[multiempresa] reparto_por_defecto` | `50` | % de la empresa A cuando el campo de reparto está vacío, es 0 o está fuera de rango, y en las tareas sin empresa. |
| `[[proyectos]] nombre` | — | Varios archivos de MS Project en un mismo libro: una sección `[[proyectos]]` por archivo ([Varios archivos](varios-archivos.md)). Nombre corto y fijo del archivo en el libro (hasta 30 caracteres, sin `:` ni `;`). |
| `[[proyectos]] origen` | — | Archivo `.mpp` o `.xml`. Una ruta relativa se toma desde la carpeta del archivo de configuración. |
| `[[proyectos]] planificacion` | la de `[exportacion]` | Planificación de ese archivo; `""` exporta todas sus tareas. |
| `[[proyectos]] linea_base` | la de `[exportacion]` | Línea base de ese archivo. |

Los nombres de campo aceptan el nombre interno en inglés (`Text10`, `Number1`), el nombre en español
(`Texto10`, `Número1`) o el alias definido en MS Project. Solo se buscan campos de tarea: un campo de
recurso con el mismo alias no se toma. Cómo crear esos campos se explica en
[Preparación del proyecto en MS Project](preparacion-ms-project.md#3-campos-personalizados).

Las opciones de la línea de comandos (`--planificacion`, `--linea-base`, `--nivel-max`, `--mes-inicio`,
`--meses`) reemplazan el valor del archivo de configuración. Con varios archivos (`[[proyectos]]`),
`--planificacion` y `--linea-base` no se admiten: se indican en la sección de cada archivo.

El archivo se debe guardar con codificación **UTF-8**, con o sin BOM (en el Bloc de notas: *Guardar
como*, *Codificación: UTF-8*). Los valores enteros (`meses`, `nivel_maximo`, `decimales`, `linea_base`)
no admiten decimales.

---

## Ejemplos

**Configuración mínima** para un proyecto con el código EDT en el campo *EDT* de MS Project y los
campos de método y peso de hito con alias:

```toml
[campos]
edt = "WBS"
metodo = "Método EVM"
peso_hito = "Peso hito"
```

**Programa con dos planificaciones** y el código EDT en un campo personalizado:

```toml
[campos]
edt = "Item EDT"
metodo = "Método EVM"
peso_hito = "Peso hito"

[horizonte]
mes_inicio = "2026-01"
meses = 120

[exportacion]
planificacion = "INTERNO"
moneda = "EUR"
```

**Multi-empresa**, con los nombres de las dos empresas (el BAC y el AC de cada una se ingresan en el
libro):

```toml
[multiempresa]
activar = true
campo_empresa = "Empresa"
empresa_a = "Constructora"
empresa_b = "Ingenieria"
ambas = "Ambas"
campo_reparto = "% Reparto"
reparto_por_defecto = 50
```

**Depuración de costes** de la empresa B, con el indicador de alineamiento, en un archivo aparte
(`config_depuracion.toml`) para no generar por error un libro de entrega con depuración
([Depuración de costes](depuracion-costes.md)):

```toml
[exportacion]
depuracion_costes = true
depuracion_empresa = "EmpresaB"
depuracion_alineamiento = true
depuracion_umbral = 10

[multiempresa]
activar = true
```

**Programa con dos archivos de MS Project**, el primero con planificaciones y el segundo sin ellas
([Varios archivos](varios-archivos.md)):

```toml
[campos]
edt = "Item EDT"
metodo = "Método EVM"

[exportacion]
planificacion = "INTERNO"
titulo = "Programa Norte"

[[proyectos]]
nombre = "P1"
origen = 'C:\Programa\Proyecto1.mpp'

[[proyectos]]
nombre = "P2"
origen = 'C:\Programa\Proyecto2.mpp'
planificacion = ""
```

**Semáforo más exigente** (verde de 0,98 a 1,02; rojo por debajo de 0,95 o por encima de 1,05):

```toml
[semaforo]
rojo_inferior = 0.95
amarillo_inferior = 0.98
amarillo_superior = 1.02
rojo_superior = 1.05
```

**Sin alerta por índices altos** (solo los umbrales inferiores): se indican umbrales superiores que no
se alcanzan en la práctica.

```toml
[semaforo]
amarillo_superior = 99
rojo_superior = 99
```
