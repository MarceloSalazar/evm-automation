# Comandos y opciones

[Inicio](../README.md) › [Documentación](README.md) › Comandos y opciones

El programa tiene nueve comandos: `validar` (revisa el archivo de MS Project sin crear ningún
archivo), `generar` (crea el libro con el % planificado de cada mes), `actualizar` (carga el % ganado
de un mes), `analizar` (actualiza la hoja `Análisis de desviaciones` sin leer MS Project),
`generar-fichas` (crea un documento con la ficha de cada tarea para el cliente), `procesar-costes`
(revisa el plan de costes de una empresa y escribe sus cálculos, en un libro con
[depuración de costes](depuracion-costes.md)), `generar-cliente` (genera el `.mpp` del cliente a partir
de la planificación interna; ver [Planificación para el cliente](generar-cliente.md)), `actualizar-edt`
(numera los códigos EDT de una rama según la jerarquía de MS Project) y `corregir-linea-base` (mueve,
borra o establece la línea base de unas tareas para corregir erratas; ver [Líneas base](lineas-base.md)). El BAC y el AC se ingresan a mano en el libro
([Cálculo del PV y el EV](calculo-pv-ev.md)). La ayuda se obtiene con:

```powershell
python evm_export.py --help
python evm_export.py validar --help
python evm_export.py generar --help
python evm_export.py actualizar --help
python evm_export.py analizar --help
python evm_export.py generar-fichas --help
python evm_export.py procesar-costes --help
python evm_export.py generar-cliente --help
python evm_export.py actualizar-edt --help
python evm_export.py corregir-linea-base --help
```

Con el [instalador de Windows](../README.md#opción-1-instalador-para-windows-sin-python), el programa se
ejecuta con `evm_export` en lugar de `python evm_export.py` (por ejemplo, `evm_export generar --help`).

Para ver mensajes de diagnóstico se agrega `-v` **antes** del comando:
`python evm_export.py -v generar ...`. Para guardarlos en un archivo, con la fecha y la hora y el detalle
técnico de los errores, sin cambiar lo que muestra la consola: `python evm_export.py --registro
diagnostico.log generar ...`.

Todos los comandos muestran en la primera línea la versión del programa (`Exportador EVM 2.7.2`). En
una consola, mientras trabajan, muestran también una barra de progreso con el % del total, el tiempo
transcurrido, el tiempo restante estimado (ETA), la tarea en curso y el % de esa tarea:

```text
[██████░░░░░░░░░░░░░░]  31 %  Transcurrido 0:19  ·  ETA 0:42  ·  Lectura de MS Project: 1.234 de 3.000 tareas (41 %)
```

El ETA se calcula con el ritmo de la tarea en curso y el tiempo que suelen llevar las siguientes (según
los meses del horizonte). Indica *calculando...* hasta que hay datos suficientes, por ejemplo mientras se
abre MS Project o, en `analizar`, mientras se lee el libro.

La barra se borra al terminar y el comando indica su duración en la última línea
(`Tiempo total: 1:05`). Con la salida redirigida a un archivo (por ejemplo, en un registro automático) no
se muestran la barra ni el tiempo total. Con `-v` se informa además la duración de cada etapa.

---

## `validar`: revisa el archivo de MS Project

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --config config.toml
```

Lee el archivo de MS Project con la configuración, igual que `generar`, y muestra en pantalla las
observaciones agrupadas por tema y la evaluación DCMA de 14 puntos. **No crea ni modifica ningún
archivo**, salvo el informe de mejoras si se indica `--informe`. En cada tema van primero los
errores y, en cada severidad, las observaciones generales, las de las filas EDT (las filas de las hojas
`PV`, `EV` y `AC`), las de las tareas de detalle y las de fuera del libro
([alcance](validaciones-y-problemas.md#1-validaciones)). Revisa lo mismo que la hoja `Validación` de
`generar`:

- los códigos EDT: que el código de cada tarea dependa del de su tarea superior (bajo `1.2.0` van
  `1.2.1`, `1.2.2`...; `1.14.2` es un error), su formato, los niveles salteados, el orden y los
  duplicados;
- las fechas y la duración: tareas sin línea base, fechas invertidas, hitos con duración, tareas con
  duración 0 que no son hitos, duraciones estimadas y tareas programadas manualmente;
- los pesos de las tareas, el campo `Visible`, el método de medición (también frente a la duración del
  paquete) y las empresas.

Además, revisa aspectos del cronograma que no afectan al libro:

- **Línea base:** las líneas base guardadas y la que usa el % planificado, y las tareas resumen cuya
  línea base no coincide con la de sus tareas de detalle (no se resumió al guardarla).
- **Avance a la fecha de estado:** *Comienzo real* y *Fin real* coherentes con el *% completado*,
  avance posterior a la fecha de estado, fecha de estado que no es el último día del mes (ni su último
  día laborable), trabajo
  pendiente programado antes de la fecha de estado (cronograma sin actualizar), avance fuera de
  secuencia, índice de ejecución de la línea base (BEI), tareas incumplidas respecto de la línea base e
  índice de longitud del camino crítico (CPLI).
- **Lógica del cronograma:** tareas sin predecesoras o sin sucesoras, adelantos, posiciones, vínculos
  comienzo a fin, proporción de vínculos fin a comienzo, vínculos en tareas resumen, delimitaciones que
  fijan fechas, holgura alta, holgura negativa, tareas de duración alta y la opción *Calcular varias
  rutas críticas* de MS Project. Con una planificación, solo cuentan los vínculos entre sus tareas, y se
  informan los que la unen con otras planificaciones.

Con estos controles, `validar` revisa 12 de los 14 puntos de la
[evaluación DCMA de calidad de cronogramas](evaluacion-dcma.md), y al final de cada revisión muestra la
tabla de los 14 puntos: valor, umbral y si cumple. El detalle de cada control está en
[Validaciones](validaciones-y-problemas.md#1-validaciones). Se recomienda ejecutar `validar` antes de
`generar` y antes de cada `actualizar`.

| Opción | Descripción |
|---|---|
| `--origen`, `--config`, `--motor`, `--planificacion`, `--linea-base`, `--nivel-max`, `--visible`, `--silenciar` | Igual que en `generar`. Con varios archivos (`[[proyectos]]`) se revisan todos. Los errores silenciados no cuentan para el código de salida. |
| `--fecha-corte AAAA-MM-DD` | Fecha de corte (fecha de estado) con la que se revisa el avance. Por defecto, la fecha de estado de cada archivo. Sin fecha de estado, solo se revisan la coherencia de las fechas reales y el avance fuera de secuencia. |
| `--todas` | Muestra todas las observaciones, una por tarea. Por defecto, hasta 20 líneas por tema, con las de las tareas de detalle y fuera del libro agrupadas por código y fila del libro. |
| `--planificacion INTERNO,CLIENTE` | Varias planificaciones separadas por comas (también en `planificacion` de la configuración): revisa cada una por separado y al final las compara ([abajo](#varias-planificaciones)). Solo con un archivo, sin `[[proyectos]]`. |
| `--informe [RUTA]` | Escribe un informe de mejoras en Markdown, sin nombres ([abajo](#informe-de-mejoras)). Por defecto, `<origen>_mejoras.md` junto al archivo de MS Project. |
| `--sobrescribir` | Reemplaza el informe de mejoras si ya existe. |

Si la planificación indicada no existe, o si dos archivos de `[[proyectos]]` tienen el mismo código
EDT (casos en los que `generar` se detiene), `validar` lo informa como error y revisa lo que puede. El
programa termina con código 1 si encuentra errores y con 0 si no los hay (aunque haya advertencias),
para usarlo en procesos automáticos. Si no puede leer la configuración o el archivo de MS Project,
termina con código 2 ([Códigos de salida](#códigos-de-salida)).

Ejemplo de salida con el avance simulado del proyecto de ejemplo a junio de 2027 (`config.toml` del
[tutorial](ejemplo-paso-a-paso.md#2-configuración); se omiten algunas líneas):

```text
Exportador EVM 2.7.2
Leyendo avance.xml ...
Validación: Programa de ejemplo EVM
  Planificación: INTERNO  |  Línea base  |  Fecha de estado: 30/06/2027
  Tareas revisadas: 51  |  Filas EDT: 36 (24 de ingreso)  |  Nivel máximo: 3

Lectura y configuración: 1 informativa
  [i] LEC-024 Líneas base guardadas en las tareas revisadas: Línea base (51 tareas). El % planificado usa la Línea base, la que el proyecto usa para el valor acumulado (Archivo > Opciones > Avanzadas).

Estructura y códigos EDT: 3 informativas
  [i] EDT-017 Planificación 'INTERNO' («Planificación interna»): se exportan las 51 tareas que dependen de ella.
  ...

Campo Visible: sin observaciones

Fechas y duración: sin observaciones

Hitos y pesos: sin observaciones

Método de medición: 6 advertencias, 1 informativa
  [!] MET-004 1.2.1 Licitación de ingeniería: Método 0/100 en un paquete de 3,6 meses (línea base), más de 1: el EV queda en 0 hasta que termina. Se recomienda dividir el paquete o usar un método de avance (% completado, hitos ponderados). Umbral: [validacion] meses_0_100.
  ...

Empresas: sin observaciones

Avance a la fecha de estado: 1 advertencia, 1 informativa
  [!] AVA-013 De 12 tareas que debían terminar a la fecha de estado según la línea base, 12 no terminaron a tiempo: 8 sin terminar y 4 terminadas después de su fin de línea base (Id 2, 1.2.1, 1.2.2, 2.1.1, 2.1.2.1, 2.1.2.2, 2.1.2.3, 2.1.2.4, 2.1.2.5, 2.1.2.6 y 2 más).
  [i] AVA-015 Quedan 898 días laborables hasta el fin previsto (09/12/2030) y el fin de la línea base es el 18/10/2030: 36 días laborables de atraso.

Lógica del cronograma: 1 advertencia
  [!] LOG-001 El cronograma no tiene vínculos entre tareas: las fechas no se recalculan cuando cambian las duraciones o el avance. Se recomienda vincular las tareas (cada tarea con al menos una predecesora y una sucesora).

Evaluación DCMA de 14 puntos:
   1. Lógica (sin predecesoras o sucesoras)  sin vínculos                 ≤ 5 %   no cumple
   2. Adelantos                              sin vínculos                 0
   ...
   9. Fechas no válidas                      0 tareas                     0       cumple
  10. Recursos                               no aplica (BAC en el libro)  —
  11. Tareas incumplidas                     12 de 12 (100 %)             ≤ 5 %   no cumple
  12. Prueba del camino crítico              prueba manual                —
  13. CPLI                                   0,96                         ≥ 0,95  cumple
  14. BEI                                    0,33                         ≥ 0,95  no cumple

Resultado: 8 advertencias, 6 informativas.
  Por alcance: generales: 2 advertencias, 6 informativas; en filas EDT: 6 advertencias.
No se creó ni se modificó ningún archivo.
```

Cada línea indica la severidad (`[X]` error, `[!]` advertencia, `[i]` informativa), el código de la
observación (por ejemplo `MET-004`, explicado en [Códigos de errores y advertencias](codigos.md)), el
código EDT y el nombre de la tarea (las tareas sin código EDT, con su Id de MS Project) y la observación.
Cada problema tiene un solo código: si afecta a varias tareas, una observación general dice cuántas son
y cada tarea tiene la suya, con el mismo código. Los valores de la evaluación DCMA (% e índices) y sus
umbrales están solo en su tabla.
Una tarea de detalle indica la fila del libro en que se suma, y las observaciones de las tareas de
detalle de una misma fila se muestran en una sola línea, con `--todas` para verlas una por una:

```text
  [!] FEC-006 Duración estimada: 12 tareas de detalle de la fila 2.7.1 (2.7.1.1, 2.7.1.3, ... y 2 más). Una por tarea con --todas.
  [!] FEC-006 2.7.1.3 Montaje (detalle de 2.7.1): Duración estimada («?»).      (con --todas)
```

La línea *Por alcance* del resultado indica cuántas observaciones son generales, de filas EDT, de tareas
de detalle y de fuera del libro. Las observaciones que ya se revisaron se pueden ocultar con
`--silenciar` o con `[validacion] silenciar`
([Silenciar observaciones](codigos.md#2-silenciar-observaciones)).

### Varias planificaciones

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --config config.toml --planificacion INTERNO,CLIENTE
```

Con varias planificaciones separadas por comas, `validar` lee el archivo una vez y revisa cada
planificación por separado, como si se ejecutara una vez por planificación: sus observaciones, su
evaluación DCMA y su resultado. Al final, una tabla las compara lado a lado (tareas, inactivas, con
`Peso` = 0, revisadas, filas del libro, tareas con comienzo en cada línea base, tareas resumen con la
línea base sin resumir, vínculos propios y con otras planificaciones, tareas críticas, fecha de estado,
fin de la línea base y fin previsto, y observaciones) y lista sus diferencias: las filas del libro que
solo están en una planificación y las filas comunes con otras fechas de plan (línea base) o con otro fin
previsto. Los vínculos entre planificaciones se informan con `LOG-018`. El código de salida
es 1 si alguna planificación tiene errores. Los demás comandos admiten una sola planificación
([CFG-053](codigos.md#cfg-053)).

### Informe de mejoras

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --config config.toml --informe
```

`--informe` escribe, además de la salida en pantalla, un documento Markdown para compartir la revisión
(por defecto `<origen>_mejoras.md`; con `[[proyectos]]` se indica la ruta, `--informe C:\Revisión\mejoras.md`):

Empieza con la **fecha y la hora** de la revisión y la **versión** del programa. Una versión publicada se
indica con su número (`2.7.2`); al ejecutar desde una copia del repositorio que no está en una versión
publicada, con la última versión publicada y el commit, como lo muestra GitHub: `2.5.0 + desarrollo
(b513fbb)`. La consola, `--version` y el libro indican la versión de la misma forma.

1. **Resumen:** las cifras de cada planificación o archivo (las de la comparación de arriba, más las
   filas «Otras tareas» y las tareas sumadas en la fila de su elemento superior) y las mejoras de
   prioridad alta con más tareas.
2. **Mejoras por prioridad:** cada código de observación con su severidad, las tareas en las que aparece
   (por caso, si el código tiene varios: p. ej. `LOG-003`, sin predecesoras o sin sucesoras) y, del
   [catálogo](codigos.md), su causa y qué hacer. Prioridad **alta**: errores y advertencias que cambian el
   libro (filas, % planificado, % ganado u horizonte; los que también informa `generar`); **media**:
   advertencias de la calidad del cronograma que solo revisa `validar` (línea base, avance y lógica);
   **baja**: observaciones informativas. Las informativas que solo describen la estructura (`EDT-016`,
   `EDT-017`, `EDT-018`, `EDT-022`, `VIS-012` y `LEC-024`) no se repiten: sus datos están en el resumen.
3. **Evaluación DCMA de 14 puntos**, con una columna por planificación.
4. **Líneas base:** las líneas base guardadas en las tareas de la planificación (también las inactivas
   y las de `Peso` = 0), para detectar las que se guardaron por error. Una tabla con cada línea base:
   si la usa el % planificado, la fecha en que se guardó (solo con un .mpp: el XML no la incluye), sus
   tareas (y cuántas son resumen), su comienzo más temprano y su fin más tardío, cuántas tareas tienen
   las mismas fechas que la línea base usada e **indicios** de un guardado por error: en pocas tareas
   (menos del 10 % de las que tienen la usada; p. ej. al guardar con *Tareas seleccionadas* en otra
   línea base), solo en tareas resumen, o copia de la usada. Después, una sección por línea base: de
   la que usa el % planificado, cuántas tareas la tienen; de cada una de las demás, sus tareas. Al
   final, las **tareas sin la línea base que usa el % planificado** (no forman parte del PV), con las
   otras líneas base que tienen (sin ninguna, o solo con otras). Si su trabajo está aprobado y falta
   guardarlo, se corrige con `corregir-linea-base --establecer N`
   ([Líneas base](lineas-base.md#32-las-tres-acciones)).

   Las tareas de cada lista van **por rama, una por línea**, por código EDT (sin código, por su
   UniqueID). Sin sangría, la primera tarea de cada rama (su tarea superior no está en la lista), que
   es la que se busca en MS Project, con dónde está (`dentro de 2.3.0.1`, o el nivel) y qué abarca
   (tarea de detalle, tarea resumen con cuántas de sus subtareas están en la lista, o *rama completa
   (25 tareas)*, que no se despliega). Debajo, con sangría, las tareas de esa rama que también están en
   la lista, en su jerarquía; en texto normal, sin negrita, las tareas resumen intermedias que no
   están. Por ejemplo, para una línea base adicional:

   ```markdown
   - **`1.0.0`** — nivel 1; tarea resumen; de sus 12 subtareas, 9 la tienen
       - `1.1.0` — no la tiene; de sus 11 subtareas, 9 la tienen
           - **`1.1.3`**
           - **`1.1.4`**
           - ...
   - **`3.0.0`** — nivel 1; rama completa (86 tareas)
   - **`2.2.3.1`** — dentro de `2.2.3`; tarea de detalle
   ```
5. **Comparación de planificaciones**, si se revisaron varias.

El informe **no incluye nombres** de tareas, de empresas ni de calendarios, ni el título del proyecto: las
tareas se identifican por su código EDT o, sin código, por su UniqueID, y en los textos de las
observaciones generales los nombres se reemplazan por `…`. Las empresas de multi-empresa aparecen como
«empresa A» y «empresa B». Así se puede compartir con quien no debe ver el contenido del cronograma. La
salida en pantalla sí muestra los nombres. Los archivos `*_mejoras.md` están en el `.gitignore` del
repositorio.

---

## `generar`: crea el libro con el PV

```powershell
python evm_export.py generar --origen "C:\Programa\Programa.mpp" --config config.toml --salida EVM_Programa.xlsx
```

| Opción | Descripción |
|---|---|
| `--origen` | Archivo `.mpp` (requiere MS Project), `.xml` de MS Project, o `activo` para usar el proyecto abierto en MS Project. **Obligatoria**, salvo con varios archivos (`[[proyectos]]` en la configuración): en ese caso se omite, o se indica `NOMBRE=RUTA` para reemplazar la ruta de un archivo ([Varios archivos](varios-archivos.md)). `NOMBRE=RUTA` se puede repetir. Una sola ruta, sin `NOMBRE=`, tiene prioridad: se lee solo ese archivo, con la planificación y la línea base de su sección si es uno de `[[proyectos]]` ([CFG-057](codigos.md#cfg-057)), en todos los comandos que leen MS Project; una ruta relativa se toma desde la carpeta actual (no desde la del archivo de configuración). |
| `--salida` | Libro `.xlsx` a crear. Por defecto `<origen>_EVM.xlsx` (con `--origen activo`, `EVM.xlsx` en la carpeta actual); obligatoria con varios archivos. |
| `--config` | Archivo de configuración TOML ([Configuración](configuracion.md)). |
| `--mes-inicio AAAA-MM` | Primer mes del horizonte (mes 1). Por defecto, `mes_inicio` de la configuración o, si no se indica, el primer mes con % planificado (en general, el de inicio de la línea base). |
| `--meses N` | Cantidad de meses (columnas). Por defecto, `meses` de la configuración (120). |
| `--planificacion CODIGO` | Código EDT de la planificación a exportar (`INTERNO`, `CLIENTE`...). Reemplaza el valor del archivo de configuración; `""` exporta todas. Con varios archivos no se admite (la planificación se indica en cada `[[proyectos]]`), salvo con una sola ruta en `--origen`, que lee solo ese archivo ([CFG-057](codigos.md#cfg-057)). |
| `--linea-base N` | 0 = Línea base, 1..10 = Línea base N. Por defecto, la que el proyecto usa para el valor acumulado. Con varios archivos no se admite (se indica en cada `[[proyectos]]`), salvo con una sola ruta en `--origen`. |
| `--nivel-max N` | Nivel máximo del código EDT. Por defecto, `nivel_maximo` de la configuración (3). Con `lista_edt`, las filas las decide la lista y el nivel solo se usa para completar los códigos con ceros ([Lista EDT](calculo-pv-ev.md#34-lista-edt-filas-definidas-en-un-archivo)). |
| `--motor auto\|com\|xml` | Forma de lectura. `auto` decide por la extensión. |
| `--migrar-desde libro.xlsx` | Copia al libro nuevo el BAC ingresado, el % ganado y el AC ya cargados, los métodos editados en la columna C, el análisis de desviaciones y la fecha de corte de un libro anterior ([El libro Excel](libro-excel.md#4-cambios-de-línea-base-o-de-edt---migrar-desde)). Con depuración de costes, también la previsión pegada en `PV-depuración` ([Depuración de costes](depuracion-costes.md#9-ciclo-de-ajuste)). El % planificado es siempre el de MS Project. |
| `--sobrescribir` | Reemplaza el libro de salida si ya existe. |
| `--ev-rellenar` | Si el archivo de MS Project ya tiene avance y fecha de estado, carga también el % ganado de cada mes hasta la fecha de estado (ver abajo). Por defecto el % ganado se carga con `actualizar`. Equivale a `[exportacion] ev_historico = true` ([Configuración](configuracion.md#claves)). |
| `--pv-fechas-actuales` | Las tareas sin línea base forman el % planificado con sus fechas actuales, como antes de la versión 2.7.0. Por defecto, en una fila con otras tareas con línea base, quedan fuera del PV y solo cuentan en el % ganado (retrabajo o trabajo no previsto; [Tareas sin línea base](calculo-pv-ev.md#tareas-sin-línea-base-retrabajo-y-trabajo-no-previsto)). |
| `--visible` | Muestra la ventana de MS Project mientras se lee (solo COM). No tiene relación con el campo `Visible`. |
| `--silenciar CODIGOS` | Códigos de las observaciones que no se muestran, separados por comas (`LOG-003,FEC-006`, o `MET-004:1.2.1` para una sola fila); se puede repetir y se suma a `[validacion] silenciar` ([Silenciar observaciones](codigos.md#2-silenciar-observaciones)). |

Con `--ev-rellenar` (o `ev_historico = true`), si el archivo de MS Project ya tiene avance y fecha de
estado (un proyecto en ejecución), `generar` también carga el % ganado de cada mes hasta la fecha de
estado, que pasa a ser la fecha de corte del `Resumen`, y lo indica en pantalla (`% ganado: cargado de
01/2026 a 11/2026`). El de los meses anteriores se reconstruye con las fechas reales
([Del % acumulado al % del mes](calculo-pv-ev.md#7-del--acumulado-al--del-mes)).
Con `--migrar-desde` se conserva el del libro anterior. Sin `--ev-rellenar` ni `ev_historico = true`, el
libro se genera sin % ganado y se carga con `actualizar`.

Ejemplo de salida con el proyecto de ejemplo (se omiten algunas líneas):

```text
Exportador EVM 2.7.2
Leyendo ejemplos/programa_ejemplo.xml ...
Proyecto: Programa de ejemplo EVM
  Planificación: INTERNO
  Línea base  |  PV: % planificado de MS Project (duración de la línea base y peso de las tareas) × BAC ingresado en el libro
  Filas EDT: 36 (24 de ingreso)  |  Nivel máximo: 3
  Horizonte: 01/2026 a 12/2035 (120 meses)
  BAC: se debe ingresar el BAC de cada fila en la columna D de PV (celdas amarillas).
  [i] LIB-001 El PV y el EV de cada fila son el BAC × el % planificado y el % ganado (hojas % PV y % EV), que se calculan con la duración de la línea base y el peso de las tareas. Se debe ingresar el BAC de cada fila en las celdas amarillas de la columna D de la hoja PV.
  [i] EDT-017 Planificación 'INTERNO' («Planificación interna»): se exportan las 51 tareas que dependen de ella.
  [i] EDT-020 Se omitió 1 tarea sin código EDT que no depende de ningún elemento EDT: no forma parte de ninguna fila del libro.
  [i] EDT-022 14 tareas sin código EDT o de nivel EDT mayor a 3 se agregaron en su elemento EDT superior.
  [i] MET-005 1 fila sin método de medición definido (columna C).
  [!] MET-004 1.2.1 Licitación de ingeniería: Método 0/100 en un paquete de 3,6 meses (línea base), más de 1: el EV queda en 0 hasta que termina. Se recomienda dividir el paquete o usar un método de avance (% completado, hitos ponderados). Umbral: [validacion] meses_0_100.
  ...
  Por alcance: generales: 5 informativas; en filas EDT: 6 advertencias.
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
| `--origen` | Proyecto con el avance a la fecha de corte (`.mpp`, `.xml` o `activo`). **Obligatoria**, salvo con varios archivos: se leen todos los de `[[proyectos]]`, y `NOMBRE=RUTA` reemplaza la ruta de un archivo (relativa a la carpeta actual; se puede repetir). Una sola ruta sin `NOMBRE=` lee solo ese archivo ([CFG-057](codigos.md#cfg-057)): un libro generado con varios archivos no se actualiza con uno solo (`ACT-002`, sin modificarlo). |
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
  de los rangos admisibles en el mes de corte, y conserva lo ya escrito ([Análisis de desviaciones](analisis-desviaciones.md)).
  Avisa si el análisis del mes anterior quedó incompleto o si hay acciones vencidas.
- Registra los cambios del BAC y de los valores de meses ya informados (% ganado y AC) desde la ejecución
  anterior en la hoja `Registro de cambios` ([El libro Excel](libro-excel.md#registro-de-cambios)), y
  revisa la coherencia del EV y el AC de cada paquete, los meses sin informe, el realismo del EAC (TCPI
  frente al CPI) y el peso del LOE en el BAC
  ([Cálculo del PV y el EV](calculo-pv-ev.md#14-controles-y-validaciones)).
- Avisa si el % planificado que resulta hoy de MS Project no coincide con el del libro: se guardó una
  línea base nueva y el libro sigue con la anterior hasta generar uno nuevo con `--migrar-desde`
  ([ACT-047](codigos.md#act-047)).
- Escribe el plazo según el cronograma de MS Project: columnas Q a U de `Indicadores EDT` (fin previsto,
  desvío de fin, holgura y crítica de cada elemento) y la sección de plazo del `Resumen`
  ([El libro Excel](libro-excel.md#indicadores-edt)).
- Al terminar, muestra los paquetes fuera de umbral, los indicadores a la fecha de corte (con los mismos
  cálculos que la hoja `Resumen`, incluida la tendencia) y el plazo del proyecto.

Ejemplo de salida:

```text
Exportador EVM 2.7.2
Leyendo ...\avance_2027-06.xml ...
Fecha de corte: 30/06/2027 (mes 18 del horizonte)
Observaciones por fila:
  [i] ACT-034 4.2.2 Cierre administrativo: método no definido: se usó % completado
  [!] ACT-044 13 paquetes fuera de umbral en 05/2027 quedaron sin analizar, sin texto en la hoja Análisis de desviaciones (1.1.1, 1.1.2, 1.1.3, 1.2.1, 1.2.2, 1.2.3, 2.1.1, 2.1.2, 2.1.3, 2.1.4 y 3 más): se quitan de la hoja, y los que siguen fuera de umbral vuelven a listarse en el mes de corte.
  [i] ACT-026 El AC se carga a mano en la hoja 'AC'; los indicadores de coste usan los valores cargados.
Paquetes fuera de umbral (SPI o CPI fuera de 0,95 a 1,05, o SV %, CV % o desviación del avance fuera de su rango verde): 13 (hoja 'Análisis de desviaciones')
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
de umbral ([Análisis de desviaciones](analisis-desviaciones.md#3-cierre-mensual)). Como `actualizar`,
registra los cambios del BAC y de meses ya informados y revisa la coherencia del EV y el AC (con el AC
recién cargado), el análisis del mes anterior, los meses sin informe, el realismo del EAC y el peso del
LOE.

| Opción | Descripción |
|---|---|
| `--libro` | Libro creado con `generar`, con al menos un `actualizar`. **Obligatoria.** |
| `--sin-respaldo` | No crea la copia de respaldo del libro. |
| `--config` | Solo se usan `[validacion] silenciar` y `loe_maximo` ([ACT-049](codigos.md#act-049)) y `[exportacion] guardado_rapido`: el libro guarda los rangos admisibles al generarse. Sin `--config`, `loe_maximo` vale 15 y el guardado rápido está activado. |
| `--silenciar CODIGOS` | Igual que en `generar`. |

---

## `procesar-costes`: revisión y cálculos del plan de costes

```powershell
python evm_export.py procesar-costes --libro EVM_Depuracion.xlsx
python evm_export.py procesar-costes --libro EVM_Depuracion.xlsx --calculos
python evm_export.py procesar-costes --libro EVM_Depuracion.xlsx --calculos --config config_depuracion.toml
```

En un libro generado con `depuracion_costes = true`, revisa el plan de costes de la hoja `PV-depuración`
y, con `--calculos`, escribe sus cálculos, **sin leer MS Project**. **No modifica el % planificado**: el
PV de todas las filas es el de MS Project, y el plan se compara con él en las filas «(coste)» de la hoja
PV de la empresa y en la hoja `Depuración costes`
([Depuración de costes](depuracion-costes.md#8-revisión-y-cálculos-del-plan-procesar-costes)).

- Lee solo las **líneas de EDT cliente** cuyo código es el de una fila de ingreso de la hoja PV de la
  empresa ([qué líneas se leen](depuracion-costes.md#42-qué-líneas-se-leen)), como las filas «(coste)».
- Muestra cada fila con plan, con el **EAC** de su línea y su **BAC**, y las observaciones: líneas de EDT
  cliente sin leer, importes después del horizonte, filas sin línea en el plan o sin importes, EDT cliente
  repetidas e importes guardados como texto
  ([DEP-012 a DEP-026](codigos.md#413-dep--depuración-de-costes)).
- Con `--calculos` (o `[exportacion] depuracion_calculos = true`), escribe en la hoja las fórmulas de los
  cálculos del plan: las sumas de los componentes de coste en su EDT interna y de las EDT interna en su
  EDT cliente, el EAC de cada línea y el BAC de cada EDT cliente, leído de la hoja PV de la empresa
  ([cálculos del plan](depuracion-costes.md#43-cálculos-del-plan-celdas-naranjas)). Sin esa opción, no
  modifica el libro.
- Con `[depuracion_precio]` en la configuración (`--config`), los meses de cada EDT interna con
  componentes son su **precio**, con los recargos de la tabla `Recargos` de la hoja
  ([precio de las EDT interna](depuracion-costes.md#44-precio-de-las-edt-interna-tabla-recargos)); si la
  hoja no tiene la tabla, `--calculos` la agrega vacía. Avisa de los componentes que no están en la lista y
  de los coeficientes que faltan.
- Una hoja de la versión 2.5 o anterior, con el plan desde la celda A8, no se procesa: se genera un libro
  nuevo con `--migrar-desde`, que pasa el plan a la celda A12 ([DEP-024](codigos.md#dep-024)).
- En un libro procesado con una versión anterior a la 2.6.0 (que reemplazaba el % planificado por el
  reparto del plan), avisa de esas filas ([DEP-021](codigos.md#dep-021)): `generar --migrar-desde` les
  devuelve el % de MS Project.

```text
Exportador EVM 2.7.2
Cálculos del plan de costes: 16 líneas de la hoja PV-depuración (sumas, EAC y BAC de las EDT cliente)
Hoja: PV-EmpresaA  |  Plan de costes: 9 líneas de EDT cliente, 3 con fila de ingreso  |  Importe leído: 278
Plan de costes en 3 de 5 filas de ingreso, 278 dentro del horizonte:
  1.1.0        Dirección                                         EAC            117  BAC            117
  1.2.1        Licitación                                        EAC             65  BAC             65
  1.2.2        Obras                                             EAC             96  BAC            100
  [!] DEP-014 3 líneas de EDT cliente de la hoja PV-depuración sin leer, por 62,00: 1.2.7 (fila 25): sin fila del libro, 2.0.0 (fila 27): resumen sin detalle en el plan, 9.9 (fila 28): sin fila del libro. Su importe no llega a ninguna fila «(coste)».
  [i] DEP-016 Filas de ingreso de la hoja PV-EmpresaA sin línea de EDT cliente en el plan de costes (2): 1.0.0-D, 2.1.0. Su fila «(coste)» queda vacía.
El % planificado (hoja '% PV-EmpresaA') no cambia: es el de MS Project. El plan de costes se compara con el PV en las filas «(coste)» de la hoja PV-EmpresaA y en la hoja 'Depuración costes'.
Respaldo: C:\Programa\EVM_Depuracion.respaldo-20261108-101500.xlsx
Libro actualizado: C:\Programa\EVM_Depuracion.xlsx
```

| Opción | Descripción |
|---|---|
| `--libro` | Libro generado con `depuracion_costes = true`, con el plan de costes en la hoja `PV-depuración`. **Obligatoria.** |
| `--calculos` | Escribe los cálculos del plan en la hoja `PV-depuración` (fórmulas en las celdas naranjas) y guarda el libro. Equivale a `[exportacion] depuracion_calculos = true` para esta ejecución. Sin esta opción, el libro no se modifica. |
| `--simular` | Con `--calculos`, muestra el resultado sin modificar el libro. |
| `--sin-respaldo` | Con `--calculos`, no crea la copia de respaldo del libro. |
| `--config` | Solo se usan `[validacion] silenciar`, `[exportacion] guardado_rapido` y `depuracion_calculos` y `[depuracion_precio]`. Para el precio de las EDT interna, se indica la configuración de la depuración. |
| `--silenciar CODIGOS` | Igual que en `generar`. |

---

## `generar-fichas`: fichas de tareas para el cliente

```powershell
python evm_export.py generar-fichas --origen "C:\Programa\Programa.mpp" --config config.toml --salida Fichas_Programa.md
```

Crea un documento Markdown (`.md`) con una ficha por cada elemento EDT, para entregar al cliente. **No
crea ni modifica el libro.** Las fichas corresponden a las filas del libro (planificación,
`nivel_maximo`, campo `Visible` o lista EDT, tareas inactivas y varios archivos), sin las filas «-D». Cada ficha
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
- Los vínculos con tareas sin ficha (de otra planificación, inactivas, con `Peso` = 0, fuera de la
  lista EDT o de otro proyecto) no se muestran, para que el documento no incluya tareas que no se entregan. La consola
  indica cuántos se omitieron.
- No se indica el tipo de vínculo (fin a comienzo, comienzo a comienzo...) ni la posición o el adelanto.

| Opción | Descripción |
|---|---|
| `--salida` | Archivo `.md` a crear. Por defecto `<origen>_fichas.md` (con `--origen activo`, `Fichas.md` en la carpeta actual); obligatoria con varios archivos. |
| `--sobrescribir` | Reemplaza el archivo de salida si ya existe. |
| `--origen`, `--config`, `--motor`, `--planificacion`, `--linea-base`, `--nivel-max`, `--visible`, `--silenciar` | Igual que en `generar`. Con varios archivos (`[[proyectos]]`) se crea un solo documento con las fichas de todos, en el orden del libro. |

Ejemplo de salida con el proyecto de ejemplo, que no tiene vínculos entre tareas:

```text
Exportador EVM 2.7.2
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

## `generar-cliente`: planificación para el cliente

```powershell
python evm_export.py generar-cliente --origen "C:\Programa\INTERNO.mpp" --config config.toml --destino "C:\Entregas\CLIENTE 2026-10.mpp"
```

Genera, a partir de la planificación interna, el `.mpp` que se entrega al cliente: el cliente ve las
tareas con `Sí` en el campo de `[cliente] campo` (y en sus tareas superiores). Por defecto, las demás
quedan como «Otras tareas», sin nombre ni código EDT, con sus fechas y su avance: el % completado y el
libro EVM son los de la planificación interna. Con `[cliente] actividad_interna = false` se borran, y la
lógica se rehace sin ellas. El comando compara el % completado del proyecto y el libro EVM con los de la
planificación interna y avisa de las diferencias. La planificación interna no se modifica: MS Project
trabaja sobre una copia. El archivo del cliente queda de solo lectura, y junto a él se escribe un informe
de control interno (`<destino>_control.md`). Con `--simular` solo calcula los cambios y la comparación, y
escribe el informe; admite el XML. Ver [Planificación para el cliente](generar-cliente.md).

---

## `actualizar-edt`: numera los códigos EDT de una rama

```powershell
python evm_export.py actualizar-edt --origen activo --config config.toml --desde 2.7.1
```

Durante la planificación, al agregar, quitar o mover tareas, los códigos EDT de las tareas de detalle
dejan de seguir la jerarquía. `actualizar-edt` los vuelve a numerar a partir de una tarea: la tarea de
`--desde` conserva su código y sus subtareas se numeran en el orden del esquema de MS Project (los
hijos `2.7.1.1`, `2.7.1.2`...; los nietos `2.7.1.1.1`...). Las tareas sin código dentro de la rama
reciben uno. Solo se numeran ramas cuya tarea superior tiene un código numérico: los hitos del contrato,
sin código y fuera de las ramas con código, no cambian. Solo escribe en un campo personalizado de texto
(`[campos] edt`, por ejemplo `"Item EDT"`); para el campo EDT de MS Project se usa su función
*Renumerar*.

**Tareas protegidas.** El código de estas tareas nunca cambia, porque son filas del libro y el cliente
debe recibir siempre el mismo código:

- las de código de nivel 1 a `nivel_maximo`;
- las de `Visible = Sí`;
- las de la lista EDT (`lista_edt`), si se usa;
- las que contienen alguna de las anteriores.

`[cliente] campo` no protege: una tarea que ve el cliente pero que no es fila del libro se puede
renumerar.

**Orden.** Dentro de cada tarea, las subtareas protegidas van primero, con códigos correlativos desde 1
(`2.7.1.1`, `2.7.1.2`...), y las demás se numeran a continuación (`2.7.1.3`...): el cliente nunca debe
ver saltos. Por ejemplo, con `2.7.1.1` y `2.7.1.2` con `Visible = Sí`:

| Antes | Después | |
|---|---|---|
| `2.7.1.1` (Visible = Sí) | `2.7.1.1` | no cambia |
| `2.7.1.1.7` | `2.7.1.1.1` | subtarea de una protegida: se numera |
| `2.7.1.2` (Visible = Sí) | `2.7.1.2` | no cambia |
| `2.7.1.9` | `2.7.1.3` | |
| (sin código) | `2.7.1.4` | tarea nueva |

**Errores.** Si hay algún error, no se escribe ningún código:

- una tarea que se numera delante de una protegida ([NUM-008](codigos.md#num-008));
- códigos protegidos con saltos o desordenados ([NUM-009](codigos.md#num-009)): es una errata y se
  corrige a mano;
- un código protegido que no cuelga directamente de su tarea superior ([NUM-010](codigos.md#num-010)),
  por ejemplo una tarea copiada de otra rama con un código de nivel 3;
- una tarea sin código que contiene tareas protegidas ([NUM-011](codigos.md#num-011)), o una tarea sin
  código con `Visible = Sí` ([NUM-012](codigos.md#num-012)): su código lo decide el jefe de proyecto.

Un código nuevo que ya tiene una tarea de fuera de la rama se informa ([NUM-014](codigos.md#num-014)) y
la rama se actualiza igualmente.

**Desde un nivel superior.** Con `--desde 2.7` (nivel 2), los paquetes de nivel 3 (`2.7.1`, `2.7.2`...)
no cambian y se numeran sus subtareas. Nunca se crea un código de nivel 1 a `nivel_maximo` (sería una
fila nueva del libro): una tarea sin paquete que cuelga directamente de `2.7` recibe `2.7.0.1`, una
actividad de `2.7.0`. Con `ignorar_ceros_intermedios = true`, `2.7.0.1` sería de nivel 3: esas tareas
no se numeran y se informan ([NUM-013](codigos.md#num-013)).

**El archivo abierto en MS Project.** Se puede tener el archivo abierto: el programa usa la instancia
de MS Project en marcha.

- Con `--origen activo`, o con la ruta de un `.mpp` que ya está abierto, los códigos se escriben en
  el proyecto abierto y **no se guarda**: se revisan en MS Project y se guarda el archivo, o se cierra
  sin guardar para descartarlos. Todos los cambios forman una sola acción de deshacer: un Ctrl+Z los
  quita ([NUM-017](codigos.md#num-017) si MS Project no lo permite).
- Con la ruta de un `.mpp` cerrado, el programa lo abre, crea la copia
  `<archivo>.respaldo-AAAAMMDD-HHMMSS.mpp`, escribe los códigos, guarda el archivo y lo cierra.

Mientras trabaja el comando no se debe usar MS Project: editar una celda o dejar un diálogo abierto
bloquea al programa. Después de escribir, el programa vuelve a leer los códigos y los compara con los
previstos ([NUM-016](codigos.md#num-016)).

**Salida.** La consola muestra los cambios (Id, código actual, código nuevo y nombre) y las
observaciones. El informe `<origen>_edt-AAAAMMDD-HHMMSS.md` (cada ejecución tiene el suyo) tiene todos
los cambios y las tareas protegidas con su motivo. Incluye nombres de tareas: es para uso interno.

**Antes de usarlo.** Cambiar códigos por debajo de `nivel_maximo` afecta a lo que use esos códigos fuera
de MS Project: el plan de costes de la [depuración de costes](depuracion-costes.md), si sus EDT interna
son de ese nivel, y las fichas ya generadas. Los vínculos no cambian: van por Id.

| Opción | Descripción |
|---|---|
| `--desde CODIGO` | Código EDT de la tarea cuyas subtareas se numeran (su código no cambia). **Obligatoria.** Se puede repetir para numerar varias ramas, que no deben estar una dentro de otra. |
| `--origen` | `.mpp`, `activo` o, solo con `--simular`, `.xml`. Con varios archivos (`[[proyectos]]`) se omite: se usa la ruta de la sección de `--proyecto`; si se indica, reemplaza esa ruta (también `activo`). |
| `--proyecto NOMBRE` | Con varios archivos (`[[proyectos]]`): el archivo que se numera, por su `nombre`, con su planificación y su línea base. No hace falta si hay una sola sección ni si `--origen` es la ruta de una sección. Sin él, con varias secciones, un `--origen` que no es la ruta de ninguna se lee solo, con la planificación y la línea base generales ([CFG-057](codigos.md#cfg-057)). |
| `--simular` | Muestra los cambios y escribe el informe, sin escribir los códigos. Admite el XML. |
| `--sin-respaldo` | No crea la copia de respaldo del `.mpp` antes de guardarlo. |
| `--informe RUTA` | Informe en Markdown. Por defecto `<origen>_edt-AAAAMMDD-HHMMSS.md` junto al archivo (con `activo`, `EDT-AAAAMMDD-HHMMSS.md` en la carpeta actual). |
| `--sobrescribir` | Reemplaza el informe de `--informe` si ya existe. |
| `--config`, `--planificacion`, `--nivel-max`, `--motor`, `--visible`, `--silenciar` | Igual que en `generar`. Si el archivo tiene varias planificaciones con los mismos códigos (`INTERNO` y `CLIENTE`), se elige una con `--planificacion` (con `[[proyectos]]`, la de la sección). |

---

## `corregir-linea-base`: mueve, borra o establece la línea base de unas tareas

```powershell
python evm_export.py corregir-linea-base --origen activo --config config.toml --tarea 3.4.5 --rama --de 2 --a 10 --motivo "Tareas no comunicadas al cliente"
```

Corrige erratas en la línea base antes de generar el libro: por ejemplo, tareas nuevas que se guardaron en
la línea base del PV sin la aprobación del cliente. **Mover** (`--de N --a M`) copia los cinco campos de
la línea base N de cada tarea (comienzo, fin, duración, trabajo y costo previstos) a la M y deja la N
vacía (NA, o ND en MS Project en español); **borrar** (`--borrar N`) la deja vacía; **establecer**
(`--establecer N`) copia a la línea base N los valores actuales de las tareas que no la tienen, como
*Establecer línea base* de MS Project (por ejemplo, para el desglose de una tarea aprobada que no se
guardó en la línea base aprobada). El motivo se agrega a las Notas de cada tarea que cambia.

Antes de escribir, avisa de lo que conviene decidir: una fila que se queda sin PV, una tarea resumen de
encima que quedaría en el PV como un bloque, un destino con valores o un destino que es la línea base
del PV. Sin `--continuar` no escribe nada (código de salida 1); con `--interactivo` pregunta caso a caso y
escribe un cambio cada vez, que se puede deshacer. El archivo abierto y el `.mpp` cerrado se tratan como
en `actualizar-edt` (sin guardar y una sola acción de deshacer, o copia de respaldo y guardado).

Los conceptos, las buenas prácticas, las opciones y los avisos están en
[Líneas base en MS Project y en el programa EVM](lineas-base.md#parte-3-comando-corregir-linea-base).

| Opción | Descripción |
|---|---|
| `--tarea CODIGO` | Código EDT de la tarea. Se puede repetir; las tareas no deben estar una dentro de otra. Se indica `--tarea` o `--todas` ([LBA-022](codigos.md#lba-022)). |
| `--todas` | Todas las ramas de la planificación a las que afecta la acción (por ejemplo, para borrar una línea base entera): cada rama es un caso, con todas sus subtareas; con `--interactivo`, se revisan rama a rama, con un resumen de sus tareas. El motivo queda solo en el informe, salvo con `--notas`. |
| `--notas` | Con `--todas`, agrega también el motivo a las Notas de cada tarea que cambia. |
| `--rama` | Incluye todas las subtareas de cada `--tarea`. Una tarea resumen sin `--rama` es un error ([LBA-004](codigos.md#lba-004)). |
| `--de N --a M` / `--borrar N` / `--establecer N` | Mover de la línea base N a la M, borrar la N, o establecer la N con los valores actuales en las tareas que no la tienen (0 a 10). Se indica una sola acción. |
| `--motivo "texto"` | Queda en el informe y se agrega al final de las Notas de cada tarea que cambia (con `--todas`, solo con `--notas`). **Obligatoria** al escribir (con `--interactivo`, si falta, se pregunta). |
| `--aprobada N` | Línea base aprobada: si no es la del PV, avisa ([LBA-011](codigos.md#lba-011)). Con `--interactivo`, si falta, se pregunta. |
| `--interactivo` | Decide caso a caso y escribe un cambio cada vez, con la opción de deshacerlo. |
| `--continuar` | Escribe aunque haya avisos que piden confirmación. |
| `--simular` | Muestra los cambios y escribe el informe, sin escribir en MS Project. Admite el XML. |
| `--informe RUTA` | Informe en Markdown. Por defecto `<origen>_linea-base-AAAAMMDD-HHMMSS.md` junto al archivo (con `activo`, `LINEA-BASE-AAAAMMDD-HHMMSS.md` en la carpeta actual). |
| `--origen`, `--proyecto`, `--sin-respaldo`, `--sobrescribir`, `--config`, `--planificacion`, `--linea-base`, `--motor`, `--visible`, `--silenciar` | Como en `actualizar-edt`. `--linea-base` indica la línea base del PV. |

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

# Revisión del plan de costes de una empresa (hoja PV-depuración), con los cálculos del plan
python evm_export.py procesar-costes --libro EVM_Depuracion.xlsx --calculos

# Planificación: numerar las subtareas de 2.7.1 en el proyecto abierto (revisar en MS Project y guardar)
python evm_export.py actualizar-edt --origen activo --config config.toml --desde 2.7.1

# Errata: sacar de la línea base del PV (la 2) una rama nueva no aprobada, conservando sus fechas en la 10
python evm_export.py corregir-linea-base --origen activo --config config.toml --tarea 3.4.5 --rama --de 2 --a 10 --motivo "No comunicada al cliente" --interactivo

# Varios archivos de MS Project ([[proyectos]] en la configuración): sin --origen
python evm_export.py generar --config config.toml --salida EVM_Programa.xlsx
python evm_export.py actualizar --config config.toml --libro EVM_Programa.xlsx
python evm_export.py actualizar-edt --config config.toml --proyecto P1 --desde 2.7.1 --simular
```

Si no se activó el entorno virtual, se usa `.venv\Scripts\python.exe` en lugar de `python`.

---

## Códigos de salida

Todos los comandos terminan con un código que permite usarlos en procesos automáticos:

| Código | Significado |
|---|---|
| 0 | El comando terminó correctamente. En `validar`: no hay errores, aunque puede haber advertencias. |
| 1 | `validar` encontró errores (`[X]`) en el archivo de MS Project: se deben corregir antes de generar o actualizar el libro. En `actualizar-edt`: hay errores y no se escribió ningún código. En `corregir-linea-base`: hay avisos que piden confirmación y no se escribió nada. Ningún otro comando termina con 1. |
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
