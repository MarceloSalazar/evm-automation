# Planificación para el cliente

[Inicio](../README.md) › [Documentación](README.md) › Planificación para el cliente

El comando `generar-cliente` obtiene, a partir de la **planificación interna**, la planificación que se
entrega al cliente: un archivo `.mpp` en el que el cliente ve sus tareas, con la misma lógica, las mismas
fechas y el mismo avance. Así se mantiene **una sola planificación**: la del cliente no se edita a mano. Se
vuelve a generar en cada entrega, junto con su libro EVM.

---

## Contenido

1. [Preparación en MS Project](#1-preparación-en-ms-project)
2. [Qué tareas ve el cliente](#2-qué-tareas-ve-el-cliente)
3. [Las tareas que no se muestran: «Otras tareas»](#3-las-tareas-que-no-se-muestran-otras-tareas)
4. [Borrar las tareas que no se muestran](#4-borrar-las-tareas-que-no-se-muestran)
5. [Comparación con la planificación interna](#5-comparación-con-la-planificación-interna)
6. [Uso](#6-uso)
7. [Cómo se genera el archivo](#7-cómo-se-genera-el-archivo)
8. [Informe de control](#8-informe-de-control)
9. [Registro de diagnóstico](#9-registro-de-diagnóstico)

---

## 1. Preparación en MS Project

- Un campo personalizado de texto (por ejemplo *Texto15*, con el alias `Etiqueta Cliente`) indica, tarea
  por tarea, si el cliente la ve: `Sí` o `No`. Se configura en `[cliente] campo`.
- El código EDT está en un campo personalizado (`[campos] edt`, por ejemplo *Texto5*): así se puede quitar
  de las tareas que el cliente no ve. Si es el campo EDT de MS Project, esas tareas conservan su código
  (`CLI-027`).
- El archivo tiene **una sola planificación**, la interna: `[exportacion] planificacion` se deja vacía (o
  sin la clave), y el comando usa todo el proyecto.
- Solo si el archivo tuviera varias planificaciones (por ejemplo, tareas de primer nivel `INTERNO` y
  `CLIENTE`), se indica la interna en `[exportacion] planificacion`: las demás no pasan al archivo del
  cliente, y la tarea de la interna pasa a llamarse como indica `[cliente] nombre_planificacion` (por
  defecto `CLIENTE`).

## 2. Qué tareas ve el cliente

Una tarea se muestra si tiene `Sí` **y también lo tienen todas sus tareas superiores** de la
planificación. Valen como `Sí`: `Sí`, `Si` y `S`, en mayúsculas o minúsculas y con o sin tilde. Con `No`,
vacío u otro valor, la tarea no se muestra, con todas sus subtareas: una tarea nueva sin valor queda oculta
hasta que se marca.

| Caso | Resultado | Aviso |
|---|---|---|
| `Sí` en la tarea y en sus tareas superiores | Se muestra | — |
| `No` | No se muestra, con su rama | — |
| Vacío | No se muestra, con su rama | `CLI-031` si su tarea superior se muestra |
| Valor no reconocido (por ejemplo `Sí, revisar` o `X`) | No se muestra | `CLI-003` |
| `Sí` debajo de una tarea sin `Sí` | No se muestra | `CLI-004` |

Una tarea vacía debajo de una que se muestra suele ser una tarea nueva que nadie marcó: el cliente deja de
verla sin que nadie lo haya decidido. `CLI-031` la avisa en cada ejecución, también con `--simular`, con la
tarea de más arriba de cada rama (las inactivas no cuentan). Si el cliente no la debe ver, se le pone
`No`: así queda decidido y no se vuelve a avisar. Con `[cliente] campo` en la configuración, `validar`
revisa las mismas marcas (`CLI-003`, `CLI-004` y `CLI-031`), sin generar el archivo del cliente: así las
tareas nuevas sin marcar aparecen en la revisión mensual de la planificación interna.

Qué pasa con las tareas que no se muestran lo decide `[cliente] actividad_interna`:

| `actividad_interna` | Tareas que no se muestran | % completado y libro EVM |
|---|---|---|
| `true` (por defecto) | Quedan como «Otras tareas», sin nombre ni código EDT (apartado 3) | Los de la planificación interna |
| `false` | Se borran (apartado 4) | Pueden cambiar: se avisa con el detalle |

## 3. Las tareas que no se muestran: «Otras tareas»

MS Project calcula el % completado de un resumen, y el del proyecto, con la duración de **todas** sus
tareas de detalle: la duración realizada (duración × % completado) dividida por la duración total. El
libro EVM también calcula cada fila con sus tareas de detalle: el % planificado de cada mes sale de las
fechas de la línea base de cada una, y el % ganado, de su avance. Si se borraran las tareas internas, su
duración y su avance dejarían de contar, y los % cambiarían. Por eso, con `actividad_interna = true`, las
tareas que no se muestran se conservan, pero sin lo que dice qué son:

| Se quita | Se conserva |
|---|---|
| El nombre, que pasa a ser «Otras tareas» | Fechas, duración, delimitaciones y fecha límite |
| El código EDT | Líneas base 0 a 10 |
| Notas, hipervínculo y contacto | % completado, % físico completado, comienzo y fin reales |
| Los campos de texto y los códigos de esquema que no usa el libro EVM | Vínculos, y los campos de texto calculados con una fórmula (`CLI-030`): MS Project calcula su valor con los datos que se conservan |
| El valor del campo del cliente (en todas las tareas) | Los campos del libro EVM: método, Visible, peso de hito, empresa y reparto |
| Las tareas inactivas que no se muestran (no cuentan en ningún cálculo) | Recursos asignados y calendario de la tarea |

El cliente sabe que existen esas tareas y cuándo ocurren, pero no qué son. El % completado de cada tarea,
de los resúmenes y del proyecto es el de la planificación interna, y las fechas no cambian: no hacen falta
vínculos nuevos ni fechas fijadas.

Por defecto (`[cliente] filtrar_vista = true`), el archivo se abre con el filtro *Tareas del cliente*, que
oculta las «Otras tareas» (`CLI-028` si MS Project no lo acepta). El filtro solo cambia lo que se ve: el %
completado sigue contando todas las tareas, y el cliente puede quitarlo. Con `filtrar_vista = false`, el
archivo se abre sin filtro, con todas las tareas a la vista.

### El libro EVM del cliente

Las «Otras tareas» no tienen código EDT: en el libro forman parte de la fila de su elemento EDT superior.
Casi todas las filas quedan iguales, pero hay tres casos que se avisan:

- **Filas -D** (`CLI-025`). Un elemento EDT que no se muestra, dentro de uno que sí tiene filas propias,
  pasa a la fila -D de su superior («Otras tareas de …»). Por ejemplo, la fila `1.2.3` pasa a ser
  `1.2.0-D`, con las mismas tareas y valores. El cliente ve esa fila en su libro: para evitarla, se muestra
  ese elemento (`Sí`) o se reorganiza la EDT.
- **Filas que se reúnen** (`CLI-024`). Si no se muestra ningún elemento EDT de un resumen, la fila del
  resumen reúne a las de esos elementos.
- **Filas que desaparecen** (`CLI-024`). Un elemento EDT que no se muestra y no depende de ninguno que se
  muestre queda sin código: sus tareas no forman parte de ninguna fila del libro del cliente.

El informe de control indica, para cada caso, la fila del otro libro que tiene esas tareas.

## 4. Borrar las tareas que no se muestran

Con `actividad_interna = false`, las tareas que no se muestran se borran. El % completado de los
resúmenes y del proyecto, y el libro EVM, pueden cambiar: el programa lo calcula y lo avisa con el detalle
(apartado 5).

### Resúmenes sin tareas visibles

Si de un resumen no se muestra ninguna subtarea, en el archivo del cliente pasa a ser una tarea. Conserva
sus fechas, su duración, su avance y sus líneas base, y recibe los vínculos de sus subtareas (`CLI-005`).

Si una subtarea oculta fija el comienzo o el fin de un resumen que sí tiene subtareas visibles, el
resumen del cliente va de la primera a la última subtarea visible (`CLI-008`).

### Lógica del cronograma

Al quitar tareas, sus vínculos desaparecen. Para que la lógica siga completa y las fechas no cambien:

- Los vínculos entre tareas que se muestran se conservan.
- Los vínculos de las subtareas de un resumen convertido pasan al resumen, con su tipo y su posición.
- Si entre dos tareas que se muestran hay tareas ocultas, se crea un **vínculo puente**. Su posición es la
  suma de las posiciones del camino, más la duración de las tareas ocultas que el camino recorre de
  comienzo a fin, menos la de las que recorre de fin a comienzo. Por ejemplo, A → (tarea oculta de 3 días)
  → B pasa a ser A → B fin a comienzo con 3 días de posición.
- Un vínculo nuevo no debe retrasar a su sucesora. Si lo haría (por ejemplo, porque la tarea oculta ya
  terminó), su posición se ajusta a la que dan las fechas actuales. Si así quedaría con adelanto, el
  vínculo parte del comienzo de la predecesora, con una posición positiva.
- Entre dos tareas solo puede haber un vínculo: si hay varios caminos, se elige el que más restringe a la
  sucesora. No se crean vínculos entre tareas ya vinculadas, entre un resumen y sus subtareas, ni los que
  cerrarían un ciclo (`CLI-007`).
- Si el archivo tuviera otras planificaciones, sus vínculos desaparecen con ellas.

Con `[cliente] fijar_fechas = true` (por defecto), las fechas que aun así cambien al recalcular el
archivo del cliente se fijan con la delimitación *No comenzar antes del* y se informan (ver el
apartado 7).

### Avance de los resúmenes

- **Resumen convertido** (no se muestra ninguna subtarea): conserva el % que tenía, calculado con todas
  sus subtareas (`CLI-021` si no es así). En su resumen superior pesa por su propia duración (de su
  comienzo a su fin), no por la suma de la de sus subtareas: si estas van en paralelo, pesa menos.
- **Resumen con subtareas visibles y ocultas:** en el archivo del cliente, MS Project lo calcula solo con
  las tareas que se muestran. El avance de las tareas ocultas no cuenta, y el % del resumen cambia. El
  programa lo calcula antes de generar el archivo (también con `--simular`), lo compara después con el de
  MS Project y lista cada resumen con su % interno y el del cliente (`CLI-020`).
- **Tareas de detalle que se muestran:** conservan su % (`CLI-021` si no es así).

## 5. Comparación con la planificación interna

En los dos modos, el programa compara la planificación del cliente con la interna. Con `--simular` es
una previsión, calculada sin MS Project; al generar el archivo, lee el archivo del cliente en MS Project
después de guardarlo.

| Qué compara | Cómo | Aviso |
|---|---|---|
| % completado del proyecto | Como lo calcula MS Project, con las tareas de detalle activas | `CLI-023` |
| Libro EVM, fila a fila | Cada fila con % planificado y % ganado: las mismas tareas de detalle, el mismo % planificado de cada mes y el mismo % ganado con el avance actual. Una fila que cambia de código se empareja por sus tareas | `CLI-024`, `CLI-025` |
| Cada tarea (con «Otras tareas», al generar el archivo) | Comienzo, fin, comienzos y fines reales, % completado, % físico, duración, línea base, activa e hito, también de los resúmenes | `CLI-026` |
| No se pudo calcular un libro | Por ejemplo, por un error de la configuración del libro | `CLI-029` |

Con `actividad_interna = true`, todo debe coincidir, salvo los cambios de código del apartado 3. Con
`actividad_interna = false`, el informe de control detalla las diferencias.

## 6. Uso

```toml
[campos]
edt = "Item EDT"            # código EDT en un campo personalizado

[exportacion]
planificacion = ""          # una sola planificación: todo el proyecto

[cliente]
campo = "Etiqueta Cliente"
actividad_interna = true    # las tareas que no se muestran quedan como «Otras tareas»
filtrar_vista = true        # el archivo se abre con un filtro que oculta las «Otras tareas»
fijar_fechas = true
```

```powershell
# Ver los cambios sin generar nada (admite el .mpp o el XML)
python evm_export.py generar-cliente --origen INTERNO.mpp --config config.toml --simular
# Generar el archivo del cliente (necesita MS Project)
python evm_export.py generar-cliente --origen INTERNO.mpp --config config.toml --destino "CLIENTE 2026-10.mpp"
```

| Opción | Uso |
|---|---|
| `--origen` | Planificación interna: `.mpp`; con `--simular`, también `.xml` |
| `--destino` | Archivo `.mpp` del cliente (por defecto `<origen>_cliente.mpp`); no puede ser el origen |
| `--informe` | Informe de control (por defecto `<destino>_control.md`) |
| `--simular` | Calcula los cambios y la comparación, y escribe el informe de control, sin generar el archivo del cliente |
| `--sobrescribir` | Reemplaza el archivo del cliente y el informe de control si existen |

El libro EVM del cliente se genera y se actualiza con el archivo del cliente, con su propia
configuración: `planificacion = ""` (con varias planificaciones, el nombre de `nombre_planificacion`) y el
campo `Visible` de siempre. Con «Otras tareas», sus filas incluyen el trabajo interno (sin nombre), y sus
% son los del libro interno. Los UID de las tareas son los de la planificación interna y no cambian de una
entrega a otra, por lo que `actualizar` encuentra las filas en cada archivo nuevo. Para pasar a este libro
los meses de un libro anterior, se usa `generar --migrar-desde`.

## 7. Cómo se genera el archivo

La planificación interna **no se abre ni se modifica**: el programa la copia a un archivo temporal en la
carpeta de destino y MS Project trabaja sobre la copia.

1. Abre la copia y la lee como el resto del programa.
2. Con el cálculo manual y sin vincular tareas automáticamente, borra las tareas de otras planificaciones,
   si las hubiera, y:
   - con «Otras tareas»: las tareas inactivas que no se muestran;
   - sin «Otras tareas»: las tareas que no se muestran.

   Comprueba que quedan exactamente las tareas esperadas (si no, se detiene sin generar el archivo:
   `CLI-018`) y quita los vínculos que MS Project pudo crear al borrar tareas (`CLI-016`).
3. Con «Otras tareas», quita el nombre, el código EDT y los demás datos de las tareas que no se muestran
   (`CLI-027` si MS Project no deja vaciar un campo; los campos calculados con una fórmula no se pueden
   vaciar y se informan con `CLI-030`). Sin «Otras tareas», devuelve a cada resumen
   convertido su duración, su avance y sus líneas base (`CLI-019` si MS Project no lo admite) y crea los
   vínculos nuevos (`CLI-015` si MS Project rechaza alguno).
4. Vacía el campo del cliente en todas las tareas y quita su nombre (`CLI-017` si no puede). Con varias
   planificaciones, renombra además la interna.
5. Recalcula y compara las fechas de cada tarea con las de la planificación interna (también las inactivas:
   MS Project las sigue programando con sus predecesoras). Fija con *No comenzar
   antes del* las tareas sin comenzar que se adelantan (`CLI-013`) y quita los vínculos nuevos de las que
   se atrasan; repite el cálculo hasta que no hay cambios. Sin «Otras tareas», las que siguen distintas
   (comenzadas, con otra delimitación...) se informan (`CLI-014`) y se compara el % completado de cada
   tarea (apartado 4).
6. Con «Otras tareas» y `filtrar_vista = true`, aplica el filtro *Tareas del cliente* a la vista.
7. Revisa lo que no son tareas (apartado 7.1): quita los calendarios sin uso y las macros.
8. Guarda la copia, la lee y la compara con la planificación interna (apartado 5).
9. Deja la copia en `--destino` como archivo **de solo lectura**: no se edita, se vuelve a generar en cada
   entrega. Si se debe reemplazar, `--sobrescribir` lo hace. Por último, lee el archivo (sin MS Project) y
   comprueba que no le queda ningún módulo de macros.

Si la planificación interna está abierta en MS Project, se usa la versión guardada en el disco
(`CLI-012`). Si algo falla, no queda ningún archivo del cliente ni la copia temporal.

Para comprobar que MS Project no creó vínculos al borrar tareas, se lee de una vez el *Id. único de
predecesoras* de cada tarea; solo se recorren, vínculo a vínculo, los de las tareas cuyas predecesoras no
son las esperadas. En una planificación de unas 1.450 tareas, este paso pasa de casi un minuto a pocos
segundos.

### 7.1 Calendarios, macros y demás contenido del archivo

El archivo del cliente es una copia del `.mpp`: sin este paso llevaría todos los calendarios y las macros
de la planificación interna.

| Contenido | Qué se hace |
|---|---|
| Calendarios base | Se quitan los que no usa el proyecto, ninguna tarea (también las «Otras tareas») ni ningún recurso. Se conserva el que crea MS Project para las líneas base de Project 98. Si no se puede leer el calendario del proyecto, no se quita ninguno |
| Macros (módulos de VBA) | Se quitan todos los módulos y el código de *ThisProject*, con el proyecto VBA. Si MS Project no da acceso a él (Centro de confianza), se quitan con el Organizador, por el nombre de los módulos leídos del archivo |
| Propiedades del archivo (título, autor...), vistas, tablas, filtros y grupos | Se conservan. El informe de control indica qué propiedades tienen valor (sin el valor) y los nombres de las vistas, tablas, filtros y grupos |

Lo que se quitó se informa con `CLI-032`. Si MS Project no deja quitar un calendario, o el archivo guardado
todavía tiene macros, se avisa con `CLI-033`: se pueden quitar a mano con el Organizador (*Archivo >
Información > Organizador*).

## 8. Informe de control

Documento interno en Markdown, que **no se entrega al cliente**: identifica las tareas por su código EDT y
su UID (Id. único de MS Project), sin sus nombres. Contiene:

- el modo (`actividad_interna`) y el resumen: tareas que se muestran y que no, por motivo; «Otras tareas»
  y tareas inactivas que se quitan, o resúmenes convertidos; tareas de otras planificaciones; vínculos que
  se conservan, desaparecen y se crean;
- las observaciones;
- la comparación con la planificación interna: el % completado del proyecto, las filas del libro EVM
  (iguales, con otro código, con otras tareas o valores, y dónde están sus tareas en el otro libro), las
  filas -D y, al generar el archivo con «Otras tareas», las tareas con datos distintos;
- con «Otras tareas», la tarea de más arriba de cada rama que no se muestra;
- las tareas con `Sí` que no se muestran y los valores no reconocidos;
- sin «Otras tareas»: los resúmenes que pasan a ser tareas, cada vínculo nuevo (predecesora, sucesora,
  tipo, posición en días y origen), los vínculos que no se crean, los resúmenes que cambian de fechas y
  los que tienen otro % completado;
- las tareas sin valor debajo de una tarea que se muestra (`CLI-031`), con la tarea superior y el número
  de tareas de su rama;
- al generar el archivo: lo que hizo MS Project (fechas fijadas, campos que no dejó vaciar, filtro,
  vínculos que no aceptó o que creó al borrar tareas, nombre del campo del cliente) y el contenido del
  archivo (calendarios que se conservan, con su uso, y los quitados; macros; propiedades con valor; vistas,
  tablas, filtros y grupos).

## 9. Registro de diagnóstico

Para revisar un problema, la opción `--registro` (antes del comando) escribe el registro de diagnóstico
completo en un archivo, con la fecha y la hora de cada paso y el detalle técnico de los errores. La consola
muestra lo mismo que sin la opción.

```powershell
python evm_export.py --registro generar-cliente.log generar-cliente --origen INTERNO.mpp --config config.toml
```

El registro indica la versión del programa, de Python y de MS Project, cada paso de la generación, las
ramas borradas, los datos quitados de las «Otras tareas» por campo, los vínculos creados (tipo y posición
pedidos y los que tiene MS Project), las fechas fijadas, los calendarios y las macros quitados, la duración
de cada etapa y el resultado de la comparación. Identifica las
tareas por su código EDT y su UID, sin sus nombres: los errores de MS Project se registran sin el texto
entre comillas, donde MS Project cita los nombres.
