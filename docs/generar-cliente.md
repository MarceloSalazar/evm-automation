# Planificación para el cliente

[Inicio](../README.md) › [Documentación](README.md) › Planificación para el cliente

El comando `generar-cliente` obtiene, a partir de la **planificación interna**, la planificación que se
entrega al cliente: un archivo `.mpp` con solo las tareas que el cliente debe ver, con la misma lógica, las
mismas fechas y el mismo avance. Así se mantiene **una sola planificación**: la del cliente no se edita a
mano. Se vuelve a generar en cada entrega, junto con su libro EVM.

---

## Contenido

1. [Preparación en MS Project](#1-preparación-en-ms-project)
2. [Qué tareas ve el cliente](#2-qué-tareas-ve-el-cliente)
3. [Resúmenes sin tareas visibles](#3-resúmenes-sin-tareas-visibles)
4. [Lógica del cronograma](#4-lógica-del-cronograma)
5. [Uso](#5-uso)
6. [Cómo se genera el archivo](#6-cómo-se-genera-el-archivo)
7. [Informe de control](#7-informe-de-control)
8. [Registro de diagnóstico](#8-registro-de-diagnóstico)

---

## 1. Preparación en MS Project

- Un campo personalizado de texto (por ejemplo *Texto15*, con el alias `Etiqueta Cliente`) indica, tarea
  por tarea, si el cliente la ve: `Sí` o `No`. Se configura en `[cliente] campo`.
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
| `No` o vacío | No se muestra, con su rama | — |
| Valor no reconocido (por ejemplo `Sí, revisar` o `X`) | No se muestra | `CLI-003` |
| `Sí` debajo de una tarea sin `Sí` | No se muestra | `CLI-004` |

## 3. Resúmenes sin tareas visibles

Si de un resumen no se muestra ninguna subtarea, en el archivo del cliente pasa a ser una tarea. Conserva
sus fechas, su duración, su avance y sus líneas base, y recibe los vínculos de sus subtareas (`CLI-005`).

Si una subtarea oculta fija el comienzo o el fin de un resumen que sí tiene subtareas visibles, el
resumen del cliente va de la primera a la última subtarea visible (`CLI-008`).

## 4. Lógica del cronograma

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
apartado 6).

## 5. Uso

```toml
[exportacion]
planificacion = ""          # una sola planificación: todo el proyecto

[cliente]
campo = "Etiqueta Cliente"
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
| `--simular` | Calcula los cambios y escribe el informe de control, sin generar el archivo del cliente |
| `--sobrescribir` | Reemplaza el archivo del cliente y el informe de control si existen |

El libro EVM del cliente se genera y se actualiza con el archivo del cliente, con su propia
configuración: `planificacion = ""` (con varias planificaciones, el nombre de `nombre_planificacion`) y el
campo `Visible` de siempre. Así, el libro solo tiene las tareas que ve el cliente y, de ellas, las que indica
`Visible`. Los UID de las tareas son los de la planificación interna y no cambian de una entrega a otra,
por lo que `actualizar` encuentra las filas en cada archivo nuevo. Para pasar a este libro los meses de
un libro anterior, se usa `generar --migrar-desde`.

## 6. Cómo se genera el archivo

La planificación interna **no se abre ni se modifica**: el programa la copia a un archivo temporal en la
carpeta de destino y MS Project trabaja sobre la copia.

1. Abre la copia y la lee como el resto del programa.
2. Con el cálculo manual y sin vincular tareas automáticamente, borra las tareas que no se muestran (y las
   de otras planificaciones, si las hubiera), y comprueba que quedan exactamente las tareas que se muestran (si no, se
   detiene sin generar el archivo: `CLI-018`).
3. Devuelve a cada resumen convertido su duración, su avance y sus líneas base (`CLI-019` si MS Project no
   lo admite).
4. Quita los vínculos que MS Project pudo crear al borrar tareas (`CLI-016`) y crea los vínculos nuevos
   (`CLI-015` si MS Project rechaza alguno).
5. Vacía el campo del cliente en todas las tareas y quita su nombre (`CLI-017` si no puede). Con varias
   planificaciones, renombra además la interna.
6. Recalcula y compara las fechas de cada tarea con las de la planificación interna. Fija con *No comenzar
   antes del* las tareas sin comenzar que se adelantan (`CLI-013`) y quita los vínculos nuevos de las que
   se atrasan; repite el cálculo hasta que no hay cambios. Las que siguen distintas (comenzadas, con otra
   delimitación...) se informan (`CLI-014`).
7. Guarda la copia y la deja en `--destino` como archivo **de solo lectura**: no se edita, se vuelve a
   generar en cada entrega. Si se debe reemplazar, `--sobrescribir` lo hace.

Si la planificación interna está abierta en MS Project, se usa la versión guardada en el disco
(`CLI-012`). Si algo falla, no queda ningún archivo del cliente ni la copia temporal.

## 7. Informe de control

Documento interno en Markdown, que **no se entrega al cliente**: identifica las tareas por su código EDT y
su UID (Id. único de MS Project), sin sus nombres. Contiene:

- el resumen: tareas que se muestran y que no, por motivo; tareas de otras planificaciones; vínculos que se
  conservan, desaparecen y se crean;
- las observaciones (`CLI-003` a `CLI-008`);
- las tareas con `Sí` que no se muestran y los valores no reconocidos;
- los resúmenes que pasan a ser tareas;
- cada vínculo nuevo: predecesora, sucesora, tipo, posición en días y origen (puente o resumen);
- los vínculos que no se crean y los resúmenes que cambian de fechas;
- al generar el archivo: las fechas fijadas, las fechas que siguen distintas (con el motivo), los vínculos
  que MS Project no aceptó o creó al borrar tareas y si se quitó el nombre del campo.

## 8. Registro de diagnóstico

Para revisar un problema, la opción `--registro` (antes del comando) escribe el registro de diagnóstico
completo en un archivo, con la fecha y la hora de cada paso y el detalle técnico de los errores. La consola
muestra lo mismo que sin la opción.

```powershell
python evm_export.py --registro generar-cliente.log generar-cliente --origen INTERNO.mpp --config config.toml
```

El registro indica la versión del programa, de Python y de MS Project, cada paso de la generación, las
ramas borradas, los vínculos creados (tipo y posición pedidos y los que tiene MS Project), las fechas
fijadas y las que siguen distintas. Identifica las tareas por su código EDT y su UID, sin sus nombres: los
errores de MS Project se registran sin el texto entre comillas, donde MS Project cita los nombres.
