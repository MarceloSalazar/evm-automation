# Guía de convenios para planificar en MS Project

Guía para **jefes de proyecto y planners**. Reúne las reglas comunes para trabajar el cronograma en MS
Project de forma que sirva, sin trabajo adicional, para el **informe EVM** (informe mensual de valor
ganado) que se entrega al cliente. Cada regla indica **qué** hacer y **por qué** se decidió así. No hace
falta conocer la herramienta que prepara el informe EVM.

> Los nombres de los campos (`Item EDT`, `Visible EVM`, `Peso`, `Método EVM`...) son los habituales; en
> un proyecto pueden tener otro nombre. Los nombres de menú son los de MS Project en español.

## 1. Por qué importan estas reglas

El informe EVM compara, para cada elemento de la EDT y cada mes, tres valores:

| Valor | Significado | De dónde sale |
|---|---|---|
| **PV** (*Planned Value*) | Trabajo que se planificó hacer | Fechas de la **línea base** y peso de cada tarea |
| **EV** (*Earned Value*) | Trabajo realmente hecho | **Avance** registrado en MS Project, según el método de medición |
| **AC** (*Actual Cost*) | Costo real | Contabilidad (no se carga en MS Project) |

El presupuesto (**BAC**) de cada elemento se carga en el informe EVM, no en MS Project. Del cronograma
solo se toma **cómo se reparte el trabajo en el tiempo y cuánto se avanzó**. Por eso, la estructura, la
línea base, los pesos y el avance tienen que seguir las mismas reglas en todos los proyectos: si no, el
informe EVM muestra desvíos que no existen o esconde los reales.

## 2. Código EDT (`Item EDT`)

**Reglas:**

- El código es numérico, con puntos: `2`, `2.7`, `2.7.1`, `2.7.1.3`. Se escribe en el campo `Item EDT`,
  no en el campo *EDT* de MS Project.
- Cada código empieza por el de su tarea superior y tiene un nivel más: bajo `2.7.1` van `2.7.1.1`,
  `2.7.1.2`...
- Sin duplicados y sin saltos: `2.7.1.1`, `2.7.1.2`, `2.7.1.3`, en el orden del esquema.
- Los ceros finales completan el código hasta el nivel 3 y no cuentan: `2.0.0` es el nivel 1 y `2.7.0`,
  el nivel 2. Un 0 intermedio indica una actividad que cuelga directamente de un elemento sin paquete:
  `2.7.0.1` es una actividad de `2.7`.
- Los niveles 1 a 3 son, en general, los **elementos del informe EVM** (el nivel 3 es el paquete de
  trabajo). Los niveles inferiores son el detalle de cada paquete. Si el proyecto usa el campo
  `Visible EVM` ([sección 3](#3-campo-visible-evm-qué-tiene-fila-propia-en-el-informe-evm)), los
  elementos del informe EVM son los marcados con `Sí`.
- Todas las tareas de detalle y los hitos de un paquete llevan código.
- Los hitos del contrato que no forman parte de ningún elemento de la EDT no llevan código.

**Por qué:** el código, y no la posición en el esquema, decide a qué elemento del informe EVM pertenece
cada tarea. Los códigos de los elementos del informe EVM son los que ve el cliente mes a mes: **una vez
entregados, no cambian nunca**. Los códigos del detalle sí se pueden renumerar durante la planificación;
el responsable del informe EVM puede hacerlo automáticamente para una rama, sin tocar los códigos de los
elementos del informe.

## 3. Campo `Visible EVM`: qué tiene fila propia en el informe EVM

| Valor | Significado |
|---|---|
| **Sí** | La tarea es un elemento del informe EVM, con fila propia, de cualquier nivel |
| vacío o No | La tarea no tiene fila propia: su trabajo se suma en el elemento superior con `Sí` |

Si ninguna tarea tiene `Sí`, los elementos del informe EVM son los de nivel 1 a 3. Si alguna lo tiene,
el campo decide: todos los elementos del informe EVM, también los de nivel 1 a 3, llevan `Sí`.

**Reglas:**

- El campo es de **texto** (con lista `Sí`, `No`), no una marca Sí/No: necesita el estado vacío.
- Dentro de cada tarea, las subtareas con `Sí` van **primero** y con códigos correlativos desde 1. Las
  demás van a continuación.
- Una tarea con `Sí` necesita código EDT.

**Por qué:** las tareas con `Sí` llegan al cliente con su código. Si estuvieran mezcladas con el detalle,
cualquier renumeración del detalle obligaría a cambiar sus códigos o dejaría saltos que el cliente vería.

## 4. Campo `Peso`: cuánto vale cada tarea dentro de su tarea superior

En los paquetes con método de medición `Ponderado`, el peso de cada subtarea es un **% de su tarea
superior**; en `Hitos`, el de cada hito es un % del paquete. El 0 vale en cualquier paquete,
sea cual sea su método de medición:

| Peso | Significado |
|---|---|
| vacío | **Sin definir**: la tarea depende de sus hermanas (ver abajo) |
| número mayor que 0 | % de la tarea superior |
| **0** | Peso cero a propósito: la tarea **se hace, pero no cuenta** en el informe EVM |

**Reglas:**

- Los pesos de las hermanas suman **100** o menos. Si suman menos de 100, el resto se reparte entre las
  hermanas con el peso vacío, según su duración. Si suman 100, las de peso vacío no reciben nada. Más de
  100 es un error.
- Un hito sin peso no recibe nada (dura 0): si debe valer, se le pone peso.
- El campo es de **texto**: en un campo numérico, MS Project no distingue un 0 de un campo vacío.
- Una tarea con 0 queda fuera con todas sus subtareas: no suma en su paquete ni forma una fila aparte de
  «otras tareas». Un elemento del informe EVM con 0 se conserva, con valor 0.
- El motivo de un peso 0 se anota en las **Notas** de la tarea ([sección 10](#10-notas-de-la-tarea)).

**Por qué:** el vacío permite cargar solo los pesos que importan (por ejemplo, los entregables) y dejar
que el resto se reparta por duración. El 0 permite planificar tareas necesarias (coordinación interna,
trámites) sin que alteren el avance que se informa.

## 5. Método de medición (`Método EVM`)

Cada paquete indica su **método de medición**: la regla con que se calcula su avance (EV) a partir de lo
registrado en MS Project. Se elige con el responsable del informe EVM y se aprueba con la línea base.
Por ahora se usan estos métodos de medición:

| Método de medición | Cómo avanza | Cuándo usarlo |
|---|---|---|
| `0/100` | 0 hasta terminar; todo al terminar | Paquetes que empiezan y terminan en el mismo mes |
| `Hitos` | Cada hito vale su `Peso` al cumplirse | Tres meses o más, con entregables intermedios (al menos un hito por mes) |
| `% Completado` | El *% completado* de sus tareas, ponderado por su duración | Trabajo continuo sin entregables intermedios |
| `Ponderado` | Cada subtarea con su `Peso` y su propio método de medición | Paquetes con subtareas de distinta naturaleza (por ejemplo, ingeniería, compras y entregables) |
| `Avance lineal` | Avanza con el tiempo, igual que lo planificado | Solo soporte sin entregables (gestión, apoyo) |

En un paquete `Ponderado`, cada subtarea indica en `Método EVM` su propio método de medición (`0/100`,
`% Completado` o `Avance lineal`; sin indicarlo, `% Completado`). Un hito cuenta al cumplirse.

**Por qué:** el método de medición se acuerda antes para que el avance no dependa de la opinión de cada
mes. El avance lineal se limita porque siempre parece «en plazo» y oculta los atrasos.

## 6. Empresas (proyectos entre dos empresas)

| `Empresa` | Significado |
|---|---|
| Empresa A o empresa B | Todo el trabajo de la tarea es de esa empresa |
| Ambas | Se reparte según `% Reparto` |
| vacío | Igual que su tarea superior |

- `% Reparto` es el **% de la empresa A**: `30` significa 30 % para A y 70 % para B. Solo se usa en tareas
  `Ambas`.
- `% Reparto` vacío o 0 significa el reparto acordado por defecto, **nunca 0 %**. Para 0 % o 100 % se
  escribe la empresa correspondiente en lugar de `Ambas`.
- Solo valen los nombres de empresa acordados: cualquier otro es un error.

**Por qué:** cada empresa recibe su propio PV y EV. MS Project no distingue un 0 de un campo vacío, por
eso el 0 no puede significar 0 %.

## 7. Línea base

La línea base es el **plan aprobado por el cliente**. Solo se guarda cuando el cliente aprueba la
planificación; mientras no la apruebe, **no se define línea base**.

- Con la aprobación: *Proyecto > Establecer línea base*, para el proyecto completo.
- Tareas nuevas de un cambio **aprobado por el cliente**: línea base de las tareas seleccionadas,
  resumida en las tareas resumen (*Resumir líneas base: para todas las tareas de resumen*).
- Retrabajo o trabajo no previsto, sin cambio aprobado: la tarea se agrega **sin línea base**.
- La línea base no se modifica al registrar el avance ni al reprogramar.
- Cada cambio de línea base se justifica en las **Notas** de las tareas afectadas.

**Cómo afecta al informe EVM:**

| Situación | PV | EV |
|---|---|---|
| Tarea con línea base | Según las fechas de su línea base | Según su avance |
| Tarea agregada sin línea base (retrabajo, trabajo no previsto) | No cambia: la tarea no forma parte del plan aprobado | **Cuenta**: su trabajo se suma al del paquete, que avanza menos hasta que la tarea termina |
| Tarea nueva con línea base de un cambio aprobado | Entra en el PV en las fechas de su línea base | Según su avance |
| Proyecto sin línea base (sin aprobación del cliente) | **Provisional**: sale de las fechas actuales y cambia con cada cambio del cronograma; el informe EVM lo indica | Según su avance |
| Línea base modificada sin aprobación | El PV deja de ser el plan aprobado | — |

Ejemplo de retrabajo: un paquete terminado (EV = 100 %) recibe una tarea sin línea base para corregir un
rechazo del cliente. Su PV sigue igual; su EV baja (queda trabajo por hacer) y vuelve al 100 % cuando la
corrección termina. En un paquete `Ponderado`, el peso de la tarea nueva cuenta en el EV; en el PV, su
parte se reparte entre sus hermanas con línea base.

**Por qué:** el PV sale de la línea base. Si se mueve sin aprobación, el plan «persigue» a la realidad y
los desvíos desaparecen del informe EVM. El trabajo no previsto, en cambio, tiene que verse: por eso
cuenta en el EV y no en el PV.

## 8. Duración, hitos y lógica

- Una tarea de **duración 0 es un hito** y se marca como hito.
- La duración de cada tarea representa su trabajo: a igual peso vacío, una tarea más larga vale más.
- Cada tarea tiene predecesora y sucesora (salvo el inicio y el fin). Se evitan las delimitaciones
  (*No comenzar antes del*...) y los adelantos; para una fecha comprometida se usa una **fecha límite**.
- Tareas de detalle de duración acotada (en general, no más de dos meses).

**Por qué:** con la lógica completa, MS Project calcula una fecha de fin y una holgura fiables, que el
informe EVM usa para saber si un atraso afecta al plazo.

## 9. Registro mensual del avance

1. **Fecha de estado** = último día del mes informado (*Proyecto > Fecha de estado*).
2. Registrar el avance **a esa fecha**: % completado, comienzo y fin reales, hitos cumplidos.
3. Reprogramar el trabajo pendiente para que empiece **después** de la fecha de estado.
4. No tocar la línea base.
5. Anotar en las **Notas** las causas de los retrasos o adelantos importantes.

**Por qué:** el EV del mes es el avance a la fecha de estado. Avance registrado en otra fecha, o trabajo
pendiente que queda en el pasado, distorsionan el mes y la previsión de fin.

## 10. Notas de la tarea

Las **Notas** de cada tarea (doble clic en la tarea > pestaña *Notas*) son el lugar para justificar los
cambios y dar contexto. Se anota, con la fecha y quién lo decide:

- por qué se descartó o se desactivó una tarea;
- por qué una tarea tiene `Peso` = 0;
- qué cambio aprobado justifica una línea base nueva, o qué rechazo origina un retrabajo;
- por qué se movieron fechas comprometidas, se cambió el método de medición o se agregó una delimitación;
- cualquier supuesto o dependencia externa que explique la planificación.

**Por qué:** el informe EVM muestra *qué* cambió; las Notas explican *por qué*. Sin ellas, al mes
siguiente nadie recuerda el motivo, y cada desvío se tiene que volver a investigar.

## 11. Tareas que no deben contar en el informe EVM

| Situación | Qué hacer |
|---|---|
| La tarea se hará, pero no debe contar en el avance informado | `Peso` = 0, con el motivo en las Notas |
| La tarea se descartó, pero se quiere conservar | Desactivarla (*Tarea > Desactivar*), con el motivo en las Notas |
| La tarea no debe tener fila propia, pero su trabajo cuenta | `Visible EVM` vacío o No |

**Por qué:** desactivar una tarea también la saca del cálculo de fechas de MS Project: sus vínculos dejan
de mover a las sucesoras. Durante la planificación conviene el `Peso` = 0, que no altera el cronograma.

## 12. Lista de control

**Al planificar:**

- [ ] Todas las tareas del alcance tienen `Item EDT`, coherente con su tarea superior, sin saltos ni
      duplicados.
- [ ] Las tareas con `Visible EVM` = Sí van primero dentro de su tarea superior.
- [ ] Los pesos de cada grupo de hermanas suman 100 o menos; 0 solo donde la tarea no debe contar.
- [ ] Cada paquete tiene su método de medición acordado.
- [ ] `Empresa` y `% Reparto` completos (si el proyecto es entre dos empresas).
- [ ] Hitos marcados como hitos; lógica completa; sin delimitaciones innecesarias.
- [ ] Línea base guardada y resumida en las tareas resumen, **solo si el cliente aprobó la
      planificación**.
- [ ] Los cambios y las decisiones, justificados en las Notas.

**Cada mes:**

- [ ] Fecha de estado = último día del mes.
- [ ] Avance registrado a esa fecha y trabajo pendiente reprogramado.
- [ ] Línea base sin cambios (salvo un cambio aprobado por el cliente, anotado en las Notas).
