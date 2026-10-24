# Análisis de desviaciones

[Inicio](../README.md) › [Documentación](README.md) › Análisis de desviaciones

La hoja **Análisis de desviaciones** lista cada mes los paquetes de trabajo fuera de umbral (SPI o CPI
en amarillo o rojo, por debajo o por encima del rango verde) y tiene columnas para explicar cada desviación: causa, impacto, acción correctiva,
responsable y fecha compromiso. Así, el análisis que acompaña al informe mensual queda en el mismo libro
que los datos, y lo escrito se conserva al actualizar el libro y al regenerarlo por un cambio de línea
base.

![Hoja Análisis de desviaciones del ejemplo](img/excel/analisis-desviaciones.png)

*Análisis de diciembre de 2027 del [proyecto de ejemplo](ejemplo-paso-a-paso.md), con el texto de tres
paquetes como ejemplo: uno por encima del rango verde (CPI 1,10) y dos por debajo.*

---

## Contenido

1. [Qué paquetes se listan](#1-qué-paquetes-se-listan)
2. [Columnas](#2-columnas)
3. [Cierre mensual](#3-cierre-mensual)
4. [Qué se conserva de un mes a otro](#4-qué-se-conserva-de-un-mes-a-otro)
5. [Cambios de línea base](#5-cambios-de-línea-base)
6. [Filtrar e imprimir](#6-filtrar-e-imprimir)
7. [Preguntas frecuentes](#7-preguntas-frecuentes)

---

## 1. Qué paquetes se listan

`actualizar` agrega a la hoja las **filas de ingreso** (paquetes de trabajo y filas `-D Otras tareas`)
cuyo **SPI o CPI acumulado** al mes de la fecha de corte está fuera del **rango verde** del semáforo:
menor que `amarillo_inferior` o mayor que `amarillo_superior` (0,95 y 1,05 por defecto,
[Configuración](configuracion.md)). Es decir, los paquetes con semáforo amarillo o rojo en
`Indicadores EDT`.

- Un índice **por debajo** del rango indica atraso (SPI) o sobrecosto (CPI).
- Un índice **por encima** del rango también se analiza: puede indicar un AC incompleto, un avance
  informado de más o un presupuesto o una línea base poco exigentes
  ([Cuadro de Mando](cuadro-de-mando.md#3-valores-desviaciones-índices-y-pronósticos)).

- El SPI y el CPI se calculan igual que en `Indicadores EDT`. Un índice sin valor (por ejemplo, el CPI
  de un paquete sin AC) no se considera.
- Los umbrales son los mismos con los que se **generó** el libro: los que muestra la fila
  *Semáforo CPI y SPI* del `Cuadro de Mando`.
- Los elementos de nivel 1 y 2 (subtotales) no se listan: su desviación se explica con la de sus
  paquetes.
- Con [varios archivos de MS Project](varios-archivos.md), se listan los paquetes de todos los
  archivos.

La pantalla también muestra la lista al terminar `actualizar`, incluso con `--simular` (ejemplo a junio
de 2027, el de [Comandos](comandos.md#actualizar-carga-el--ganado-del-mes)):

```text
Paquetes fuera de umbral (SPI o CPI fuera de 0,95 a 1,05): 13 (hoja 'Análisis de desviaciones')
  1.1.1      Oficina de gestión del programa (PMO)     SPI  1,00  CPI  1,11
  1.1.2      Control de proyectos y EVM                SPI  1,00  CPI  1,11
  1.1.3      Gestión de riesgos                        SPI  1,00  CPI  1,09
  1.2.1      Licitación de ingeniería                  SPI  0,00  CPI  0,00
  1.2.2      Licitación de obras                       SPI  0,50  CPI  0,54
  ...
```

---

## 2. Columnas

| Columna | Contenido |
|---|---|
| Mes | Mes analizado (el de la fecha de corte). |
| Código EDT, Actividad | Paquete de trabajo. |
| BAC, PV, EV y AC acumulados, SV, CV, SPI, CPI | **Fórmulas** sobre las hojas `PV`, `EV` y `AC`, acumuladas hasta el mes de la fila. Si se corrige un valor en esas hojas, el análisis se recalcula. |
| Semáforo | `Rojo`, `Amarillo` o `Verde`, según el índice más alejado del rango verde: `Rojo` si alguno de los dos está por debajo de `rojo_inferior` o por encima de `rojo_superior`. Una fila puede quedar en `Verde` si se corrigieron los datos después de agregarla. |
| Antecedente | `Nuevo` (primer mes fuera de umbral), `También en MM/AAAA` (estaba fuera de umbral el mes anterior, sin texto) o `Texto copiado de MM/AAAA: se debe revisar` (ver la [sección 4](#4-qué-se-conserva-de-un-mes-a-otro)). |
| **Causa**, **Impacto en plazo y costo**, **Acción correctiva**, **Responsable**, **Fecha compromiso** | Celdas amarillas: se completan a mano. |

La columna oculta `clave` asocia cada fila con su paquete (el *Id. único* de la tarea en MS Project). No
se debe modificar.

---

## 3. Cierre mensual

```mermaid
flowchart LR
    A["actualizar<br/>(% ganado del mes)"] --> B{"¿AC del mes cargado<br/>antes de actualizar?"}
    B -- "Sí" --> D["Análisis de desviaciones<br/>con los paquetes del mes"]
    B -- "No" --> C["Cargar el AC en la hoja AC<br/>y ejecutar analizar"]
    C --> D
    D --> E["Completar causa, impacto,<br/>acción, responsable y fecha"]
    E --> F["Informe mensual"]
```

1. `actualizar` agrega los paquetes fuera de umbral del mes de corte.
2. Si el AC del mes se carga en la hoja `AC` **después** de `actualizar`, al ejecutarlo todavía no se
   conocían los paquetes con CPI bajo. Se ejecuta entonces `analizar`, que vuelve a calcular la lista con
   los valores del libro, **sin leer MS Project**:

   ```powershell
   python evm_export.py analizar --libro EVM_Programa.xlsx
   ```

   El libro debe estar cerrado en Excel. Como `actualizar`, guarda antes una copia de respaldo (salvo
   con `--sin-respaldo`).
3. Se completan las columnas amarillas de los paquetes del mes, con los responsables de cada paquete
   ([Proceso de reporte](proceso-reporte-evm.md#43-pasos-detallados)).

---

## 4. Qué se conserva de un mes a otro

- **Mismo mes.** Repetir `actualizar` (o `analizar`) para el mismo mes conserva el texto de cada
  paquete. Se agregan los paquetes que pasaron a estar fuera de umbral y se quitan los que volvieron
  dentro del umbral **si no tienen texto**; los que tienen texto se conservan.
- **Mes siguiente.** Al actualizar un mes nuevo:
  - las filas de los meses anteriores **con texto** se conservan, como registro de lo informado;
  - las filas de los meses anteriores **sin texto** se eliminan;
  - si un paquete sigue fuera de umbral y tenía texto, el texto se **copia** a la fila del mes nuevo,
    con el antecedente `Texto copiado de MM/AAAA: se debe revisar`, para actualizarlo en lugar de
    escribirlo de nuevo.
- **Filas agregadas a mano.** Una fila escrita a mano debajo de la lista (por ejemplo, el comentario de
  un elemento de nivel 2) se conserva, sin fórmulas.

No se deben insertar ni eliminar columnas. Las filas se pueden ordenar o filtrar con los botones del
encabezado.

---

## 5. Cambios de línea base

`generar --migrar-desde` traslada el análisis del libro anterior al libro nuevo
([El libro Excel](libro-excel.md#4-cambios-de-línea-base-o-de-edt---migrar-desde)):

- Cada fila pasa a la fila nueva de su paquete, y sus fórmulas usan los valores del libro nuevo. Como el
  PV corresponde a la nueva línea base, los índices de los meses ya informados pueden cambiar: el
  registro de lo informado con la línea base anterior es el libro anterior.
- Si el paquete ya no existe en el libro nuevo, el texto se conserva en una fila sin valores
  (antecedente `Fila inexistente en el libro nuevo`) y se advierte en `Validación`.

---

## 6. Filtrar e imprimir

- El encabezado tiene **filtros**: por ejemplo, se filtra la columna `Mes` para ver solo el mes del
  informe, o `Semáforo` para ver solo los rojos.
- La hoja se imprime en horizontal, con el ancho ajustado a una página y el encabezado repetido en cada
  página, para adjuntarla al informe mensual junto con el `Cuadro de Mando`.

---

## 7. Preguntas frecuentes

**¿Se pueden cambiar los umbrales?** Son los del semáforo (`[semaforo]`) con que se generó el libro. Para
usar otros se cambia la configuración y se regenera el libro con `--migrar-desde`, que conserva el
análisis. Para no analizar los índices altos se indican umbrales superiores que no se alcanzan en la
práctica ([Configuración](configuracion.md#ejemplos)).

**¿Por qué aparece un paquete de nivel de esfuerzo (LOE) con SPI 1?** Porque su CPI está fuera del rango
verde: el SPI de un paquete LOE siempre es 1, pero su costo puede desviarse. En el ejemplo, los paquetes
LOE de gestión (`1.1.1` a `1.1.3` y `1.2.3`) tienen CPI de 1,08 a 1,10 en diciembre de 2027 porque su
costo real es menor que el planificado.

**¿Qué pasa si borro una fila del análisis?** Si el paquete sigue fuera de umbral en el mes de corte,
vuelve a aparecer (sin texto) la próxima vez que se ejecute `actualizar` o `analizar` para ese mes.

**¿Se puede analizar un mes anterior?** Para ver sus valores basta con elegirlo como mes de reporte en el
`Cuadro de Mando`. Para volver a calcular su lista, `actualizar` con la fecha de corte de ese mes (y
`--forzar` si hay meses posteriores cargados) la recalcula y conserva su texto, pero también vuelve a
escribir el % ganado de ese mes con el avance que tiene el archivo de MS Project: se debe usar el
archivo (o el XML) guardado en ese cierre y, después, volver a ejecutar `actualizar` con la última fecha
de corte. Las filas sin texto de los demás meses se eliminan.
