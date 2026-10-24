# El libro Excel

[Inicio](../README.md) › [Documentación](README.md) › El libro Excel

`generar` crea un libro `.xlsx` con el % planificado de cada mes para cada elemento de la EDT, hojas
para ingresar el BAC y el AC, y hojas de indicadores y gráficos calculados con fórmulas de Excel.
`actualizar` carga cada mes el % ganado en ese mismo libro.

El BAC de cada fila se ingresa en la hoja `PV`, el PV y el EV son el BAC × el % de las hojas `% PV` y
`% EV`, y el AC se carga a mano en la hoja `AC`
([Cálculo del PV y el EV](calculo-pv-ev.md#11-hojas-del-libro)).

---

## Contenido

1. [Hojas del libro](#1-hojas-del-libro)
2. [Hojas PV, EV y AC](#2-hojas-pv-ev-y-ac)
3. [Reglas de uso](#3-reglas-de-uso)
4. [Cambios de línea base o de EDT (`--migrar-desde`)](#4-cambios-de-línea-base-o-de-edt---migrar-desde)
5. [Hojas de cálculo y control](#5-hojas-de-cálculo-y-control)
6. [Libros de versiones anteriores](#6-libros-de-versiones-anteriores)

---

## 1. Hojas del libro

| Hoja | Contenido |
|---|---|
| **Cuadro de Mando** | Tablero del mes de reporte: valores acumulados, desviaciones, índices con semáforo, pronósticos y tres gráficos. Ver [Cuadro de Mando e indicadores](cuadro-de-mando.md). |
| **Resumen** | Fecha de corte (celda amarilla `B5`, con nombre `FechaCorte`), parámetros de la exportación e indicadores del proyecto a la fecha de corte, incluido el Cronograma Ganado. |
| **PV** | Valor planificado: columnas A–F con la EDT y el BAC, y una columna por mes. El BAC se ingresa en la columna D (celdas amarillas) y cada mes es el BAC × el % planificado. |
| **% PV** | % planificado de cada mes, que calcula `generar`. |
| **EV** | Valor ganado de cada mes: el BAC × el % ganado. |
| **% EV** | % ganado de cada mes, que calcula `actualizar`. |
| **AC** | Costo real de cada mes (celdas amarillas), que se ingresa a mano. |
| **PV-EmpresaA**, **% PV-EmpresaA**, **EV-EmpresaA**, **% EV-EmpresaA**, **AC-EmpresaA** y las de EmpresaB | Solo en multi-empresa: valores de cada empresa, en lugar de las hojas `% PV` y `% EV`. Las hojas `PV`, `EV` y `AC` son la suma de las dos ([Multi-empresa](multiempresa.md)). |
| **Evolución mensual** | Tabla mensual del proyecto (datos de la curva S). |
| **Indicadores EDT** | Indicadores de cada elemento de la EDT a la fecha de corte. |
| **Análisis de desviaciones** | Paquetes fuera de umbral de cada mes, con causa, impacto, acción correctiva, responsable y fecha ([Análisis de desviaciones](analisis-desviaciones.md)). |
| **Validación** | Advertencias de la exportación ([Validaciones](validaciones-y-problemas.md#1-validaciones)). |
| **Listas** | Métodos de medición del desplegable de la columna C y su regla de cálculo. |
| **Depuración costes** y **PV-depuración** | Solo con `depuracion_costes = true`: gráfico comparativo y previsión pegada del plan de costes de una empresa; la hoja PV de esa empresa lleva una fila «(coste)» debajo de cada fila, y en las demás hojas esas filas quedan vacías y ocultas ([Depuración de costes](depuracion-costes.md)). |
| `_graficos`, `_estructura`, `_meta` (y `_depuracion`) | Hojas ocultas con los datos de los gráficos, los datos que usan `actualizar` y `--migrar-desde` y, con depuración de costes, la asignación de la previsión pegada. **No se deben modificar.** |

---

## 2. Hojas PV, EV y AC

Las tres hojas tienen la misma estructura:

| Columna | Contenido |
|---|---|
| A | Código EDT (`1.0.0`, `1.1.0`, `1.2.3`...). |
| B | Actividad (nombre de la tarea). |
| C | Método de medición (lista desplegable; solo se edita en la hoja `PV`). |
| D | BAC: presupuesto de la fila. Se ingresa a mano en la hoja `PV` (en multi-empresa, en `PV-EmpresaA` y `PV-EmpresaB`), en las celdas amarillas; las demás hojas lo toman de allí. |
| E | Fecha inicio Plan (PV): inicio de la línea base. |
| F | Fecha fin Plan (PV): fin de la línea base. |
| G ... DV | Un mes por columna, del mes 1 al último mes del horizonte (`meses`; con 120, el valor por defecto, hasta la columna DV). Fila 3: número de mes; fila 4: fecha del mes. |
| Al final | Columna **Total horizonte** (en la hoja `PV`, en rojo si no coincide con el BAC) y fila **TOTAL**. |

![Hoja PV del proyecto de ejemplo](img/excel/pv.png)

*Hoja `PV`: las filas de nivel 1 y 2 (fondo azul) son subtotales; las de nivel 3 son paquetes de
trabajo, con el BAC ingresado en la columna D.*

Las **celdas amarillas** son de ingreso: el BAC de la hoja `PV` y el AC del mes de la hoja `AC` (en
multi-empresa, los de las hojas de cada empresa). Las hojas `% PV` y `% EV` las calcula el programa desde
MS Project.

Los valores mensuales son el **valor del mes**, no el acumulado.

![Hoja AC del proyecto de ejemplo](img/excel/ac.png)

*Hoja `AC` con los meses 13 a 24 del ejemplo (los meses 1 a 12 están ocultos en la imagen).*

---

## 3. Reglas de uso

- Las **filas con subelementos** (las de nivel 1 y 2 y, si el campo `Visible` agrega detalle, también
  otras) son **subtotales** (fórmulas `SUM` de sus hijos). Los valores se ingresan solo en las filas sin
  subelementos: paquetes de trabajo y filas `-D`.
- Las filas se pueden **contraer por nivel** con los botones de esquema (1, 2, 3) a la izquierda.
- **No se deben insertar, eliminar ni ordenar filas o columnas** en `PV`, `% PV`, `EV`, `% EV`, `AC` ni
  en las hojas de cada empresa: `actualizar` ubica cada fila por su posición. Si los códigos EDT de la columna A no
  coinciden con los del libro generado, o si falta una hoja o se modificaron las hojas ocultas,
  `actualizar` se detiene sin modificar el libro. Para cambiar la EDT se modifica MS Project y se
  vuelve a generar el libro con `--migrar-desde`.
- No se debe escribir sobre las celdas con fórmulas (los meses de las hojas `PV` y `EV`, los subtotales,
  la fila TOTAL y, en multi-empresa, la suma de las empresas): `actualizar` informa las que perdieron su
  fórmula para restaurarlas.
- El **método de medición** se puede cambiar en la columna C de la hoja `PV`. Se pueden agregar métodos
  al final de la columna A de la hoja `Listas`, sin dejar filas vacías.
- Los gráficos no requieren ninguna acción: al ingresar el BAC o el AC, al cargar el % ganado, al
  cambiar la fecha de corte en `Resumen!B5` o el mes de reporte del `Cuadro de Mando`, Excel recalcula
  tablas y gráficos.
- `actualizar` vuelve a crear los gráficos del `Cuadro de Mando` (la biblioteca que escribe el libro no
  conserva gráficos). Por eso no se deben agregar gráficos ni imágenes propias en este libro; los
  gráficos adicionales se deben hacer en otro libro que lea estos datos.
- Las fórmulas se calculan al abrir el libro en Excel. En la vista protegida o en vistas previas de
  archivos, los indicadores pueden aparecer vacíos hasta habilitar la edición.
- El PV acumulado a la fecha de corte incluye **el mes completo** de esa fecha. Como fecha de corte se
  debe usar el **último día del mes**.
- `actualizar` guarda una **copia de respaldo** antes de modificar el libro
  (`<libro>.respaldo-AAAAMMDD-HHMMSS.xlsx`) y lo guarda primero en un archivo temporal, de modo que un
  error al guardar no daña el libro anterior. Si el libro está abierto en Excel, avisa antes de leer
  MS Project. Se recomienda conservar, además, una copia del libro de cada informe emitido.

---

## 4. Cambios de línea base o de EDT (`--migrar-desde`)

Cuando se aprueba una nueva línea base o cambia la EDT, se genera un libro nuevo y se traslada el
historial del libro anterior:

```powershell
python evm_export.py generar --origen "C:\Programa\Programa.mpp" --config config.toml `
    --salida EVM_Programa_LB2.xlsx --migrar-desde EVM_Programa.xlsx
```

```mermaid
flowchart LR
    A["Libro anterior<br/>BAC, % ganado y AC,<br/>métodos editados,<br/>análisis, fecha de corte"] --> M{"--migrar-desde"}
    P["MS Project<br/>nueva línea base / EDT"] --> G["generar"]
    M --> G
    G --> N["Libro nuevo<br/>% planificado nuevo +<br/>BAC, % ganado y AC"]
    G --> V["Validación:<br/>filas que no se pudieron trasladar"]
```

Qué se traslada:

- El **BAC** ingresado, el **% ganado** de cada mes y el **AC** de cada fila, asociados por el
  **UniqueID** de la tarea en MS Project (o por código EDT si el proyecto cambió de nombre), sin
  duplicarlos.
- Si una tarea quedó dentro de un paquete nuevo (por ejemplo, al bajar `nivel_maximo` o al poner `No` en
  el campo `Visible` de todas las subtareas de una tarea), su historial se suma a ese paquete: el BAC y
  el AC se suman, y el % ganado se combina ponderado por el BAC de cada fila anterior.
- El **método** de la columna C se conserva solo si fue **editado en Excel**. Si el método cambió en
  MS Project, se usa el nuevo.
- El texto de la hoja `Análisis de desviaciones`
  ([Análisis de desviaciones](analisis-desviaciones.md#5-cambios-de-línea-base)).
- La **fecha de corte** de la hoja `Resumen`.
- Con depuración de costes, el contenido de la hoja `PV-depuración`, si el libro nuevo también depura la
  misma empresa ([Depuración de costes](depuracion-costes.md#8-ciclo-de-ajuste)).

Lo que no se puede trasladar se informa como **Error** en `Validación` para cargarlo a mano: filas que
se dividieron en varias y tareas eliminadas. Tampoco se admite migrar entre un libro de una sola empresa
y uno multi-empresa, ni desde un libro de versiones anteriores
([sección 6](#6-libros-de-versiones-anteriores)).

Con varios archivos de MS Project, cada fila se asocia por el nombre del archivo y el UniqueID de la
tarea; al agregar o quitar archivos, o al pasar de un solo archivo a varios, se aplica lo indicado en
[Varios archivos](varios-archivos.md#7-cambios-de-línea-base-y-de-archivos).

> El % planificado del libro nuevo corresponde a la nueva línea base: la curva PV de los meses ya
> informados cambia (el BAC trasladado se mantiene). Se recomienda conservar el libro anterior como
> registro de lo informado con la línea base previa y documentar la fecha y el motivo del cambio
> ([Proceso de reporte](proceso-reporte-evm.md#5-cambios-de-línea-base)).

---

## 5. Hojas de cálculo y control

### Evolución mensual

Tabla calculada con fórmulas a partir de la fila TOTAL de las hojas `PV`, `EV` y `AC`: PV, EV y AC del
mes y acumulados, % de avance, SPI y CPI acumulados y BAC. Un mes tiene EV cargado cuando `actualizar`
escribió su % ganado. EV y AC acumulados se muestran desde el primer mes con datos hasta la fecha de
corte; fuera de ese tramo aparecen como `#N/A` para que los gráficos no dibujen esos meses.

![Hoja Evolución mensual](img/excel/evolucion-mensual.png)

### Indicadores EDT

Para cada fila de la EDT, a la fecha de corte: BAC, PV, EV y AC acumulados, SV, CV, SPI, CPI, EAC, VAC
y porcentajes planificado, completado y gastado, con semáforo en SPI y CPI. Permite identificar qué
paquetes explican la desviación del proyecto. El EAC de cada fila es BAC / CPI; si la fila no tiene CPI
(sin AC cargado), es AC + (BAC − EV). No depende del método EAC elegido en el `Cuadro de Mando`.

![Hoja Indicadores EDT](img/excel/indicadores-edt.png)

### Análisis de desviaciones

Paquetes de trabajo con SPI o CPI en amarillo o rojo (por debajo o por encima del rango verde) en el
mes de corte, que `actualizar` (o
`analizar`) agrega cada mes, con columnas amarillas para explicar cada desviación. El texto se conserva
al actualizar y con `--migrar-desde`. Ver [Análisis de desviaciones](analisis-desviaciones.md).

### Validación

Advertencias y observaciones de la exportación, con su severidad (`Info`, `Advertencia`, `Error`), el
código EDT, la tarea afectada y el tema (columna E, para filtrar). Ver
[Validaciones](validaciones-y-problemas.md#1-validaciones). El comando `validar` muestra en pantalla las
mismas observaciones, y además las del avance y la lógica del cronograma, sin crear el libro.

![Hoja Validación](img/excel/validacion.png)

### Listas

Métodos de medición del desplegable de la columna C y su regla de cálculo.

![Hoja Listas](img/excel/listas.png)

---

## 6. Libros de versiones anteriores

Los libros generados con versiones anteriores del programa, que calculaban el PV y el EV con los costos
de MS Project, no se pueden usar: `actualizar`, `analizar` y `--migrar-desde` se detienen sin
modificarlos. Se debe generar un libro nuevo, ingresar el BAC y cargar a mano el historial que se
necesite: el % ganado de cada mes en la hoja `% EV` y el AC en la hoja `AC`
([Cálculo del PV y el EV](calculo-pv-ev.md#7-del--acumulado-al--del-mes)).
