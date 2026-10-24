# Glosario

[Inicio](../README.md) › [Documentación](README.md) › Glosario

La documentación y el libro usan las siglas en inglés de la gestión del valor ganado (PV, EV, AC...),
que son las del estándar del PMI y las más difundidas. MS Project en español usa otros nombres: llama
**valor acumulado** al valor ganado y usa siglas propias (CPTP, CPTR, CRTR...). Esta tabla relaciona
ambos.

---

## Indicadores EVM

| Sigla | Nombre en español | Nombre en inglés | Campo de MS Project en español | Fórmula |
|---|---|---|---|---|
| **BAC** | Presupuesto a la conclusión | Budget at Completion | *Costo previsto* (descrito también como CPF, costo presupuestado al finalizar) | Σ PV |
| **PV** | Valor planificado | Planned Value (BCWS) | **CPTP** (costo presupuestado del trabajo programado) | — |
| **EV** | Valor ganado | Earned Value (BCWP) | **CPTR** (costo presupuestado del trabajo realizado), "valor acumulado" | según el método |
| **AC** | Costo real | Actual Cost (ACWP) | **CRTR** (costo real del trabajo realizado) | — |
| **SV** | Variación del cronograma | Schedule Variance | **VP** | EV − PV |
| **CV** | Variación del costo | Cost Variance | **VC** | EV − AC |
| **SV %**, **CV %** | Variación porcentual | SV %, CV % | **%VP**, **%VC** | SV / PV, CV / EV |
| **SPI** | Índice de desempeño del cronograma | Schedule Performance Index | **IRP** | EV / PV |
| **CPI** | Índice de desempeño del costo | Cost Performance Index | **IRC** | EV / AC |
| **EAC** | Estimación a la conclusión | Estimate at Completion | **CEF** | BAC / CPI (u otras) |
| **ETC** | Estimación hasta la conclusión | Estimate to Complete | — | EAC − AC |
| **VAC** | Variación a la conclusión | Variance at Completion | **VAF** | BAC − EAC |
| **TCPI** | Índice de desempeño del trabajo por completar | To-Complete Performance Index | **IRPC** | (BAC − EV) / (BAC − AC) |
| **ES** | Cronograma ganado | Earned Schedule | — | Momento en que PV acumulado = EV |
| **AT** | Tiempo real transcurrido | Actual Time | — | Meses hasta la fecha de corte |
| **SV(t)** | Variación del cronograma en tiempo | Schedule Variance (time) | — | ES − AT |
| **SPI(t)** | Índice de desempeño del cronograma en tiempo | SPI (time) | — | ES / AT |

> **Atención:** en MS Project en español, **VP** es la variación del cronograma (SV), no el valor
> planificado. El valor planificado es el **CPTP**.

El programa no lee estos campos de MS Project: el cronograma no necesita costos, el PV y el EV son el
BAC que se ingresa en el libro × el % planificado y el % ganado, y el AC se carga a mano
([Cálculo del PV y el EV](calculo-pv-ev.md)).

---

## Términos de MS Project usados en la documentación

| Español (MS Project) | Inglés | Uso en el programa |
|---|---|---|
| Línea base, Línea base 1..10 | Baseline, Baseline1..10 | Sus fechas definen el % planificado. |
| Comienzo previsto, Fin previsto | Baseline Start, Baseline Finish | Fechas del plan (columnas E y F de la hoja `PV`). |
| Plan provisional | Interim plan | Copia de líneas base para el historial. |
| Fecha de estado | Status Date | Fecha de corte. |
| % completado | % Complete | Métodos `% Completado`, `0/100`, `X/Y`, hitos. |
| % físico completado | Physical % Complete | Métodos `% Físico completado` y `Unidades completadas`. |
| Método de valor acumulado | Earned Value Method | Cómo calcula MS Project su valor acumulado (CPTR). El programa no lo usa. |
| Comienzo real, Fin real | Actual Start, Actual Finish | Comienzo y fin de los paquetes. |
| EDT, Número de esquema | WBS, Outline Number | Código EDT de cada tarea. |
| Texto1..30, Número1..20 | Text1..30, Number1..20 | Campos personalizados (método, peso de los hitos, empresa, visible). |
| Hito | Milestone | Su `Peso hito` es un % del BAC del paquete; método `Hitos ponderados`. |
| Activo / Desactivar | Active / Inactivate | Tareas inactivas excluidas. |

---

## Términos del programa

| Término | Significado |
|---|---|
| BAC de una fila | Presupuesto del paquete, que se ingresa a mano en la columna D de la hoja `PV` (o de `PV-EmpresaA` y `PV-EmpresaB`). |
| % planificado | Parte del BAC de una fila planificada en cada mes (hoja `% PV`); suma 100 % en cada fila. Lo calcula `generar` con las fechas de línea base de las tareas y el `Peso hito` de los hitos ([Cálculo del PV y el EV](calculo-pv-ev.md#4--planificado-peso-de-cada-tarea)). |
| % ganado | Parte del BAC de una fila ganada en cada mes (hoja `% EV`), según su método de medición. La escribe `actualizar` ([Cálculo del PV y el EV](calculo-pv-ev.md#6--ganado-métodos-de-medición)). |
| Peso de una tarea | Fracción del BAC de la fila que le corresponde, por su duración de línea base o, en un hito, por su `Peso hito`. |
| Fecha de corte | Último día del mes informado. Por defecto, la fecha de estado de MS Project. |
| Fila de ingreso | Fila de nivel 3 (o sin hijos) donde se ingresan el BAC y el AC, y cuyo % ganado escribe `actualizar`. Las de nivel 1 y 2 son subtotales. |
| Fila `-D Otras tareas` | Fila `<código>-D Otras tareas de «tarea»` que reúne las tareas sin código EDT que dependen directamente de un elemento de nivel 1 o 2, con su propio BAC. |
| Planificación | Rama del proyecto bajo una tarea de primer nivel con código EDT (`INTERNO`, `CLIENTE`) que se exporta a un libro. |
| Horizonte | Meses del libro (120 por defecto). Mes 1 = mes de inicio de la línea base. |
| Mes de reporte | Mes que muestra el `Cuadro de Mando` (celda `E5`). |
| Valor del mes | Valor de un mes, no acumulado: el % ganado que `actualizar` escribe en la hoja `% EV` (el EV del mes es el BAC × ese %) o el AC que se ingresa en la hoja `AC`. |
