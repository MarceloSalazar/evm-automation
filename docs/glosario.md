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
| **BEI** | Índice de ejecución de la línea base | Baseline Execution Index | — | Tareas terminadas / tareas que debían terminar según la línea base (comando `validar`) |
| **CPLI** | Índice de longitud del camino crítico | Critical Path Length Index | — | (Días laborables hasta el fin previsto + holgura frente al fin de la línea base) / días hasta el fin previsto (comando `validar`) |

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
| Texto1..30, Número1..20 | Text1..30, Number1..20 | Campos personalizados (método, peso de los hitos, empresa, % de reparto, visible y descripción de las fichas). |
| Hito | Milestone | Su `Peso hito` es un % del BAC del paquete; método `Hitos ponderados`. |
| Activo / Desactivar | Active / Inactivate | Tareas inactivas excluidas. |
| Duración (estimada, con `?`) | Duration (Estimated) | `validar`: hitos con duración, tareas con duración 0 y duraciones estimadas. |
| Programada manualmente | Manually Scheduled | `validar`: MS Project no recalcula sus fechas. |
| Notas | Notes | `generar-fichas`: descripción de cada ficha (`[campos] descripcion = "Notas"`). |
| Predecesoras, sucesoras, vínculos FC, CC, FF y CF | Predecessors, successors, FS, SS, FF and SF links | `validar`: lógica del cronograma. `generar-fichas`: predecesoras y sucesoras de cada ficha. |
| Posiciones (positivas o negativas, adelantos) | Lag (lead) | `validar`: esperas y adelantos en los vínculos. |
| Delimitación (*Debe comenzar el*, *No finalizar después del*...) | Constraint (Must Start On, Finish No Later Than...) | `validar`: delimitaciones que fijan las fechas. |
| Comienzo, Fin | Start, Finish | Cronograma vigente: el libro llama **fin previsto** al *Fin* de cada paquete. No se debe confundir con el campo *Fin previsto* de MS Project en español, que es el fin de la línea base. |
| Holgura total | Total Slack | Tiempo que una tarea puede atrasarse sin atrasar el fin del proyecto (columna `Holgura` de `Indicadores EDT`, en días laborables). |
| Crítica, camino crítico | Critical, critical path | Tareas sin holgura: su atraso atrasa el fin del proyecto (columna `Crítica`). |
| Variación de fin | Finish Variance | Equivale al desvío de fin del libro (que se cuenta con el calendario del programa). |

---

## Términos del programa

| Término | Significado |
|---|---|
| BAC de una fila | Presupuesto del paquete, que se ingresa a mano en la columna D de la hoja `PV` (o de `PV-EmpresaA` y `PV-EmpresaB`). |
| % planificado | Parte del BAC de una fila planificada en cada mes (hoja `% PV`); suma 100 % en cada fila. Lo calcula `generar` con las fechas de línea base de las tareas y el `Peso hito` de los hitos ([Cálculo del PV y el EV](calculo-pv-ev.md#4--planificado-peso-de-cada-tarea)). |
| % ganado | Parte del BAC de una fila ganada en cada mes (hoja `% EV`), según su método de medición. La escribe `actualizar` ([Cálculo del PV y el EV](calculo-pv-ev.md#6--ganado-métodos-de-medición)). |
| Peso de una tarea | Fracción del BAC de la fila que le corresponde, por su duración de línea base o, en un hito, por su `Peso hito`. |
| Fecha de corte | Último día del mes informado. Por defecto, la fecha de estado de MS Project. |
| Fila de ingreso | Fila sin filas inferiores (por defecto, de nivel 3) donde se ingresan el BAC y el AC, y cuyo % ganado escribe `actualizar`: paquetes de trabajo y filas `-D`. Las filas con filas inferiores son subtotales. |
| Fila `-D Otras tareas` | Fila `<código>-D Otras tareas de «tarea»` que reúne las tareas sin fila propia (sin código EDT, con código más detallado que `nivel_maximo` o, con `lista_edt`, fuera de la lista) que dependen directamente de una fila resumen, con su propio BAC. |
| Planificación | Rama del proyecto bajo una tarea de primer nivel con código EDT (`INTERNO`, `CLIENTE`) que se exporta a un libro. |
| Horizonte | Meses del libro (120 por defecto). Por defecto, mes 1 = primer mes con % planificado (en general, el de inicio de la línea base); se cambia con `mes_inicio` o `--mes-inicio`. |
| Mes de reporte | Mes que muestra el `Cuadro de Mando` (celda `E5`). |
| EIA-748 | Norma de EE. UU. para los sistemas de gestión del valor ganado (revisión E, 2026: 27 guías). Evalúa el sistema de gestión de la organización; qué partes cubre el programa está en [Norma EIA-748-E](evaluacion-eia-748-e.md). |
| DCMA de 14 puntos | Evaluación de calidad de cronogramas de la *Defense Contract Management Agency* de EE. UU.: lógica, adelantos, posiciones, tipos de vínculo, delimitaciones, holgura, duración, fechas no válidas, recursos, tareas incumplidas, camino crítico, CPLI y BEI. `validar` revisa 12 de los 14 puntos ([Evaluación DCMA de 14 puntos](evaluacion-dcma.md)). |
| Valor del mes | Valor de un mes, no acumulado: el % ganado que `actualizar` escribe en la hoja `% EV` (el EV del mes es el BAC × ese %) o el AC que se ingresa en la hoja `AC`. |
| Desvío de fin | Días laborables entre el fin de la línea base y el fin previsto de una fila o del proyecto, según el cronograma de MS Project (> 0 = atraso) ([Indicadores EDT](libro-excel.md#indicadores-edt)). |
| Mes anterior con datos | Último mes con EV cargado antes del mes de corte o de reporte: con informes bimestrales, el del informe anterior. Base de las variaciones de la tendencia. |
| Últimos 3 meses | CPI y SPI calculados con el EV, el AC y el PV de los 3 meses que terminan en el mes de corte o de reporte: el desempeño reciente ([Cuadro de Mando](cuadro-de-mando.md#4-tendencia-y-plazo)). |
| Lista EDT | Archivo de texto con los códigos EDT que tienen fila en el libro (`[exportacion] lista_edt`); reemplaza a `nivel_maximo` y al campo `Visible` ([Cálculo del PV y el EV](calculo-pv-ev.md#34-lista-edt-filas-definidas-en-un-archivo)). |
