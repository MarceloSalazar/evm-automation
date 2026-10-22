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

---

## Términos de MS Project usados en la documentación

| Español (MS Project) | Inglés | Uso en el programa |
|---|---|---|
| Línea base, Línea base 1..10 | Baseline, Baseline1..10 | Define el PV: sus fechas (modo porcentaje) o su costo (modo costo). |
| Costo previsto | Baseline Cost | BAC de cada tarea (modo costo). |
| Comienzo previsto, Fin previsto | Baseline Start, Baseline Finish | Fechas del plan (columnas E y F de la hoja `PV`). |
| Plan provisional | Interim plan | Copia de líneas base para el historial. |
| Fecha de estado | Status Date | Fecha de corte. |
| % completado | % Complete | Métodos `% Completado`, `0/100`, `X/Y`, hitos. |
| % físico completado | Physical % Complete | Métodos `% Físico completado` y `Unidades completadas`. |
| Método de valor acumulado | Earned Value Method | Cómo calcula MS Project el CPTR. |
| Comienzo real, Fin real | Actual Start, Actual Finish | Comienzo y fin de los paquetes. |
| Costo real | Actual Cost | AC (si no hay CRTR). |
| Costo fijo, Acumulación de costos fijos | Fixed Cost, Fixed Cost Accrual | Presupuesto por tarea y su distribución. |
| Tasa estándar, Costo/Uso, Acumular | Standard Rate, Cost/Use, Accrue At | Costo de los recursos y su distribución. |
| Recurso de trabajo, material, de costo | Work, Material, Cost resource | Formas de cargar el presupuesto. |
| EDT, Número de esquema | WBS, Outline Number | Código EDT de cada tarea. |
| Texto1..30, Número1..20 | Text1..30, Number1..20 | Campos personalizados (método, peso, empresa). |
| Hito | Milestone | Método `Hitos ponderados`; en el modo porcentaje, su `Peso hito` es un % del BAC del paquete. |
| Activo / Desactivar | Active / Inactivate | Tareas inactivas excluidas. |
| Uso de tareas | Task Usage | Vista para verificar el PV mensual. |

---

## Términos del programa

| Término | Significado |
|---|---|
| Modo porcentaje | Modo por defecto: MS Project no tiene importes; el PV y el EV son el BAC ingresado en el libro × el % planificado y el % ganado ([Modo porcentaje](modo-porcentaje.md)). |
| Modo costo | El PV y el BAC salen del costo de línea base de MS Project (`modo = "costo"`). |
| % planificado | Parte del BAC de una fila planificada en cada mes (hoja `% PV`); suma 100 % en cada fila. |
| % ganado | Parte del BAC de una fila ganada en cada mes (hoja `% EV`), según su método de medición. |
| Peso de una tarea | Modo porcentaje: fracción del BAC de la fila que le corresponde, por su duración de línea base o, en un hito, por su `Peso hito`. |
| Fecha de corte | Último día del mes informado. Por defecto, la fecha de estado de MS Project. |
| Fila de ingreso | Fila de nivel 3 (o sin hijos) donde se cargan el BAC (modo porcentaje), el EV y el AC. Las de nivel 1 y 2 son subtotales. |
| Fila `-D Otros costos` / `-D Otras tareas` | Fila que recoge el costo propio de una tarea resumen o las tareas sin código EDT bajo un elemento de nivel 1 o 2 (en el modo porcentaje, con su propio BAC). |
| Planificación | Rama del proyecto bajo una tarea de primer nivel con código EDT (`INTERNO`, `CLIENTE`) que se exporta a un libro. |
| Horizonte | Meses del libro (120 por defecto). Mes 1 = mes de inicio de la línea base. |
| Mes de reporte | Mes que muestra el `Cuadro de Mando` (celda `E5`). |
| Valor del mes | EV o AC de un mes (no acumulado), que se escribe en las hojas `EV` y `AC` (en el modo porcentaje, el % ganado del mes en la hoja `% EV`). |
