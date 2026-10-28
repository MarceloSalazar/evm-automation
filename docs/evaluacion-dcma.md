# Evaluación DCMA de 14 puntos

[Inicio](../README.md) › [Documentación](README.md) › Evaluación DCMA de 14 puntos

La **evaluación DCMA de 14 puntos** es la referencia más usada para medir la calidad de un cronograma.
Este documento explica qué revisa cada punto, cómo lo calcula el comando
[`validar`](comandos.md#validar-revisa-el-archivo-de-ms-project) y qué hacer en MS Project cuando no se
cumple.

---

## Contenido

1. [Qué es y por qué importa](#1-qué-es-y-por-qué-importa)
2. [Cómo cuenta el programa](#2-cómo-cuenta-el-programa)
3. [Los 14 puntos y qué verifica el programa](#3-los-14-puntos-y-qué-verifica-el-programa)
4. [Detalle de cada punto](#4-detalle-de-cada-punto)
5. [Otros controles del cronograma](#5-otros-controles-del-cronograma)
6. [Lo que el programa no verifica](#6-lo-que-el-programa-no-verifica)
7. [Uso](#7-uso)
8. [Referencias](#8-referencias)

---

## 1. Qué es y por qué importa

En EE. UU., la norma **EIA-748** define los requisitos de un sistema de gestión del valor ganado (EVM).
Uno de ellos es un cronograma integrado: la secuencia del trabajo con sus dependencias, que es la base
de la línea base de medición (el PV). La *Defense Contract Management Agency* (DCMA), que revisa esos
sistemas en los contratos de defensa, publicó 14 controles automáticos sobre el cronograma. Se
convirtieron en la referencia habitual de las herramientas de análisis de cronogramas. Las métricas de
cumplimiento que la DCMA usa hoy en sus revisiones (DECM) incluyen controles equivalentes. Qué otros
requisitos de la norma cubre el programa está en [Norma EIA-748-E](evaluacion-eia-748-e.md).

Para este programa importa por dos motivos:

- El **% planificado** de cada fila sale de las fechas de la línea base. Un cronograma sin lógica, con
  delimitaciones o con tareas muy largas da un PV poco confiable.
- El **plazo** (desvío de fin, holgura y tareas críticas del [Cuadro de Mando](cuadro-de-mando.md#4-tendencia-y-plazo))
  es el que calcula MS Project. Solo es una previsión realista si el cronograma tiene la lógica completa
  y está actualizado a la fecha de estado.

Los umbrales de la DCMA son **señales de alerta**, no un aprobado o reprobado automático: un control
fuera de umbral indica dónde revisar el cronograma, y puede tener una explicación válida. Fuera de los
contratos del gobierno de EE. UU. no son obligatorios; se usan como buena práctica.

---

## 2. Cómo cuenta el programa

- **Tareas revisadas:** las tareas de detalle (sin subtareas) que exporta el libro: las de la
  planificación elegida, activas y sin `Visible = No`. Las tareas resumen no se cuentan.
- **Tareas sin terminar:** las que no tienen el 100 % completado. Casi todos los controles cuentan solo
  estas tareas: lo terminado ya no se puede corregir.
- **Días laborables:** la duración y la holgura se leen de MS Project y se pasan a días con las horas
  por día del proyecto (*Archivo > Opciones > Programación*; 8 horas por defecto). Las fechas (CPLI) se
  cuentan con el mismo calendario que reparte el % planificado: el del proyecto (XML) o de lunes a viernes
  con los `feriados` de la configuración (`.mpp`).
- **44 días laborables** (unos 2 meses) es el límite de la DCMA para una holgura o una duración alta.
- **Fecha de estado:** la del proyecto en MS Project (*Proyecto > Información del proyecto*) o la de
  `--fecha-corte`. Sin fecha de estado, no se calculan los puntos que dependen de ella (9, 11, 13 y 14).
- **Filas LOE:** las tareas de las filas con método *Avance lineal (LOE)* no cuentan para la holgura
  alta ni para la duración alta: un nivel de esfuerzo dura todo el proyecto por definición.
- **Varios archivos** (`[[proyectos]]`): cada archivo se evalúa por separado y sus observaciones llevan
  su nombre (`[P1]`).

---

## 3. Los 14 puntos y qué verifica el programa

| # | Punto | Qué mide | Criterio DCMA | `validar` | Código |
|---|---|---|---|---|---|
| 1 | Lógica | Tareas sin terminar sin predecesoras o sin sucesoras | Hasta 5 % | Sí | [LOG-002](codigos.md#log-002) a [LOG-005](codigos.md#log-005) ([LOG-001](codigos.md#log-001) sin vínculos) |
| 2 | Adelantos | Vínculos con posición negativa | Ninguno | Sí | [LOG-006](codigos.md#log-006) |
| 3 | Posiciones | Vínculos con posición positiva (esperas) | Hasta 5 % de los vínculos | Sí | [LOG-008](codigos.md#log-008) |
| 4 | Tipos de vínculo | Vínculos fin a comienzo (FC) | Al menos 90 % | Sí | [LOG-009](codigos.md#log-009) ([LOG-007](codigos.md#log-007) para los CF) |
| 5 | Delimitaciones rígidas | Tareas sin terminar con delimitaciones que fijan sus fechas | Hasta 5 % | Sí | [LOG-011](codigos.md#log-011) y [LOG-012](codigos.md#log-012) |
| 6 | Holgura alta | Tareas sin terminar con más de 44 días laborables de holgura total | Hasta 5 % | Sí | [LOG-013](codigos.md#log-013) |
| 7 | Holgura negativa | Tareas sin terminar con holgura total negativa | Ninguna | Sí | [LOG-014](codigos.md#log-014) y [LOG-015](codigos.md#log-015) |
| 8 | Duración alta | Tareas sin terminar con más de 44 días laborables de duración en la línea base | Hasta 5 % | Sí | [LOG-016](codigos.md#log-016) |
| 9 | Fechas no válidas | Fechas reales posteriores a la fecha de estado y fechas previstas anteriores a ella | Ninguna | Sí | [AVA-008](codigos.md#ava-008); [AVA-009](codigos.md#ava-009) a [AVA-011](codigos.md#ava-011) |
| 10 | Recursos | Tareas con duración sin recursos ni costo | Ninguna (si el cronograma tiene recursos) | No aplica | — |
| 11 | Tareas incumplidas | Tareas que debían terminar según la línea base y no terminaron, o terminaron tarde | Hasta 5 % | Sí | [AVA-014](codigos.md#ava-014) |
| 12 | Prueba del camino crítico | Un atraso en una tarea crítica mueve el fin del proyecto lo mismo | Debe cumplirse | No (prueba manual) | — |
| 13 | CPLI | Eficiencia necesaria para terminar en la fecha de la línea base | Al menos 0,95 | Sí | [AVA-015](codigos.md#ava-015) |
| 14 | BEI | Tareas terminadas frente a las que debían terminar según la línea base | Al menos 0,95 | Sí | [AVA-013](codigos.md#ava-013) |

`validar` revisa 12 de los 14 puntos. Los puntos 10 y 12 se explican en
[Lo que el programa no verifica](#6-lo-que-el-programa-no-verifica).

En los puntos con un umbral en %, la observación es una **advertencia** si se supera el umbral y una
**informativa** si hay casos pero quedan dentro del umbral. En los puntos sin tolerancia (adelantos,
holgura negativa, fechas no válidas), cada caso es una advertencia o un error.

---

## 4. Detalle de cada punto

### 4.1 Lógica (punto 1)

- **Qué cuenta:** tareas sin terminar sin predecesora o sin sucesora, sobre el total de tareas sin
  terminar. No cuentan la tarea que empieza el proyecto (no necesita predecesora) ni la que lo termina
  (no necesita sucesora). Un vínculo de una tarea resumen cuenta para sus subtareas, aunque se recomienda
  vincular las tareas de detalle ([LOG-010](codigos.md#log-010)).
- **Qué hacer:** cada tarea con al menos una predecesora y una sucesora, y cada rama del cronograma
  vinculada hasta el hito de fin. Sin vínculos, MS Project no mueve las fechas cuando cambian las
  duraciones o el avance.

### 4.2 Adelantos (punto 2)

- **Qué cuenta:** vínculos con posición negativa (por ejemplo, `FC-5 días`). Cada uno es una advertencia.
- **Qué hacer:** un adelanto supone que la sucesora puede empezar antes de que termine la predecesora,
  sin decir cuánto trabajo de la predecesora hace falta. Se reemplaza por un vínculo comienzo a comienzo
  con posición, o se divide la predecesora.

### 4.3 Posiciones (punto 3)

- **Qué cuenta:** vínculos con posición positiva (por ejemplo, `FC+10 días`), sobre el total de vínculos.
- **Qué hacer:** una espera fija suele ocultar trabajo (un plazo de aprobación, un curado, un transporte)
  que conviene representar como tarea, con su duración y su responsable.

### 4.4 Tipos de vínculo (punto 4)

- **Qué cuenta:** el % de vínculos fin a comienzo (FC). Los comienzo a fin (CF) se informan uno a uno.
- **Qué hacer:** los vínculos FC son los más claros de seguir. Los CC y FF son válidos, pero en exceso
  suelen indicar tareas demasiado largas que conviene dividir.

### 4.5 Delimitaciones rígidas (punto 5)

- **Qué cuenta:** tareas sin terminar con *Debe comenzar el*, *Debe finalizar el*, *No comenzar
  después del* o *No finalizar después del*. Las delimitaciones *No comenzar antes del* y *No finalizar
  antes del* no se cuentan: no impiden que la tarea se atrase.
- **Qué hacer:** usar *Lo antes posible* y, para una fecha comprometida, una **fecha límite**
  (*Información de la tarea > Avanzado*): MS Project avisa si se supera, sin fijar la fecha de la tarea.

### 4.6 Holgura alta (punto 6)

- **Qué cuenta:** tareas sin terminar con más de 44 días laborables de holgura total, la que calcula MS
  Project. No cuentan las tareas de las filas LOE. En un cronograma sin vínculos no se revisa: toda la
  holgura llega hasta el fin del proyecto y el problema es la falta de lógica (punto 1).
- **Qué hacer:** una holgura alta suele indicar que a la tarea le falta una sucesora: su resultado no
  alimenta a ninguna otra tarea, salvo al fin del proyecto. Se revisan sus sucesoras.

### 4.7 Holgura negativa (punto 7)

- **Qué cuenta:** tareas sin terminar con holgura total negativa. Cada una se informa con su holgura, de
  la más negativa a la menos negativa.
- **Cuándo aparece:** MS Project calcula holgura negativa cuando el fin previsto de una tarea supera su
  fecha límite o una delimitación. Sin fechas límite ni delimitaciones no hay holgura negativa, aunque el
  proyecto esté atrasado: en ese caso, el atraso frente a la línea base lo mide el CPLI (punto 13).
- **Qué hacer:** recuperar el atraso en el trabajo que queda (cambiar la secuencia, dividir tareas,
  sumar recursos) o, si la fecha comprometida cambió, actualizar la fecha límite.

> Para que el cronograma muestre la holgura frente a la fecha de fin comprometida, se recomienda poner
> esa fecha como **fecha límite del hito de fin**.

### 4.8 Duración alta (punto 8)

- **Qué cuenta:** tareas sin terminar con más de 44 días laborables de duración en la línea base (si la
  tarea no tiene línea base, su duración actual). No cuentan los hitos ni las tareas de las filas LOE.
- **Qué hacer:** dividir esas tareas en tareas más cortas, con entregables intermedios. Una tarea larga es
  difícil de seguir: su avance se estima en lugar de medirse.
- **Relación con [MET-004](codigos.md#met-004):** MET-004 revisa la duración del **paquete** (la fila del
  libro) frente a su método de medición (0/100 en más de 1 mes, 50/50 en más de 2). El punto 8 revisa
  cada **tarea** del cronograma.

### 4.9 Fechas no válidas (punto 9)

- **Qué cuenta:** con la fecha de estado,
  - *Comienzo real* o *Fin real* posteriores a la fecha de estado: avance registrado en el futuro
    ([AVA-008](codigos.md#ava-008), error);
  - trabajo pendiente programado antes de la fecha de estado: tareas sin terminar con *Fin* anterior a
    ella ([AVA-009](codigos.md#ava-009)) o sin avance con *Comienzo* anterior a ella
    ([AVA-010](codigos.md#ava-010)), resumidas en [AVA-011](codigos.md#ava-011).
- **Qué hacer:** corregir las fechas reales y reprogramar el trabajo pendiente después de la fecha de
  estado (*Proyecto > Actualizar proyecto*, opción para reprogramar el trabajo no completado). Sin eso,
  el fin previsto del proyecto no es una previsión realista.

### 4.10 Recursos (punto 10)

No se revisa. El programa no lee recursos ni costos de MS Project: el PV y el EV son el BAC que se
ingresa en el libro × el % planificado y el % ganado ([Cálculo del PV y el EV](calculo-pv-ev.md)). Un
cronograma sin recursos asignados es válido para este programa. Las herramientas de análisis suelen
tratar este punto como informativo, porque no todos los cronogramas tienen recursos asignados.

### 4.11 Tareas incumplidas (punto 11)

- **Qué cuenta:** de las tareas cuyo fin de línea base es anterior o igual a la fecha de estado, las que
  no terminaron, o que terminaron con un *Fin real* posterior a su fin de línea base. El mensaje separa
  las dos causas y lista las tareas.
- **Relación con el BEI (punto 14):** el BEI compara cantidades (cuántas terminaron frente a cuántas
  debían terminar) y da por buena una tarea terminada tarde. Este punto mide el cumplimiento de cada tarea.
  [AVA-012](codigos.md#ava-012) (informativa) lista además las tareas que debían **comenzar** y no comenzaron.

### 4.12 Prueba del camino crítico (punto 12)

No se revisa de forma automática: la prueba exige reprogramar el cronograma, y el programa abre el
archivo de MS Project en solo lectura. Se hace a mano, sobre una **copia** del archivo:

1. Guardar una copia del proyecto (*Archivo > Guardar como*).
2. Mostrar las tareas críticas (*Formato > Tareas críticas*) y elegir una tarea crítica sin terminar,
   cerca del comienzo del trabajo pendiente.
3. Anotar el fin del proyecto y aumentar la duración de la tarea en una cantidad grande y fácil de
   reconocer (por ejemplo, 100 días).
4. Comprobar que el fin del proyecto se atrasa **la misma cantidad** de días laborables.
5. Si se atrasa menos o no se mueve, el camino crítico no es continuo: buscar delimitaciones, tareas sin
   sucesoras o vínculos con tareas resumen entre esa tarea y el fin. Corregirlos en el archivo original.
6. Cerrar la copia sin guardar.

### 4.13 CPLI (punto 13)

El índice de longitud del camino crítico (*Critical Path Length Index*) indica con qué eficiencia hay que
ejecutar el trabajo que queda para terminar en la fecha de la línea base:

```text
CPLI = (días que quedan + holgura del proyecto) / días que quedan
```

- **Días que quedan:** días laborables desde la fecha de estado hasta el fin previsto del proyecto (el
  último *Fin* de las tareas revisadas).
- **Holgura del proyecto:** días laborables entre el fin previsto y el fin de la línea base: positiva si
  el fin previsto es anterior (adelanto) y negativa si es posterior (atraso). Es el mismo desvío de fin
  del [Cuadro de Mando](cuadro-de-mando.md#4-tendencia-y-plazo), con el signo contrario.
- **Ejemplo:** quedan 100 días laborables y el fin previsto está 5 días atrasado: CPLI = (100 − 5) / 100
  = 0,95, en el límite. Con 20 días de atraso, 0,80: hay que recuperar 20 días en los 100 que quedan.
- No se calcula sin fecha de estado, sin línea base o sin trabajo pendiente después de la fecha de
  estado.

> Las herramientas de análisis de cronogramas suelen usar la holgura del hito de fin frente a la fecha
> contractual. Si esa fecha es una **fecha límite** del hito de fin igual al fin de la línea base, el
> resultado es el mismo.

### 4.14 BEI (punto 14)

- **Qué cuenta:** tareas terminadas (todas, también las que terminaron antes de lo previsto) dividido por
  las tareas que debían terminar a la fecha de estado según la línea base. Es informativa si llega a
  0,95 y una advertencia si no; el mensaje lista las tareas que debían terminar y no terminaron.

---

## 5. Otros controles del cronograma

`validar` revisa también controles que no forman parte de los 14 puntos, pero que la DCMA (en sus
métricas DECM) y las guías de buenas prácticas incluyen:

| Control | Código | Por qué |
|---|---|---|
| Avance fuera de secuencia: la tarea comenzó (o terminó) antes de lo que permite el vínculo con su predecesora | [AVA-016](codigos.md#ava-016) y [AVA-017](codigos.md#ava-017) | El avance contradice la lógica: las fechas y la holgura del trabajo que queda pueden no ser realistas. Se corrige el avance o el vínculo. |
| Fechas reales incoherentes (*Fin real* sin *Comienzo real*, invertidas, 100 % sin *Fin real*...) | [AVA-001](codigos.md#ava-001) a [AVA-005](codigos.md#ava-005) | El avance mal registrado distorsiona el % ganado y los índices. |
| Tareas resumen con vínculos | [LOG-010](codigos.md#log-010) | Ocultan la lógica de las tareas de detalle. |
| Vínculos comienzo a fin (CF) | [LOG-007](codigos.md#log-007) | Tipo de vínculo poco habitual y difícil de seguir. |
| Proyecto o tareas sin línea base | [FEC-009](codigos.md#fec-009) a [FEC-011](codigos.md#fec-011) | Sin línea base no hay PV ni puntos 8, 11, 13 y 14. |
| Duraciones estimadas y tareas programadas manualmente | [FEC-006](codigos.md#fec-006) a [FEC-008](codigos.md#fec-008) | Una duración sin confirmar define el % planificado; una tarea manual no se reprograma con la lógica. |
| Método de medición frente a la duración del paquete | [MET-004](codigos.md#met-004) | 0/100 y fórmulas fijas solo en paquetes cortos (1 y 2 meses por defecto). |

---

## 6. Lo que el programa no verifica

| Aspecto | Motivo | Alternativa |
|---|---|---|
| Recursos (punto 10) | El programa no lee recursos ni costos: el PV es el BAC × el % planificado. | No aplica. |
| Prueba del camino crítico (punto 12) | Exige reprogramar el cronograma; el programa abre el archivo en solo lectura. | Prueba manual de la [sección 4.12](#412-prueba-del-camino-crítico-punto-12). |
| Continuidad del camino crítico y tareas LOE en el camino crítico | Dependen del cálculo de MS Project. | La prueba del punto 12 los pone en evidencia. |
| Análisis de riesgo del cronograma (simulación de Monte Carlo) | Necesita rangos de duración y una herramienta de simulación. | Herramientas especializadas. |
| Calidad de las estimaciones de duración | No se puede comprobar con los datos del cronograma. | Revisión del responsable de cada paquete. |

---

## 7. Uso

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --config config.toml
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --fecha-corte 2026-10-31 --todas
```

- Los puntos 1 a 8 se muestran en el tema **Lógica del cronograma**; los puntos 9, 11, 13 y 14, en
  **Avance a la fecha de estado**.
- Se recomienda ejecutar `validar` antes de `generar` (con la línea base recién guardada: puntos 1 a 8)
  y antes de cada `actualizar` (con el avance cargado: todos los puntos).
- Un control que no aplica al proyecto, o una observación ya revisada, se puede silenciar con
  `[validacion] silenciar` o `--silenciar` (por ejemplo `LOG-013`, o `LOG-015:1.2.1` para una sola
  tarea) ([Silenciar observaciones](codigos.md#2-silenciar-observaciones)).
- Los controles DCMA son advertencias o informativas: no cambian el código de salida de `validar`,
  salvo el avance registrado en el futuro ([AVA-008](codigos.md#ava-008)), que es un error.

Para que la evaluación sea útil, el cronograma de MS Project debe tener:

- la **línea base** guardada ([Preparación en MS Project](preparacion-ms-project.md#6-línea-base));
- la **fecha de estado** de cada cierre y el avance registrado hasta esa fecha
  ([Fecha de estado y registro del avance](preparacion-ms-project.md#7-fecha-de-estado-y-registro-del-avance));
- las tareas vinculadas, con *Lo antes posible* y **fechas límite** en lugar de delimitaciones;
- el trabajo pendiente reprogramado después de la fecha de estado (*Proyecto > Actualizar proyecto*).

---

## 8. Referencias

| Recurso | Contenido |
|---|---|
| [GAO Schedule Assessment Guide (GAO-16-89G)](https://www.gao.gov/products/gao-16-89g) | Las 10 buenas prácticas de un cronograma confiable, con los controles de lógica, holgura, camino crítico y actualización. |
| [DCMA — Earned Value Management Systems](https://www.dcma.mil/HQ/EVMS) | Página de la DCMA sobre la revisión de sistemas EVM y sus métricas de cumplimiento (DECM). |
| [NDIA IPMD — Guías de la división](https://www.ndia.org/divisions/ipmd/division-guides-and-resources) | Guía de intención de la norma EIA-748 y guía de buenas prácticas de planificación y programación (PASEG). |

Más referencias en [Recursos externos](recursos-externos.md).
