# Norma EIA-748-E: requisitos y cobertura del programa

[Inicio](../README.md) › [Documentación](README.md) › Norma EIA-748-E

La norma **EIA-748** es la referencia de EE. UU. para los sistemas de gestión del valor ganado (EVMS).
Este documento resume las 27 guías de su revisión E e indica, para cada una, qué hace el programa, qué
controles automáticos la verifican y qué queda a cargo de la organización.

---

## Contenido

1. [Qué es la norma](#1-qué-es-la-norma)
2. [Cómo leer este documento](#2-cómo-leer-este-documento)
3. [Resumen de cobertura](#3-resumen-de-cobertura)
4. [Organización (guías 1 a 4)](#4-organización-guías-1-a-4)
5. [Planificación, programación y presupuesto (guías 5 a 13)](#5-planificación-programación-y-presupuesto-guías-5-a-13)
6. [Avance y recolección de datos (guías 14 a 16)](#6-avance-y-recolección-de-datos-guías-14-a-16)
7. [Análisis e informes de gestión (guías 17 a 23)](#7-análisis-e-informes-de-gestión-guías-17-a-23)
8. [Revisiones y mantenimiento de datos (guías 24 a 27)](#8-revisiones-y-mantenimiento-de-datos-guías-24-a-27)
9. [Lo que aporta la organización](#9-lo-que-aporta-la-organización)
10. [Referencias](#10-referencias)

---

## 1. Qué es la norma

- **EIA-748** (*Earned Value Management Systems*) la publica SAE International y la mantiene la división
  de gestión integrada de programas de la NDIA (IPMD). La **revisión E** (febrero de 2026) reemplaza a la
  D: reduce las guías de 32 a 27 (fusiona algunas, elimina cuatro y agrega dos) y las ordena según el
  ciclo de vida del proyecto, en las mismas cinco categorías.
- La norma dice **qué** debe tener un sistema EVM, no **cómo** hacerlo. La NDIA publica una **guía de
  intención** (*EIA-748-E Intent Guide*, mayo de 2026) con el propósito de cada guía y sus atributos
  verificables. Este documento se basa en ella.
- En los contratos del Departamento de Defensa de EE. UU. por encima de ciertos montos, el contratista
  debe usar un sistema que cumpla la norma y, en los mayores, validado por la DCMA. Fuera de esos
  contratos es una referencia de buenas prácticas; su equivalente internacional es la norma
  ISO 21508.

> **La norma evalúa el sistema de gestión de la organización, no un programa informático.** El sistema
> son los procesos (descritos en una *descripción del sistema EVM*), las herramientas y las personas.
> Ningún programa "cumple" la norma por sí solo: puede ser una de las herramientas de un sistema que la
> cumpla. Este documento indica qué partes de ese sistema aporta el Exportador EVM.

**Correspondencia de términos.** En este programa, las **filas de ingreso** del libro (paquetes de
trabajo y filas `-D Otras tareas`) son el nivel donde se ingresan el BAC y el AC y se mide el EV. Según
el nivel elegido (`nivel_maximo` o `lista_edt`), equivalen a las **cuentas de control** de la norma o a
sus **paquetes de trabajo**. El PV de todas las filas es la línea base de medición del desempeño (PMB).

---

## 2. Cómo leer este documento

Cada guía tiene una cobertura:

| Cobertura | Significado |
|---|---|
| **Cubre** | El programa produce lo que pide la guía y controla sus datos. La organización sigue siendo responsable del proceso. |
| **Parcial** | El programa cubre una parte; el resto lo hace la organización o se registra fuera del programa. |
| **No cubre** | El programa no tiene la información que pide la guía, aunque podría incorporarla. |
| **Fuera de alcance** | Corresponde a otros sistemas (contabilidad, costos indirectos, materiales) o no aplica por diseño: el programa calcula el PV y el EV como el BAC × el %, sin leer costos ni recursos de MS Project. |

La columna **Controles** indica los códigos de los controles automáticos relacionados
([Códigos de errores y advertencias](codigos.md)); un guion indica que no hay control automático.

---

## 3. Resumen de cobertura

| # | Guía | Cobertura | Controles |
|---|---|---|---|
| **Organización** | | | |
| 1 | Descomponer el alcance con una EDT | Parcial | [EDT](codigos.md#43-edt--estructura-y-códigos-edt), [VIS](codigos.md#44-vis--campo-visible) |
| 2 | Identificar las responsabilidades de la organización | Parcial | [EMP-003](codigos.md#emp-003) a [EMP-005](codigos.md#emp-005) |
| 3 | Integrar la EDT y la OBS en cuentas de control | Parcial | — |
| 4 | Integrar los procesos de gestión con la EDT y la OBS | Parcial | [EDT-009](codigos.md#edt-009), [EDT-011](codigos.md#edt-011), [ACT-011](codigos.md#act-011) a [ACT-017](codigos.md#act-017) |
| **Planificación, programación y presupuesto** | | | |
| 5 | Identificar los indicadores de avance | Parcial | [HIT](codigos.md#46-hit--hitos-y-pesos), [FEC-004](codigos.md#fec-004), [FEC-005](codigos.md#fec-005) |
| 6 | Programar el trabajo autorizado | Cubre | [LOG](codigos.md#410-log--lógica-del-cronograma), [AVA-008](codigos.md#ava-008) a [AVA-015](codigos.md#ava-015) ([Evaluación DCMA](evaluacion-dcma.md)) |
| 7 | Establecer y mantener una línea base distribuida en el tiempo | Cubre | [FEC-009](codigos.md#fec-009) a [FEC-012](codigos.md#fec-012), [CFG-044](codigos.md#cfg-044), [CFG-045](codigos.md#cfg-045), [ACT-022](codigos.md#act-022) |
| 8 | Autorizar alcance, plazo y presupuesto por elementos de costo | Fuera de alcance | — |
| 9 | Planificar en paquetes de trabajo y de planificación | Parcial | [MET-004](codigos.md#met-004), [LOG-016](codigos.md#log-016), [HIT-005](codigos.md#hit-005), [EDT-018](codigos.md#edt-018) |
| 10 | Establecer los criterios de medición de cada paquete | Parcial | [MET](codigos.md#47-met--método-de-medición), [LEC-019](codigos.md#lec-019) |
| 11 | Aplicar tasas indirectas a los presupuestos | Fuera de alcance | — |
| 12 | Identificar el presupuesto no distribuido y la reserva de gestión | No cubre | — |
| 13 | Conciliar con el objetivo de costo | No cubre | — |
| **Avance y recolección de datos** | | | |
| 14 | Medir el avance y determinar el EV | Cubre | [AVA](codigos.md#49-ava--avance-a-la-fecha-de-estado), [ACT-004](codigos.md#act-004) a [ACT-006](codigos.md#act-006), [ACT-019](codigos.md#act-019) |
| 15 | Registrar los costos reales por elementos de costo | Parcial | [ACT-026](codigos.md#act-026) |
| 16 | Contabilizar el material comprado | Fuera de alcance | — |
| **Análisis e informes de gestión** | | | |
| 17 | Generar las variaciones de cronograma y de costo | Cubre | [ACT-018](codigos.md#act-018) |
| 18 | Identificar y evaluar las variaciones significativas | Parcial | — |
| 19 | Evaluar las variaciones de costos indirectos | Fuera de alcance | — |
| 20 | Actualizar el EAC de cada cuenta de control | Parcial | — |
| 21 | Resumir, revisar y evaluar los datos de desempeño | Cubre | — |
| 22 | Implementar acciones de gestión | Parcial | — |
| 23 | Revisar el EAC del proyecto | Parcial | — |
| **Revisiones y mantenimiento de datos** | | | |
| 24 | Incorporar los cambios del cliente | Parcial | [ACT-011](codigos.md#act-011) a [ACT-017](codigos.md#act-017), [MIG](codigos.md#412-mig--migración---migrar-desde) |
| 25 | Documentar y conciliar las replanificaciones internas | Parcial | [ACT-015](codigos.md#act-015), [FEC-010](codigos.md#fec-010) |
| 26 | Controlar los cambios retroactivos | Parcial | [ACT-007](codigos.md#act-007), [ACT-020](codigos.md#act-020), [ACT-021](codigos.md#act-021), [ACT-024](codigos.md#act-024), [ACT-027](codigos.md#act-027) |
| 27 | Línea base o cronograma por encima del objetivo (OTB/OTS) | No cubre | — |

**En resumen:** de las 27 guías, el programa **cubre 5**, cubre **en parte 15**, **no cubre 3** que
podría incorporar, y **4 están fuera de su alcance** (costos por elemento, indirectos y materiales). Lo
más sólido es lo que se apoya en el cronograma de MS Project: la EDT, la lógica del cronograma, la
línea base distribuida en el tiempo, la medición del EV con métodos objetivos y las variaciones. Lo más
débil es el control de los cambios y de los presupuestos que no son el BAC de cada paquete (reserva,
presupuesto no distribuido, presupuesto del contrato).

---

## 4. Organización (guías 1 a 4)

### Guía 1 · Descomponer el alcance con una EDT

- **Qué exige:** una sola EDT para gestionar el proyecto, con todo el trabajo autorizado (también los
  cambios), extendida al menos hasta las cuentas de control y con un diccionario que defina el alcance de
  cada elemento.
- **Qué hace el programa:** usa el código EDT de MS Project como estructura del libro y revisa su
  coherencia: formato, dependencia de la tarea superior, niveles salteados, orden y duplicados
  (`EDT-001` a `EDT-009`). Informa el trabajo que queda fuera de la EDT: tareas sin código o fuera de la
  lista EDT que no dependen de ningún elemento (`EDT-019`, `EDT-020`). Con una sola planificación por
  libro, garantiza una sola EDT. `generar-fichas` crea una ficha por elemento con su descripción, fechas
  y vínculos, que puede servir de base del diccionario EDT.
- **Qué queda a cargo de la organización:** que la EDT contenga **todo** el alcance del contrato
  (comparación con el enunciado del trabajo) y el diccionario EDT aprobado.

### Guía 2 · Identificar las responsabilidades de la organización

- **Qué exige:** una estructura organizativa (OBS) a la que se asigna todo el trabajo, incluidos los
  subcontratistas principales.
- **Qué hace el programa:** en [multi-empresa](multiempresa.md), asigna cada tarea a una de dos empresas
  (o a las dos) y avisa de las tareas sin empresa (`EMP-003`, `EMP-004`) o con una empresa no reconocida
  (`EMP-005`). No tiene un campo de responsable (área o persona) por paquete.
- **Qué queda a cargo de la organización:** la OBS y la asignación de cada paquete a un responsable.

### Guía 3 · Integrar la EDT y la OBS en cuentas de control

- **Qué exige:** cuentas de control en la intersección de la EDT y la OBS, cada una con un responsable
  (CAM), alcance, plazo y presupuesto.
- **Qué hace el programa:** las filas de ingreso tienen alcance (sus tareas), plazo (sus fechas de línea
  base) y presupuesto (su BAC). En multi-empresa, cada fila tiene un PV, un EV y un AC por empresa. No
  registra el responsable de cada fila.
- **Qué queda a cargo de la organización:** la matriz de responsabilidades y la autorización de cada
  cuenta de control.

### Guía 4 · Integrar los procesos de gestión con la EDT y la OBS

- **Qué exige:** que el cronograma, el presupuesto, la autorización del trabajo y la acumulación de
  costos compartan los mismos códigos, para medir el desempeño por EDT u OBS.
- **Qué hace el programa:** el código EDT de MS Project es la clave común del cronograma, el PV, el EV y
  el AC; por eso los códigos no se pueden repetir (`EDT-009`, `EDT-011`). `actualizar` comprueba que el
  libro y el cronograma sigan correspondiéndose (`ACT-011` a `ACT-017`). La
  [depuración de costes](depuracion-costes.md) relaciona la EDT del cronograma con el plan de costes de
  una empresa.
- **Qué queda a cargo de la organización:** que el AC se registre en contabilidad con los mismos
  códigos (números de imputación por paquete).

---

## 5. Planificación, programación y presupuesto (guías 5 a 13)

### Guía 5 · Identificar los indicadores de avance

- **Qué exige:** identificar en el cronograma los productos, hitos y objetivos técnicos que miden el
  avance, con criterios de cumplimiento definidos de antemano.
- **Qué hace el programa:** revisa los hitos y sus pesos: hitos con duración, tareas de duración 0 sin
  marca de hito, pesos fuera de rango o que no suman el 100 % (`FEC-004`, `FEC-005`, `HIT-001` a
  `HIT-009`). El método «Hitos ponderados» mide el EV con los hitos cumplidos.
- **Qué queda a cargo de la organización:** el criterio de cumplimiento de cada hito (qué evidencia lo
  da por cumplido). Puede registrarse en la descripción de la tarea, que muestran las fichas.

### Guía 6 · Programar el trabajo autorizado

- **Qué exige:** un cronograma con la secuencia del trabajo y sus dependencias, trazable a la EDT, con
  hitos y un camino crítico para todo el período del contrato.
- **Qué hace el programa:** `validar` revisa 12 de los 14 puntos de la evaluación DCMA: lógica,
  adelantos, posiciones, tipos de vínculo, delimitaciones, holgura, duración, fechas no válidas, tareas
  incumplidas, CPLI y BEI ([Evaluación DCMA](evaluacion-dcma.md)). La trazabilidad a la EDT la revisan los
  controles de la guía 1.
- **Qué queda a cargo de la organización:** la prueba del camino crítico (manual, en una copia del
  archivo) y la carga de recursos, si el contrato la exige.

### Guía 7 · Establecer y mantener una línea base distribuida en el tiempo

- **Qué exige:** una línea base (PMB) con alcance, plazo y presupuesto por cuenta de control,
  distribuida en el tiempo **de forma coherente con el cronograma**.
- **Qué hace el programa:** el PV de cada fila es su BAC distribuido en los meses según las fechas de la
  línea base de MS Project: la coherencia con el cronograma está garantizada. Avisa si falta la línea base
  o la tienen solo algunas tareas (`FEC-009` a `FEC-012`), si parte del PV cae fuera del horizonte
  (`CFG-044`, `CFG-045`) y si hay filas sin BAC (`ACT-022`, `ACT-023`).
- **Qué queda a cargo de la organización:** el BAC de cada fila, ingresado en el libro, y el presupuesto
  de los paquetes lejanos que todavía no se detallan (ver guía 12).

### Guía 8 · Autorizar alcance, plazo y presupuesto por elementos de costo

- **Qué exige:** presupuestos por elemento de costo (mano de obra, material, subcontratos, otros
  directos), en importes, horas u otras unidades.
- **Fuera de alcance:** el BAC de cada fila es un importe único. El programa no lee recursos ni costos de
  MS Project. La desagregación por elemento de costo se lleva en el sistema de presupuestos de la
  organización.

### Guía 9 · Planificar en paquetes de trabajo y de planificación

- **Qué exige:** dividir cada cuenta de control en paquetes de trabajo (cortos o con hitos intermedios,
  con fechas y presupuesto) y paquetes de planificación para el trabajo lejano. La suma de los paquetes es
  el presupuesto de la cuenta.
- **Qué hace el programa:** las filas resumen del libro suman las de sus filas inferiores (fórmulas), y
  el trabajo sin fila propia se reúne en filas `-D Otras tareas` con su propio BAC (`EDT-018`). Avisa de
  los paquetes largos con métodos 0/100 o fórmulas fijas (`MET-004`), de las tareas largas (`LOG-016`) y
  de las filas sin % planificado (`HIT-005`). No distingue los paquetes de planificación.
- **Qué queda a cargo de la organización:** decidir qué trabajo lejano queda como paquete de
  planificación y detallarlo a tiempo.

### Guía 10 · Establecer los criterios de medición de cada paquete

- **Qué exige:** una sola técnica de EV por paquete, que mida el avance técnico (no el consumo de
  recursos), con respaldo cuantificable en los paquetes largos. El nivel de esfuerzo (LOE) se reserva
  para el trabajo que no se puede medir, se mantiene al mínimo y su EV es igual a su PV.
- **Qué hace el programa:** un método por fila de ingreso, del campo de MS Project o de la columna C
  ([Métodos de medición](metodos-de-medicion.md)). Avisa de los métodos sin definir, no reconocidos o
  distintos en las subtareas (`MET-001` a `MET-005`, `LEC-019`) y de los métodos 0/100 o de fórmula fija
  en paquetes largos (`MET-004`). El LOE gana su PV.
- **Qué falta:** no controla el peso del LOE en el BAC ni advierte del *% completado* sin respaldo
  objetivo.

### Guía 11 · Aplicar tasas indirectas a los presupuestos

- **Fuera de alcance:** las tasas de costos indirectos y sus presupuestos son del sistema de costos de
  la organización. Si el BAC incluye indirectos, el programa los trata como parte del importe.

### Guía 12 · Identificar el presupuesto no distribuido y la reserva de gestión

- **Qué exige:** identificar el presupuesto no distribuido (UB) y la reserva de gestión (MR), con un
  registro mensual de sus movimientos.
- **No cubre:** el libro solo tiene el BAC de cada fila; la reserva y el presupuesto no distribuido se
  llevan fuera del libro.

### Guía 13 · Conciliar con el objetivo de costo

- **Qué exige:** que el presupuesto del contrato (CBB) sea igual a la suma de los presupuestos internos
  más la reserva de gestión: CBB = PMB + MR, y PMB = suma de las cuentas de control + UB.
- **No cubre:** el BAC total del libro es la suma de las filas, pero el libro no tiene el presupuesto del
  contrato ni la reserva para conciliarlos.

---

## 6. Avance y recolección de datos (guías 14 a 16)

### Guía 14 · Medir el avance y determinar el EV

- **Qué exige:** actualizar el cronograma a la fecha de estado y medir el avance físico con los criterios
  definidos de antemano. El EV del material no puede ser anterior a su recepción.
- **Qué hace el programa:** `actualizar` calcula el % ganado de cada fila con su método y el avance de
  MS Project, con la fecha de estado ([Cálculo del PV y el EV](calculo-pv-ev.md)). `validar` revisa que el
  cronograma esté actualizado: fechas reales coherentes, avance en el futuro, trabajo pendiente antes de
  la fecha de estado y avance fuera de secuencia (`AVA-001` a `AVA-017`). `actualizar` exige la fecha
  de corte y avisa si difiere de la fecha de estado (`ACT-004` a `ACT-006`, `ACT-019`).
- **Fuera de alcance:** el material (el programa no distingue elementos de costo).

### Guía 15 · Registrar los costos reales por elementos de costo

- **Qué exige:** AC directos e indirectos, por elemento de costo, de un sistema controlado por la
  contabilidad general, **en el mismo período que el EV**. Si el costo real todavía no está registrado,
  se usa una estimación.
- **Qué hace el programa:** el AC de cada fila y mes se ingresa a mano en la hoja `AC` (`ACT-026`), con
  la misma fecha de corte que el EV. No se conecta con la contabilidad ni distingue elementos de costo.
- **Qué falta:** controles de coherencia entre el EV y el AC del mes (EV sin AC, AC sin EV, AC en
  paquetes terminados).
- **Qué queda a cargo de la organización:** el AC del ERP con la misma fecha de corte y las estimaciones
  de costos devengados no facturados ([Proceso de reporte](proceso-reporte-evm.md#2-roles)).

### Guía 16 · Contabilizar el material comprado

- **Fuera de alcance:** el inventario, los consumos y el material sobrante son del sistema de materiales
  de la organización.

---

## 7. Análisis e informes de gestión (guías 17 a 23)

### Guía 17 · Generar las variaciones de cronograma y de costo

- **Qué exige:** **al menos cada mes**, la SV, la CV y el VAC por cuenta de control y en los niveles
  superiores, con AC conciliable con la contabilidad.
- **Qué hace el programa:** el libro calcula el PV, el EV y el AC de cada mes, y la SV, la CV, los
  índices, el EAC y el VAC del proyecto ([Cuadro de Mando](cuadro-de-mando.md)) y de cada elemento EDT
  ([Indicadores EDT](libro-excel.md#indicadores-edt)). Avisa si la fecha de corte no es fin de mes
  (`ACT-018`).
- **Qué queda a cargo de la organización:** el informe mensual. La norma pide la información **cada
  mes**: un informe cada dos meses no la cumple. Si el informe es bimestral, conviene igualmente
  actualizar el libro al cierre de cada mes, con la fecha de estado de ese mes.

### Guía 18 · Identificar y evaluar las variaciones significativas

- **Qué exige:** cada mes y por cuenta de control, identificar las variaciones que superan los umbrales
  y explicar su causa, su impacto y la acción correctiva, que se sigue hasta su cierre.
- **Qué hace el programa:** la hoja [Análisis de desviaciones](analisis-desviaciones.md) lista cada mes
  los paquetes con el SPI o el CPI fuera del rango verde, con el plazo según MS Project y columnas para la
  causa, el impacto, la acción, el responsable y la fecha.
- **Qué falta:** no comprueba que las explicaciones estén completas. Los umbrales son sobre los índices,
  no sobre importes (SV, CV, VAC) ni sobre el desvío de fin.

### Guía 19 · Evaluar las variaciones de costos indirectos

- **Fuera de alcance:** el programa no distingue los costos indirectos.

### Guía 20 · Actualizar el EAC de cada cuenta de control

- **Qué exige:** que el responsable de cada cuenta de control estime el costo del trabajo que queda
  (ETC) a partir del cronograma y de los recursos, y calcule el EAC y el VAC de la cuenta.
- **Qué hace el programa:** calcula el EAC estadístico de cada fila (BAC / CPI) en `Indicadores EDT`.
- **Qué falta:** un lugar para el EAC del responsable de cada paquete y su comparación con el EAC
  estadístico.

### Guía 21 · Resumir, revisar y evaluar los datos de desempeño

- **Qué exige:** resumir los datos y las variaciones por la EDT (u OBS), con el TCPI y estimaciones
  independientes del EAC.
- **Qué hace el programa:** las filas resumen y el proyecto suman sus filas; el `Cuadro de Mando`
  muestra el TCPI, tres métodos de EAC, el Cronograma Ganado, la tendencia de los índices y el plazo
  según MS Project. En multi-empresa, los valores de cada empresa.

### Guía 22 · Implementar acciones de gestión

- **Qué exige:** implementar las acciones que surgen del análisis, verificar que se cumplan y que sean
  razonables, y considerar los riesgos en el EAC.
- **Qué hace el programa:** el análisis de desviaciones registra la acción, el responsable y la fecha
  compromiso, y la columna `Antecedente` indica si el paquete ya estaba fuera de umbral.
- **Qué falta:** el estado de cada acción (abierta o cerrada) y el aviso de las acciones vencidas.

### Guía 23 · Revisar el EAC del proyecto

- **Qué exige:** un EAC del proyecto basado en el desempeño, en los recursos que faltan y en los riesgos,
  comparado con el presupuesto del contrato.
- **Qué hace el programa:** EAC por tres métodos, ETC, VAC, TCPI y término estimado con el SPI(t).
- **Qué falta:** el EAC detallado (suma de los EAC de los responsables) y la comparación con el
  presupuesto del contrato y la reserva.

---

## 8. Revisiones y mantenimiento de datos (guías 24 a 27)

### Guía 24 · Incorporar los cambios del cliente

- **Qué exige:** incorporar a tiempo los cambios del cliente, documentando y conciliando su efecto en el
  alcance, el plazo y el presupuesto.
- **Qué hace el programa:** `generar --migrar-desde` crea el libro con la línea base o la EDT nuevas y
  traslada el BAC, el % ganado y el AC de cada fila (códigos `MIG`). `actualizar` avisa de los cambios del
  cronograma que el libro no tiene: tareas eliminadas, inactivas o nuevas, y cambios del campo `Visible`
  o de la lista EDT (`ACT-011` a `ACT-017`).
- **Qué falta:** el registro de los cambios de BAC (fecha, importe anterior y nuevo, motivo), propuesto en
  el ticket [#24](https://github.com/MarceloSalazar/evm-automation/issues/24).

### Guía 25 · Documentar y conciliar las replanificaciones internas

- **Qué exige:** documentar los cambios internos de alcance, plazo y presupuesto, y comparar la línea
  base del período con la del anterior para identificar todas las diferencias.
- **Qué hace el programa:** avisa de las tareas nuevas sin fila (`ACT-015`) y de las tareas sin línea base
  (`FEC-010`).
- **Qué falta:** comparar la línea base entre dos lecturas (cambios de fechas, del % planificado de cada
  mes o del BAC).

### Guía 26 · Controlar los cambios retroactivos

- **Qué exige:** no modificar los valores ya informados de AC, EV o presupuesto, salvo correcciones de
  errores, ajustes contables o cambios dirigidos, siempre explicados.
- **Qué hace el programa:** `actualizar` no carga meses anteriores a la última fecha de corte sin avisar
  (`ACT-020`), no sobrescribe meses posteriores sin `--forzar` (`ACT-007`, `ACT-021`), avisa de las celdas
  sin la fórmula del libro (`ACT-024`, `ACT-025`) y del % ganado negativo de un mes (`ACT-027`).
- **Limitación importante:** el PV y el EV de cada mes son el BAC × el %. **Un cambio del BAC de una fila
  recalcula el PV y el EV de todos sus meses, también los ya informados**
  ([Cálculo del PV y el EV](calculo-pv-ev.md#13-cambios-de-bac-y-de-línea-base)). Para la norma es un
  cambio retroactivo. Hasta que el programa lo registre, cada cambio de BAC se debe documentar fuera del
  libro, con el valor anterior.
- **Qué falta:** detectar cambios en los valores de los meses ya informados.

### Guía 27 · Línea base o cronograma por encima del objetivo (OTB/OTS)

- **Qué exige:** cuando el trabajo que queda no se puede gestionar con el presupuesto o el plazo del
  contrato, proponer, documentar y aplicar (con aviso previo al cliente) una línea base mayor que el
  presupuesto del contrato o un cronograma mayor que el plazo del contrato.
- **No cubre:** es una decisión de gestión. El programa muestra las señales que la anticipan (TCPI muy
  por encima del CPI, CPLI bajo, desvío de fin), pero no las evalúa juntas ni registra la replanificación.

---

## 9. Lo que aporta la organización

Para cumplir la norma, la organización debe aportar lo que el programa no hace:

- una **descripción del sistema EVM** con sus procesos, roles y responsables
  ([Proceso de reporte EVM](proceso-reporte-evm.md#2-roles));
- la **OBS** y un **responsable** por cuenta de control;
- los **presupuestos** por elemento de costo, la **reserva de gestión**, el **presupuesto no distribuido** y
  la conciliación con el presupuesto del contrato;
- el **AC** de la contabilidad, con la misma fecha de corte que el EV, y los costos indirectos y de
  material;
- el **control de cambios** de la línea base, con su registro
  ([Cambios de línea base](proceso-reporte-evm.md#5-cambios-de-línea-base));
- el **EAC** de cada responsable y del proyecto, con los riesgos;
- la **frecuencia mensual** del análisis y el seguimiento de las acciones correctivas;
- la **vigilancia** del sistema (revisiones internas periódicas).

---

## 10. Referencias

| Recurso | Contenido |
|---|---|
| [SAE EIA-748-E](https://saemobilus.sae.org/standards/eia748e-earned-value-management-systems) | La norma (de pago). |
| [NDIA IPMD — Guías de la división](https://www.ndia.org/divisions/ipmd/division-guides-and-resources) | Guía de intención de la norma (*EIA-748-E Intent Guide*), con el propósito y los atributos verificables de cada guía. |
| [Humphreys & Associates — Revisión E](https://blog.humphreys-assoc.com/eia-748-standard-for-evms-revision-e-published/) | Resumen de los cambios de la revisión E y correspondencia con las guías de la revisión D. |
| [DoD — Integrated Program Management](https://www.acq.osd.mil/asda/dpc/api/ipm/policy-guidance.html) | Guía de interpretación de la norma del Departamento de Defensa (EVMSIG). |
| [Evaluación DCMA de 14 puntos](evaluacion-dcma.md) | Controles de calidad del cronograma que aplica `validar` (guía 6). |

Más referencias en [Recursos externos](recursos-externos.md).
