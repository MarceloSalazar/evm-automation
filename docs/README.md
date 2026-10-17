# Documentación

[Inicio](../README.md) › Documentación

Documentación detallada del Exportador EVM desde MS Project. La presentación, la instalación y el uso
rápido están en el [README principal](../README.md).

---

## Por dónde empezar

| Si se necesita... | Leer |
|---|---|
| Ver el programa funcionando sin MS Project | [Tutorial con el proyecto de ejemplo](ejemplo-paso-a-paso.md) |
| Preparar un proyecto real por primera vez | [Preparación en MS Project](preparacion-ms-project.md) → [Presupuesto (BAC)](presupuesto-bac.md) → [Métodos de medición](metodos-de-medicion.md) |
| Organizar el cierre mensual | [Proceso de reporte EVM](proceso-reporte-evm.md) |
| Reunir varios archivos de MS Project en un informe | [Varios archivos de MS Project](varios-archivos.md) |
| Interpretar los indicadores y los gráficos | [Cuadro de Mando e indicadores](cuadro-de-mando.md) |
| Explicar las desviaciones del mes | [Análisis de desviaciones](analisis-desviaciones.md) |
| Entender un valor del libro | [Cálculo del PV](calculo-pv.md) y [Cálculo de EV y AC](calculo-ev-ac.md) |
| Resolver un mensaje de error | [Validaciones y solución de problemas](validaciones-y-problemas.md) |
| Presentar el programa al equipo de proyecto | [Presentación de 30 minutos](presentacion-jefes-de-proyecto.pptx) ([PDF](presentacion-jefes-de-proyecto.pdf)) |

---

## Índice

**Uso del programa**

| Documento | Contenido |
|---|---|
| [Tutorial con el proyecto de ejemplo](ejemplo-paso-a-paso.md) | Ciclo completo con el XML de ejemplo: generar el libro, simular 24 meses de avance y revisar el resultado. |
| [Proceso de reporte EVM](proceso-reporte-evm.md) | Puesta en marcha, cierre mensual, cambios de línea base, roles, calendario, buenas prácticas y lista de control. |
| [Comandos y opciones](comandos.md) | `generar` y `actualizar`, todas sus opciones y ejemplos de salida. |
| [Configuración](configuracion.md) | Claves del archivo `config.toml` y ejemplos. |
| [Presentación para el equipo de proyecto](presentacion-jefes-de-proyecto.pptx) ([PDF](presentacion-jefes-de-proyecto.pdf)) | 36 diapositivas para unos 30 minutos, con notas para el presentador: conceptos EVM y semáforo, preparación en MS Project (incluida la línea base), `generar` y `actualizar`, Cuadro de Mando, análisis de desviaciones y cierre mensual. Se abre con PowerPoint o LibreOffice Impress. |

**Preparación en MS Project**

| Documento | Contenido |
|---|---|
| [Preparación del proyecto en MS Project](preparacion-ms-project.md) | Lista de control, código EDT, campos personalizados, línea base, fecha de estado, registro del avance y XML. |
| [Presupuesto (BAC) de cada tarea](presupuesto-bac.md) | Recursos de trabajo, materiales y de costo, costo fijo; guardar y verificar la línea base; errores frecuentes. |
| [Métodos de medición](metodos-de-medicion.md) | 0/100, fórmulas fijas, hitos ponderados, % completado, % físico, unidades, LOE, CPTR y el método de valor acumulado de MS Project. |
| [Multi-empresa (opcional)](multiempresa.md) | PV, EV y AC de dos empresas: preparación, reparto, cálculo y ejemplo. |
| [Varios archivos de MS Project (opcional)](varios-archivos.md) | Un informe con los elementos EDT de varios archivos: códigos EDT, `[[proyectos]]`, cierre mensual, cambios de archivos y ejemplo. |

**Cálculos y libro**

| Documento | Contenido |
|---|---|
| [Cálculo del PV](calculo-pv.md) | Origen del PV mensual, planificación, niveles por código EDT, tareas sin código y horizonte. |
| [Cálculo de EV y AC](calculo-ev-ac.md) | Datos que se leen, EV por método, AC, valor del mes, fecha de corte y controles. |
| [El libro Excel](libro-excel.md) | Hojas, reglas de uso y `--migrar-desde`. |
| [Cuadro de Mando e indicadores](cuadro-de-mando.md) | Tablero, gráficos, fórmulas, Cronograma Ganado y cómo leer los indicadores. |
| [Análisis de desviaciones](analisis-desviaciones.md) | Paquetes fuera de umbral de cada mes, causa, impacto y acción correctiva; comando `analizar`. |

**Referencia**

| Documento | Contenido |
|---|---|
| [Validaciones y solución de problemas](validaciones-y-problemas.md) | Hoja `Validación`, limitaciones conocidas y mensajes frecuentes. |
| [Glosario](glosario.md) | Siglas EVM y su equivalencia con los campos de MS Project en español (CPTP, CPTR, CRTR...). |
| [Registro de cambios](../CHANGELOG.md) | Cambios de cada versión: nuevas funciones, correcciones y otros. |
| [Recursos externos](recursos-externos.md) | Documentación de Microsoft, estándares y guías de buenas prácticas. |
| [Desarrollo](desarrollo.md) | Pruebas automáticas, estructura del código, flujo del repositorio e imágenes de la documentación. |

---

## Mapa de la documentación

```mermaid
flowchart LR
    T["Tutorial"] --> P["Proceso de<br/>reporte EVM"]
    P --> MSP["Preparación en<br/>MS Project"]
    MSP --> BAC["Presupuesto<br/>(BAC)"]
    MSP --> MET["Métodos de<br/>medición"]
    MSP --> ME["Multi-empresa"]
    BAC --> CPV["Cálculo<br/>del PV"]
    MET --> CEV["Cálculo de<br/>EV y AC"]
    CPV & CEV --> LIB["El libro<br/>Excel"]
    LIB --> CM["Cuadro de Mando<br/>e indicadores"]
```
