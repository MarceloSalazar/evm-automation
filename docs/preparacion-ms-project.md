# Preparación del proyecto en MS Project

[Inicio](../README.md) › [Documentación](README.md) › Preparación en MS Project

Este documento reúne lo que se debe configurar en MS Project para que el programa obtenga el PV, el EV
y el AC correctos: la estructura y el código EDT, los campos personalizados, el presupuesto, la línea
base, el registro del avance y la exportación a XML. Los temas más extensos tienen su propio documento:
[Presupuesto (BAC)](presupuesto-bac.md) y [Métodos de medición](metodos-de-medicion.md).

> Los nombres de menú corresponden a MS Project en español 2016 a 2024 y a la versión de suscripción
> (*Project Online Desktop Client*); pueden variar levemente según la versión. Las imágenes son
> **esquemas ilustrativos**, no capturas.

---

## Contenido

1. [Lista de control](#1-lista-de-control)
2. [Estructura y código EDT](#2-estructura-y-código-edt)
3. [Campos personalizados](#3-campos-personalizados)
4. [Presupuesto de cada tarea](#4-presupuesto-de-cada-tarea)
5. [Método de medición e hitos](#5-método-de-medición-e-hitos)
6. [Línea base](#6-línea-base)
7. [Fecha de estado y registro del avance](#7-fecha-de-estado-y-registro-del-avance)
8. [Exportación a XML (sin COM)](#8-exportación-a-xml-sin-com)
9. [Proyectos grandes y complejos](#9-proyectos-grandes-y-complejos)

---

## 1. Lista de control

**Una vez, antes del primer `generar`:**

- [ ] Todas las tareas con costo tienen código EDT o dependen de una tarea con código EDT.
- [ ] Los paquetes de trabajo (nivel 3) agrupan las tareas de detalle.
- [ ] Los costos están en las tareas de **detalle**, no en las tareas resumen.
- [ ] Cada paquete tiene su **método de medición** (`Método EVM`) y, si usa hitos, sus hitos con
      `Peso hito`.
- [ ] Si hay dos planificaciones (`INTERNO`, `CLIENTE`), cada una está bajo una tarea de primer nivel con
      ese código EDT.
- [ ] Multi-empresa (opcional): columnas `Empresa` y `% Reparto` completas ([Multi-empresa](multiempresa.md)).
- [ ] **Línea base guardada** y resumida en las tareas resumen.
- [ ] `config.toml` con los nombres de los campos ([Configuración](configuracion.md)).

**Cada mes, antes de `actualizar`:**

- [ ] Avance registrado (*% completado*, *% físico completado*, fechas reales, hitos cumplidos).
- [ ] Trabajo pendiente reprogramado después de la fecha de estado.
- [ ] Costo real: calculado por MS Project o escrito desde el ERP (o cargado luego en Excel).
- [ ] **Fecha de estado** = último día del mes.
- [ ] Si se usa XML: archivo guardado **después** de fijar la fecha de estado.

---

## 2. Estructura y código EDT

El programa toma el **nivel** de cada tarea de su **código EDT**, no de su posición en el esquema:
`2.0.0` es nivel 1, `2.1.0` nivel 2 y `2.1.1` nivel 3 (los ceros finales no cuentan). Las tareas con
código hasta el nivel 3 tienen fila en el libro; las de nivel inferior y las que no tienen código se
suman en su elemento superior ([Cálculo del PV](calculo-pv.md#3-filas-del-libro-niveles-por-código-edt)).

El código EDT se puede leer de tres campos (clave `[campos] edt`):

| Opción | Cuándo usarla |
|---|---|
| Campo **EDT** de MS Project (`WBS`, por defecto) | Cuando el código EDT de MS Project ya sigue la numeración del proyecto. Se puede definir una máscara en *Proyecto > EDT > Definir código*. |
| **Número de esquema** (`OutlineNumber`) | Proyectos simples en los que el esquema coincide con la EDT. |
| **Campo personalizado de texto** (por ejemplo `Item EDT` en *Texto2*) | Cuando la EDT del contrato no coincide con el esquema (por ejemplo, con dos planificaciones o con tareas auxiliares intercaladas). Es la opción más flexible: el código se escribe solo en las tareas que deben tener fila. |

Recomendaciones:

- Los paquetes de trabajo (nivel 3) son la unidad de medición del EV: cada uno con un responsable, un
  presupuesto y un método.
- Las actividades de detalle y los hitos de un paquete **no necesitan** código EDT.
- Las tareas de gestión que no pertenecen a ningún elemento (por ejemplo, una reunión de arranque sin
  costo) pueden quedar sin código; si tienen costo, se informan en `Validación`.

---

## 3. Campos personalizados

El programa puede leer cuatro campos personalizados de tarea. Se crean con clic derecho en el
encabezado de una columna > **Campos personalizados** (o *Proyecto > Campos personalizados*):

| Campo sugerido | Tipo | Nombre (alias) | Valores | Clave de configuración |
|---|---|---|---|---|
| *Texto2* | Texto | `Item EDT` | Código EDT (`1.0.0`, `1.1.0`, `1.1.1`) | `[campos] edt` |
| *Texto10* | Texto | `Método EVM` | Lista de búsqueda con los métodos | `[campos] metodo` |
| *Número1* | Número | `Peso hito` | Peso de cada hito | `[campos] peso_hito` |
| *Texto3* | Texto | `Empresa` | Lista: `EmpresaA`, `EmpresaB`, `Ambas` | `[multiempresa] campo_empresa` |
| *Número2* | Número | `% Reparto` | % de la empresa A en tareas `Ambas` | `[multiempresa] campo_reparto` |

Pasos para cada campo:

1. **Campo: Tarea** y el **Tipo** (*Texto* o *Número*).
2. Se selecciona un campo libre y con **Cambiar nombre...** se le asigna el nombre.
3. Para listas cerradas (método, empresa): **Atributos personalizados > Búsqueda...**, se escriben los
   valores y se desmarca *Permitir introducir elementos adicionales en los campos*.
4. Se inserta la columna en la vista (clic derecho en un encabezado > *Insertar columna*).

![Esquema del cuadro Campos personalizados](img/ms-project/campos-personalizados.svg)

En la configuración, los campos se indican por su nombre interno en inglés (`Text10`, `Number1`), en
español (`Texto10`, `Número1`) o por su alias (`Método EVM`).

> Se recomienda guardar una vista y una tabla propias (por ejemplo, la tabla *EVM* con las columnas
> `Item EDT`, `Método EVM`, `Peso hito`, *% completado*, *% físico completado*, *Comienzo real* y *Fin
> real*) para el registro mensual del avance. Con el *Organizador* se pueden copiar a otros proyectos o
> a la plantilla global.

![Esquema de las columnas de medición en MS Project](img/ms-project/columnas-evm.svg)

---

## 4. Presupuesto de cada tarea

Los costos se asignan **solo a las tareas de detalle**, con cualquiera de estas opciones (se pueden
combinar):

- **Recursos de trabajo** con tarifa (*Vista > Hoja de recursos*, *Tasa estándar*).
- **Recursos materiales** con precio por unidad.
- **Recursos de costo** (montos por partida) asignados a la tarea, con el monto en *Información de la
  tarea > Recursos > Costo*.
- **Costo fijo** de la tarea (*Vista > Tablas > Costo*, columnas *Costo fijo* y *Acumulación de costos
  fijos*).

Cada opción, su configuración paso a paso y su efecto en la curva PV se explican en
[Presupuesto (BAC) de cada tarea](presupuesto-bac.md).

---

## 5. Método de medición e hitos

- Campo `Método EVM` con lista de búsqueda, completado en los paquetes de nivel 3.
- Para `Hitos ponderados`: hitos (duración 0) como subtareas del paquete, con su peso en `Peso hito`.
- Opcional: *Método de valor acumulado* de MS Project (*% completado* o *% físico completado*).

Detalle y ejemplos en [Métodos de medición](metodos-de-medicion.md).

---

## 6. Línea base

1. *Proyecto > Establecer línea base > Establecer línea base...*, **Línea base**, **Para: Proyecto
   completo**.
2. Tareas nuevas aprobadas: **Tareas seleccionadas** y **Resumir líneas base: Para todas las tareas de
   resumen**.
3. Copia de cada línea base aprobada en *Línea base 1..10* (*Establecer plan provisional*, *Copiar:
   Línea base*, *En: Línea base N*).

El *Costo previsto* (costo de línea base) de cada tarea es su **BAC**, y su distribución en el tiempo es
el **PV** que exporta el programa. Ver [Presupuesto: guardar la línea base](presupuesto-bac.md#6-guardar-la-línea-base).

---

## 7. Fecha de estado y registro del avance

### 7.1 Fecha de estado

- Antes de cada `actualizar` se define **Proyecto > Fecha de estado** (grupo *Estado*) = **último día
  del mes** informado. También en *Proyecto > Información del proyecto > Fecha de estado*.
- Sin fecha de estado, MS Project calcula el valor acumulado a la fecha de hoy y el programa exige
  `--fecha-corte`.
- Con archivos `.mpp`, si se indica `--fecha-corte`, el programa fija esa fecha de estado en la copia
  abierta (no guarda el archivo).

**Registro del avance** (con la fecha de estado ya definida):

| Dato | Dónde se registra | Para qué método |
|---|---|---|
| *% completado* | Botones *0%* a *100%* de la pestaña *Tarea*, tabla *Seguimiento* o *Proyecto > Actualizar proyecto* | `% Completado`, `0/100`, `X/Y`, hitos (100 % al cumplirse) |
| *% físico completado* | Columna *% físico completado* o tabla *Seguimiento* | `% Físico completado`, `Unidades completadas` |
| *Comienzo real* y *Fin real* | Tabla *Seguimiento* o cuadro *Actualizar tareas* | `X/Y` (comienzo), `0/100` (fin) |
| Trabajo real | Vista *Uso de tareas* o tabla *Trabajo* | Costo real de recursos de trabajo |

Buenas prácticas del registro:

- Registrar el avance **a la fecha de estado**, no a la fecha en que se hace la actualización.
- **Reprogramar el trabajo pendiente** a partir de la fecha de estado (*Proyecto > Actualizar proyecto >
  Reprogramar trabajo restante para que empiece después de*), para que las fechas futuras sigan siendo
  una previsión realista.
- No modificar la línea base al registrar el avance.

### 7.2 Costo real

- Por defecto, MS Project **calcula** el costo real con el trabajo real y las tarifas, más el costo fijo
  y los recursos de costo según su acumulación.
- Para escribir costos reales del ERP en MS Project se debe desmarcar *Archivo > Opciones >
  Programación > Los costos reales siempre se calculan por Project* (el texto exacto puede variar según
  la versión). Luego se escribe el *Costo real* de cada tarea o asignación (por ejemplo, en la vista
  *Uso de tareas*, con el costo real por período).
- Alternativa más simple: no registrar el costo real en MS Project, ejecutar `actualizar --solo ev` y
  cargar el AC del mes en la hoja `AC` del libro ([Cálculo de EV y AC](calculo-ev-ac.md#41-origen-del-costo-real)).

---

## 8. Exportación a XML (sin COM)

Si la PC que ejecuta el programa no tiene MS Project, se guarda el proyecto en formato XML:
*Archivo > Guardar como*, tipo **XML**. El XML se debe guardar **después** de fijar la fecha de estado:
el CPTR y el CRTR del archivo se calculan a esa fecha.

Limitaciones del XML: los tramos de costo que cruzan de un mes a otro se reparten con el calendario del
proyecto (no con los de recursos o tareas) ([Limitaciones](validaciones-y-problemas.md#2-limitaciones-conocidas)).

---

## 9. Proyectos grandes y complejos

- **Proyectos maestros con subproyectos:** el programa lee el proyecto maestro; el grupo de recursos se
  abre en solo lectura. Se debe verificar en `Validación` que el BAC total coincida con el del maestro.
- **Proyectos en Project Online, SharePoint u OneDrive:** se abren en MS Project y se usa
  `--origen activo`, o se guarda una copia local `.mpp`.
- **Varias planificaciones en un archivo** (`INTERNO` y `CLIENTE`): se genera un libro por
  planificación con `--planificacion` ([Cálculo del PV](calculo-pv.md#2-planificación-exportada)).
- **Tareas inactivas** (*Tarea > Desactivar*): no se incluyen, salvo `incluir_inactivas = true`. Es útil
  para conservar alcance descartado sin que afecte el BAC.
- **Calidad del cronograma:** un cronograma con lógica completa (predecesoras y sucesoras), sin
  restricciones innecesarias y con duraciones acotadas es la base de un PV confiable. Las guías de
  evaluación de cronogramas listadas en [Recursos externos](recursos-externos.md) (por ejemplo, la del
  GAO) describen los controles habituales.
