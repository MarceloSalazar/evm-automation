# Multi-empresa (opcional)

[Inicio](../README.md) › [Documentación](README.md) › Multi-empresa

Cuando el proyecto se ejecuta entre **dos empresas** (por ejemplo, un consorcio, una unión temporal o
una empresa y su socio tecnológico), el libro puede separar el PV, el EV y el AC de cada una, además
del total del proyecto. Este modo está **desactivado por defecto**.

---

## Contenido

1. [Cuándo usarlo](#1-cuándo-usarlo)
2. [Cómo funciona](#2-cómo-funciona)
3. [Preparación en MS Project](#3-preparación-en-ms-project)
4. [Configuración](#4-configuración)
5. [Parte de cada empresa en cada tarea](#5-parte-de-cada-empresa-en-cada-tarea)
6. [Cálculo del EV y el AC por empresa](#6-cálculo-del-ev-y-el-ac-por-empresa)
7. [Validaciones y controles](#7-validaciones-y-controles)
8. [Hojas del libro](#8-hojas-del-libro)
9. [Ciclo mensual con dos empresas](#9-ciclo-mensual-con-dos-empresas)
10. [Preguntas frecuentes](#10-preguntas-frecuentes)

---

## 1. Cuándo usarlo

- Cada empresa es responsable de una parte del alcance y se necesita conocer el **desempeño de cada
  una** (CPI y SPI por empresa) además del total.
- Hay paquetes compartidos cuyo presupuesto se reparte en un porcentaje acordado (por ejemplo, la
  gestión del programa 60/40).
- Cada empresa registra sus **costos reales en su propio sistema** (ERP) y el AC del proyecto es la suma
  de ambos.

Si el proyecto lo ejecuta una sola empresa, este modo no es necesario.

---

## 2. Cómo funciona

```mermaid
flowchart LR
    subgraph MSP["MS Project"]
        T["Tareas con columna Empresa<br/>(EmpresaA, EmpresaB o Ambas)<br/>y % Reparto"]
    end
    subgraph LIBRO["Libro Excel"]
        PVA["PV-EmpresaA"] --> PV["PV = A + B"]
        PVB["PV-EmpresaB"] --> PV
        EVA["EV-EmpresaA"] --> EV["EV = A + B"]
        EVB["EV-EmpresaB"] --> EV
        ACA["AC-EmpresaA"] --> AC["AC = A + B"]
        ACB["AC-EmpresaB"] --> AC
        PV & EV & AC --> CM["Cuadro de Mando,<br/>Resumen e Indicadores EDT<br/>(totales)"]
    end
    T -- "generar: PV y BAC<br/>repartidos por tarea" --> PVA & PVB
    T -- "actualizar: EV<br/>de cada empresa" --> EVA & EVB
    ERP["ERP de cada empresa"] -- "AC a mano<br/>(por defecto)" --> ACA & ACB
```

- El **PV y el BAC** se reparten **tarea por tarea** según la parte de cada empresa. Una fila de nivel 3
  puede incluir tareas de las dos empresas.
- En las hojas `PV`, `EV` y `AC` cada celda es la **suma** de las dos empresas (fórmula).
- **Cuadro de Mando, Resumen e Indicadores EDT** usan los totales.

---

## 3. Preparación en MS Project

1. Se crea un campo personalizado de tarea de tipo *Texto* (por ejemplo *Texto3*) con el nombre
   `Empresa` y, en *Atributos personalizados > Búsqueda*, los valores `EmpresaA`, `EmpresaB` y `Ambas`
   ([Campos personalizados](preparacion-ms-project.md#3-campos-personalizados)). La tabla de búsqueda
   evita errores de escritura; los valores deben coincidir con `empresa_a`, `empresa_b` y `ambas` de la
   configuración (sin distinguir mayúsculas ni tildes).
2. Se crea un campo de tipo *Número* (por ejemplo *Número2*) con el nombre `% Reparto`: porcentaje de la
   empresa A en las tareas `Ambas` (por ejemplo, 80 = 80 % para la empresa A y 20 % para la empresa B).
3. Se insertan las dos columnas en la vista *Diagrama de Gantt* y se completa la columna `Empresa`. Una
   tarea sin empresa **hereda** la de su tarea superior (con su `% Reparto`): basta con completarla en
   las tareas resumen y en las subtareas que sean de otra empresa.

![Esquema de las columnas Empresa y % Reparto en MS Project](img/ms-project/columnas-multiempresa.svg)

*Esquema ilustrativo (no es una captura de MS Project): la empresa se indica en las tareas resumen y
solo se repite en las subtareas que son de otra empresa.*

> **Recomendación:** acordar el reparto de los paquetes `Ambas` antes de guardar la línea base y
> documentarlo (por ejemplo, en el acuerdo de consorcio). Cambiar la empresa o el `% Reparto` de una
> tarea después de guardar la línea base cambia el BAC de cada empresa en el siguiente `generar`.

---

## 4. Configuración

```toml
[multiempresa]
activar = true
campo_empresa = "Empresa"        # campo de texto (Text3) o su alias
empresa_a = "EmpresaA"           # valor de la columna y sufijo de las hojas
empresa_b = "EmpresaB"
ambas = "Ambas"
campo_reparto = "% Reparto"      # campo numérico (Number2) o su alias
reparto_por_defecto = 50         # % de A si % Reparto está vacío, es 0 o está fuera de rango
ac_desde_project = false         # false: AC a mano; true: AC desde MS Project
```

Los nombres de las empresas se usan en los nombres de las hojas: máximo 15 caracteres, sin
`[ ] : * ? / \ '`. Ver todas las claves en [Configuración](configuracion.md).

`actualizar` usa los nombres de las empresas y los campos con los que se **generó** el libro (guardados
en la hoja oculta `_meta`), aunque la configuración no tenga la sección `[multiempresa]`. La opción
`ac_desde_project` se toma de la configuración actual.

---

## 5. Parte de cada empresa en cada tarea

| Columna `Empresa` | Parte de la empresa A (la empresa B recibe el resto) |
|---|---|
| Empresa A | 100 % |
| Empresa B | 0 % |
| `Ambas` | El valor de `% Reparto`. Si está vacío o es 0, el reparto por defecto (`reparto_por_defecto`, 50 %). Si es negativo o mayor que 100, el reparto por defecto, y se informa en `Validación`. |
| Vacía | La de la tarea superior. Si ninguna tarea superior tiene empresa, se trata como `Ambas` con el reparto por defecto y se informa en `Validación`. |
| Otro valor | Se trata como `Ambas` con el reparto por defecto y se informa en `Validación`. |

- `% Reparto` solo se usa en las tareas `Ambas`; en las demás se ignora.
- En MS Project un campo numérico vacío vale 0, por eso 0 equivale a vacío. Para asignar el 100 % a una
  empresa se indica su nombre en `Empresa`.
- Una subtarea marcada `Ambas` sin `% Reparto` usa el reparto por defecto, **no** el de su tarea
  superior. Para heredar el reparto de la tarea superior, la columna `Empresa` de la subtarea se deja
  vacía.
- La herencia también llega desde las tareas que están por encima de la planificación exportada (por
  ejemplo, si la tarea `INTERNO` tiene empresa).

**Ejemplo de herencia** (proyecto de ejemplo):

```mermaid
flowchart TD
    G["1.0.0 Gestión del programa<br/>Empresa: Ambas · % Reparto: 60"] --> G1["1.1.1 Oficina de gestión (PMO)<br/>(vacía) → Ambas 60/40"]
    I["2.0.0 Ingeniería<br/>Empresa: EmpresaA"] --> I1["2.1.1 Estudios de terreno<br/>Empresa: EmpresaB → 0/100"]
    I --> I2["2.1.3 Ingeniería básica civil<br/>(vacía) → EmpresaA 100/0"]
    C["3.0.0 Construcción<br/>Empresa: EmpresaB"] --> C1["3.1.1 Movimiento de tierras<br/>(vacía) → EmpresaB 0/100"]
    P["4.0.0 Puesta en marcha<br/>Empresa: Ambas · % Reparto vacío"] --> P1["4.1.1 Pruebas en frío<br/>(vacía) → Ambas 50/50 por defecto"]
```

*Entre paréntesis, el valor de la columna `Empresa` de la tarea; después de la flecha, el reparto A/B
resultante.*

---

## 6. Cálculo del EV y el AC por empresa

### 6.1 PV y BAC

Cada tarea aporta al PV y al BAC de cada empresa su costo de línea base multiplicado por la parte de
esa empresa. El BAC de una fila para la empresa A es:

```text
BAC_A de la fila = Σ (costo de línea base de cada tarea de la fila × parte de A en esa tarea)
BAC_B de la fila = BAC de la fila − BAC_A
```

### 6.2 EV

`actualizar` escribe el EV del mes en `EV-EmpresaA` y `EV-EmpresaB`. El EV de cada empresa se calcula
con **su BAC en la fila** y **su propio avance**:

| Método (columna C) | EV acumulado de cada empresa |
|---|---|
| `% Completado`, `% Físico completado`, `Unidades completadas` | BAC de la empresa × avance de sus tareas, ponderado por el costo de línea base que le corresponde en cada tarea. |
| `CPTR de MS Project` | CPTR de las tareas de la fila, repartido según la parte de cada empresa en cada tarea. |
| `Avance lineal (LOE)` | PV acumulado de la empresa. |
| `0/100`, `X/Y`, `Hitos ponderados` | BAC de la empresa × avance del paquete completo. |

### 6.3 AC

| `ac_desde_project` | AC de cada empresa |
|---|---|
| `false` (por defecto) | Se carga **a mano** en `AC-EmpresaA` y `AC-EmpresaB` (celdas amarillas). `actualizar` no lo modifica y no admite `--solo ac`. |
| `true` | `actualizar` reparte el CRTR de cada tarea según su parte y lo escribe en las hojas de cada empresa. |

### 6.4 Ejemplo

Paquete `3.1.1`, BAC 1.000.000 €, con dos tareas:

| Tarea | Empresa | Costo previsto | Parte de A | BAC A | BAC B | % completado | CRTR |
|---|---|---:|---:|---:|---:|---:|---:|
| Excavación | EmpresaA | 600.000 | 100 % | 600.000 | 0 | 40 % | 250.000 |
| Relleno compactado | Ambas, 70 | 400.000 | 70 % | 280.000 | 120.000 | 50 % | 210.000 |
| **Fila 3.1.1** | | **1.000.000** | | **880.000** | **120.000** | | **460.000** |

| Método | EV empresa A | EV empresa B | EV total |
|---|---:|---:|---:|
| `% Completado` | 880.000 × (600.000 × 40 % + 280.000 × 50 %) / 880.000 = **380.000** | 120.000 × 50 % = **60.000** | 440.000 |
| `50/50` (paquete comenzado) | 880.000 × 50 % = **440.000** | 120.000 × 50 % = **60.000** | 500.000 |

AC con `ac_desde_project = true`: empresa A = 250.000 + 70 % × 210.000 = **397.000**; empresa B =
30 % × 210.000 = **63.000**.

---

## 7. Validaciones y controles

`generar` informa en `Validación`:

- El **BAC de cada empresa**.
- Tareas **sin empresa** (ni propia ni heredada): se reparten con el reparto por defecto.
- Valores de empresa **no reconocidos** (por ejemplo, un error de escritura).
- `% Reparto` **fuera de rango** (negativo o mayor que 100).

![Hoja Validación en modo multi-empresa](img/excel/multi-validacion.png)

Restricciones:

- `--solo ac` requiere `ac_desde_project = true`.
- `--migrar-desde` traslada el historial de cada empresa entre libros multi-empresa. No se admite entre
  un libro de una sola empresa y uno multi-empresa, porque el historial no se puede repartir entre las
  empresas.

---

## 8. Hojas del libro

Además de `PV`, `EV` y `AC`, se crean `PV-EmpresaA`, `PV-EmpresaB`, `EV-EmpresaA`, `EV-EmpresaB`,
`AC-EmpresaA` y `AC-EmpresaB` (con los nombres configurados), con el mismo formato.

![Hoja PV-EmpresaA](img/excel/multi-pv-empresa-a.png)

*`PV-EmpresaA` del ejemplo: la gestión aporta el 60 %, la ingeniería el 100 % (salvo los estudios de
terreno) y la construcción nada.*

![Hoja EV-EmpresaB](img/excel/multi-ev-empresa-b.png)

*`EV-EmpresaB` (meses 13 a 24): las celdas amarillas son el EV del mes de la empresa B.*

---

## 9. Ciclo mensual con dos empresas

```mermaid
sequenceDiagram
    participant A as Empresa A
    participant B as Empresa B
    participant JP as Jefe de Proyecto
    participant MSP as MS Project
    participant XL as Libro EVM
    A->>JP: Avance físico y AC devengado del mes
    B->>JP: Avance físico y AC devengado del mes
    JP->>MSP: Registra el avance y fija la fecha de estado
    JP->>XL: actualizar (EV de cada empresa)
    JP->>XL: Carga el AC de cada empresa (AC-EmpresaA, AC-EmpresaB)
    XL-->>JP: Cuadro de Mando y Resumen (totales)
    JP->>A: Informe mensual
    JP->>B: Informe mensual
```

Se recomienda acordar con ambas empresas la **misma fecha de corte** y el mismo criterio de costo real
(devengado), y que cada una revise su EV antes de emitir el informe.

---

## 10. Preguntas frecuentes

**¿Se pueden ver los indicadores de cada empresa por separado?** El `Cuadro de Mando`, `Resumen` e
`Indicadores EDT` muestran los totales. Los valores de cada empresa están en sus hojas `PV-`, `EV-` y
`AC-`; el CPI y el SPI de cada empresa se pueden calcular a partir de la fila TOTAL de esas hojas.

**¿Se admiten más de dos empresas?** No. Con más socios, se puede agrupar a los socios menores en una
de las dos empresas.

**¿Qué pasa si cambio la empresa de una tarea después de generar el libro?** El PV del libro no cambia
hasta regenerarlo. `actualizar` reparte el EV con la empresa actual de cada tarea, por lo que se
recomienda regenerar el libro con `--migrar-desde` si el cambio es definitivo.
