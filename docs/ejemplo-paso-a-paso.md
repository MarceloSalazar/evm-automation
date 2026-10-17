# Tutorial: ciclo completo con el proyecto de ejemplo

[Inicio](../README.md) › [Documentación](README.md) › Tutorial con el proyecto de ejemplo

El repositorio incluye un programa ficticio en formato XML de MS Project para probar el programa sin
MS Project. Este tutorial genera el libro, simula 24 meses de avance y recorre el resultado. Todas las
imágenes de la documentación se obtuvieron con estos pasos.

---

## Contenido

1. [El proyecto de ejemplo](#1-el-proyecto-de-ejemplo)
2. [Configuración](#2-configuración)
3. [Generar el libro](#3-generar-el-libro)
4. [Simular el avance y actualizar mes a mes](#4-simular-el-avance-y-actualizar-mes-a-mes)
5. [Revisar el resultado](#5-revisar-el-resultado)
6. [Pruebas adicionales](#6-pruebas-adicionales)

---

## 1. El proyecto de ejemplo

`ejemplos/programa_ejemplo.xml` es un programa ficticio de unos 5 años (desde enero de 2026) con dos
planificaciones, `INTERNO` y `CLIENTE` (esta con costos un 15 % mayores). Cada una tiene 4 elementos de
nivel 1:

| Código | Elemento | Paquetes y métodos de medición |
|---|---|---|
| `1.0.0` | Gestión del programa | PMO, control de proyectos y riesgos (`Avance lineal (LOE)`); licitaciones (`0/100`, `50/50`). |
| `2.0.0` | Ingeniería | Estudios (`50/50`), diseño conceptual y revisión (`Hitos ponderados`), ingeniería básica y de detalle (`% Completado`, `% Físico completado`). |
| `3.0.0` | Construcción | Obras civiles (`Unidades completadas`, `% Físico completado`), suministro de equipos (`Hitos ponderados`), montaje (`% Completado`). |
| `4.0.0` | Puesta en marcha | Pruebas (`0/100`, `20/80`), capacitación (`50/50`) y un paquete sin método (`Por definir`). |

- Los códigos EDT son del tipo `1.0.0`, `1.1.0`, `1.1.1` y `1.1.1.1`, y están también en el campo
  personalizado `Item EDT`, el que indica `config.ejemplo.toml`.
- Los hitos de los paquetes y la reunión de arranque no tienen código EDT.
- El método está en el campo `Método EVM` (*Texto10*) y el peso de los hitos en `Peso hito`
  (*Número1*).
- Los campos `Empresa` y `% Reparto` permiten probar el modo multi-empresa: gestión compartida 60/40,
  ingeniería de EmpresaA (salvo los estudios de terreno, de EmpresaB), construcción de EmpresaB y puesta
  en marcha compartida al 50 %.

`ejemplos/generar_ejemplo.py` vuelve a crear el XML y, con `--fecha-estado`, genera una versión con
**avance simulado** a esa fecha: un atraso del 15 % y un sobrecosto del 8 % en todas las tareas.

> En la simulación las tareas avanzan como máximo hasta el 85 % (nunca se marcan terminadas) y los
> hitos se cumplen 20 días después de su fecha prevista. Por eso los paquetes `0/100` y algunos de
> `Hitos ponderados` muestran costo real sin valor ganado: es el comportamiento esperado de esos métodos
> cuando un paquete no termina.

---

## 2. Configuración

```powershell
copy config.ejemplo.toml config.toml
```

En `config.toml` se definen los campos de método y de peso de los hitos:

```toml
[campos]
edt = "Item EDT"
metodo = "Método EVM"
peso_hito = "Peso hito"
```

El resto de los valores del archivo de ejemplo se deja igual (`planificacion = "INTERNO"`,
`mes_inicio = "2026-01"`).

---

## 3. Generar el libro

```powershell
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --config config.toml --salida ejemplo.xlsx
```

```text
Leyendo ejemplos\programa_ejemplo.xml ...
Proyecto: Programa de ejemplo EVM
  Planificación: INTERNO
  Línea base  |  PV: Costo de línea base mensual del XML de MS Project (datos escalonados)
  Filas EDT: 36 (24 de ingreso)  |  Nivel máximo: 3
  Horizonte: 01/2026 a 12/2035 (120 meses)
  BAC total: 44,210,000.00
  [i] Planificación 'INTERNO': se exportan las 51 tareas que dependen de 'Planificación interna'.
  [i] Se omitió 1 tarea sin código EDT que no depende de ningún elemento EDT (costo de línea base 0.00, que no forma parte del BAC).
  [i] 14 tareas sin código EDT o de nivel EDT mayor a 3 se agregaron en su elemento EDT superior.
  [i] 1 fila sin método de medición definido (columna C).
Libro generado: ejemplo.xlsx
```

Las observaciones son las esperadas: la reunión de arranque no tiene código EDT ni costo, las subtareas
y los hitos sin código se suman en su paquete, y `4.2.2 Cierre administrativo` no tiene método.

![Hoja PV del libro generado](img/excel/pv.png)

---

## 4. Simular el avance y actualizar mes a mes

Para un mes:

```powershell
python ejemplos\generar_ejemplo.py --fecha-estado 2026-01-31 --salida avance.xml
python evm_export.py actualizar --origen avance.xml --config config.toml --libro ejemplo.xlsx
```

Para simular 24 cierres mensuales seguidos (enero de 2026 a diciembre de 2027), en PowerShell:

```powershell
foreach ($m in 1..24) {
    $fecha = (Get-Date -Year 2026 -Month 1 -Day 1).AddMonths($m).AddDays(-1).ToString("yyyy-MM-dd")
    python ejemplos\generar_ejemplo.py --fecha-estado $fecha --salida avance.xml
    python evm_export.py actualizar --origen avance.xml --config config.toml --libro ejemplo.xlsx --sin-respaldo
}
```

Cada `actualizar` usa como fecha de corte la fecha de estado del XML (último día del mes), calcula el EV
y el AC acumulados y escribe la diferencia en el mes correspondiente. Al final de cada ejecución muestra
los indicadores; el último mes:

```text
Fecha de corte: 31/12/2027 (mes 24 del horizonte)
...
BAC (EUR)              44,210,000
PV acumulado             14,714,032   % planificado 33.3 %
EV acumulado             10,190,540   % completado  23.1 %
AC acumulado             13,491,360   % gastado     30.5 %
SV = EV - PV             -4,523,492   SPI = 0.69
CV = EV - AC             -3,300,820   CPI = 0.76
EAC = BAC/CPI            58,530,071   VAC = -14,320,071
TCPI                           1.11
ES (meses)                    21.67   AT = 24  SV(t) = -2.33  SPI(t) = 0.90
```

> `--sin-respaldo` evita crear 24 copias de respaldo durante la prueba. En el uso real conviene
> conservar el respaldo.

---

## 5. Revisar el resultado

Al abrir `ejemplo.xlsx` en Excel, el `Cuadro de Mando` muestra diciembre de 2027:

![Cuadro de Mando del ejemplo a diciembre de 2027](img/excel/cuadro-de-mando.png)

Puntos para observar (detalle en [Cuadro de Mando e indicadores](cuadro-de-mando.md)):

1. **Curva S:** EV por debajo de PV (atraso) y AC por encima de EV (sobrecosto).
2. **CPI y SPI:** en los primeros meses oscilan mucho porque los montos son pequeños; desde septiembre
   de 2026 se estabilizan cerca de 0,91 y 0,84, y caen al final de 2027 cuando comienzan la
   construcción y el suministro de equipos.
3. **CPI vs SPI por elemento:** `1.1.0 Dirección y control` (LOE) está en SPI = 1, porque el LOE no
   genera variación de cronograma; `3.2.0 Montaje electromecánico` está en el límite inferior porque el
   suministro de equipos aún no cumplió hitos.
4. **Indicadores EDT:** identifica los paquetes que explican la desviación.

![Hoja Indicadores EDT del ejemplo](img/excel/indicadores-edt.png)

---

## 6. Pruebas adicionales

**Ver un mes anterior.** En la celda amarilla `E5` del `Cuadro de Mando` se elige, por ejemplo,
junio de 2027.

**Cambiar el método EAC.** En la celda amarilla `R5` se elige `AC + (BAC - EV) / (CPI × SPI)`.

**Cambiar el método de un paquete.** En la hoja `PV` se cambia el método de `4.2.2 Cierre
administrativo` o de `1.2.1 Licitación de ingeniería` (por ejemplo, de `0/100` a `% Completado`), se
guarda el libro y se repite `actualizar` para el último mes: el valor de ese mes se reemplaza con el
nuevo cálculo.

```powershell
python ejemplos\generar_ejemplo.py --fecha-estado 2027-12-31 --salida avance.xml
python evm_export.py actualizar --origen avance.xml --config config.toml --libro ejemplo.xlsx --simular
```

**Regenerar con historial.** Simula un cambio de línea base conservando el EV y el AC cargados:

```powershell
python evm_export.py generar --origen ejemplos\programa_ejemplo.xml --config config.toml --salida ejemplo_lb2.xlsx --migrar-desde ejemplo.xlsx
```

**Planificación del cliente.** Con `--planificacion CLIENTE` se genera el libro de la otra
planificación (BAC un 15 % mayor).

**Multi-empresa.** En `config.toml` se indica `activar = true` (y, para cargar también el AC desde el
XML, `ac_desde_project = true`) en `[multiempresa]`, se genera un libro nuevo y se repiten las
actualizaciones. Ver [Multi-empresa](multiempresa.md).

**Dos archivos de MS Project.** El segundo proyecto de ejemplo (`ejemplos\proyecto2_ejemplo.xml`, código
EDT `9.0.0`) se combina con el programa de ejemplo en un mismo libro. Ver el ejemplo de
[Varios archivos](varios-archivos.md#9-ejemplo-con-dos-archivos).

Los archivos `config.toml`, `*.xlsx` y los XML de la carpeta raíz no se suben al repositorio (están en
`.gitignore`).
