# Comandos y opciones

[Inicio](../README.md) › [Documentación](README.md) › Comandos y opciones

El programa tiene dos comandos: `generar` (crea el libro con el PV) y `actualizar` (carga el EV y el AC
de un mes). La ayuda se obtiene con:

```powershell
python evm_export.py --help
python evm_export.py generar --help
python evm_export.py actualizar --help
```

Para ver mensajes de diagnóstico se agrega `-v` **antes** del comando:
`python evm_export.py -v generar ...`

---

## `generar`: crea el libro con el PV

```powershell
python evm_export.py generar --origen "C:\Programa\Programa.mpp" --config config.toml --salida EVM_Programa.xlsx
```

| Opción | Descripción |
|---|---|
| `--origen` | Archivo `.mpp` (requiere MS Project), `.xml` de MS Project, o `activo` para usar el proyecto abierto en MS Project. **Obligatoria.** |
| `--salida` | Libro `.xlsx` a crear. Por defecto `<origen>_EVM.xlsx`. |
| `--config` | Archivo de configuración TOML ([Configuración](configuracion.md)). |
| `--mes-inicio AAAA-MM` | Primer mes del horizonte (mes 1). Por defecto, el mes en que comienza la línea base. |
| `--meses N` | Cantidad de meses (columnas). Por defecto 120. |
| `--planificacion CODIGO` | Código EDT de la planificación a exportar (`INTERNO`, `CLIENTE`...). Reemplaza el valor del archivo de configuración; `""` exporta todas. |
| `--linea-base N` | 0 = Línea base, 1..10 = Línea base N. Por defecto, la que el proyecto usa para el valor acumulado. |
| `--nivel-max N` | Nivel máximo del código EDT. Por defecto 3. |
| `--motor auto\|com\|xml` | Forma de lectura. `auto` decide por la extensión. |
| `--migrar-desde libro.xlsx` | Copia al libro nuevo los EV/AC ya cargados, los métodos editados en la columna C y la fecha de corte de un libro anterior ([El libro Excel](libro-excel.md#4-cambios-de-línea-base-o-de-edt---migrar-desde)). |
| `--sobrescribir` | Reemplaza el libro de salida si ya existe. |
| `--visible` | Muestra MS Project mientras se lee (solo COM). |

Ejemplo de salida con el proyecto de ejemplo:

```text
Leyendo ejemplos/programa_ejemplo.xml ...
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
Libro generado: ...\EVM_Ejemplo.xlsx
```

---

## `actualizar`: carga EV y AC del mes

```powershell
python evm_export.py actualizar --origen "C:\Programa\Programa.mpp" --config config.toml --libro EVM_Programa.xlsx --fecha-corte 2026-10-31
```

| Opción | Descripción |
|---|---|
| `--origen` | Proyecto con el avance a la fecha de corte (`.mpp`, `.xml` o `activo`). **Obligatoria.** |
| `--libro` | Libro creado con `generar`. **Obligatoria.** |
| `--fecha-corte AAAA-MM-DD` | Fecha de corte. Por defecto, la fecha de estado del proyecto; si el proyecto no tiene fecha de estado, es obligatoria. Con `.mpp`, el programa fija esa fecha de estado en la copia abierta (no guarda el archivo). |
| `--solo ev\|ac` | Actualiza solo EV o solo AC. En multi-empresa, `--solo ac` requiere `ac_desde_project = true` ([Multi-empresa](multiempresa.md)). |
| `--simular` | Muestra el EV y AC calculados para cada fila y los indicadores resultantes, sin modificar el libro. |
| `--forzar` | Permite actualizar aunque haya valores en meses posteriores. |
| `--sin-respaldo` | No crea la copia `<libro>.respaldo-AAAAMMDD-HHMMSS.xlsx` (si ya existe una con ese nombre, se agrega `-2`, `-3`...). |
| `--config`, `--motor`, `--linea-base`, `--nivel-max`, `--visible` | Igual que en `generar`. |
| `--planificacion` | No se aplica: `actualizar` usa siempre la planificación con la que se generó el libro (y lo advierte si la configuración indica otra). |

- Repetir `actualizar` para el mismo mes **reemplaza** el valor de ese mes; no lo duplica.
- `actualizar` controla la consistencia con MS Project (tareas eliminadas, costos sin fila, diferencia
  de AC con el CRTR del proyecto, celdas con texto o fórmulas) y lo informa en pantalla
  ([Cálculo de EV y AC](calculo-ev-ac.md#7-controles-de-consistencia)).
- Al terminar, muestra los indicadores a la fecha de corte, con los mismos cálculos que la hoja
  `Resumen`.

Ejemplo de salida:

```text
Leyendo ...\avance_2027-06.xml ...
Fecha de corte: 30/06/2027 (mes 18 del horizonte)
Observaciones por fila:
  4.2.2 Cierre administrativo: método no definido: se usó % completado

BAC (EUR)              44,210,000
PV acumulado              5,895,918   % planificado 13.3 %
EV acumulado              4,946,380   % completado  11.2 %
AC acumulado              5,393,844   % gastado     12.2 %
SV = EV - PV               -949,538   SPI = 0.84
CV = EV - AC               -447,464   CPI = 0.92
EAC = BAC/CPI            48,209,366   VAC = -3,999,366
TCPI                           1.01
ES (meses)                    15.98   AT = 18  SV(t) = -2.02  SPI(t) = 0.89
Libro actualizado: ...\EVM_Ejemplo.xlsx
Los indicadores y los gráficos se calculan en el libro (hojas 'Cuadro de Mando', 'Resumen' e 'Indicadores EDT').
```

---

## Ejemplos frecuentes

```powershell
# Proyecto abierto en MS Project (por ejemplo, desde Project Online o SharePoint)
python evm_export.py generar --origen activo --config config.toml --salida EVM_Programa.xlsx

# Exportar la planificación del cliente en lugar de la indicada en la configuración
python evm_export.py generar --origen Programa.mpp --config config.toml --planificacion CLIENTE --salida EVM_Cliente.xlsx

# Revisar el cálculo antes de modificar el libro
python evm_export.py actualizar --origen Programa.mpp --config config.toml --libro EVM_Programa.xlsx --simular

# AC del ERP cargado a mano: actualizar solo el EV
python evm_export.py actualizar --origen Programa.mpp --config config.toml --libro EVM_Programa.xlsx --solo ev

# Nueva línea base: libro nuevo conservando el historial
python evm_export.py generar --origen Programa.mpp --config config.toml --salida EVM_Programa_LB2.xlsx --migrar-desde EVM_Programa.xlsx
```

Si no se activó el entorno virtual, se usa `.venv\Scripts\python.exe` en lugar de `python`.
