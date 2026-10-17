# Archivo de configuración

[Inicio](../README.md) › [Documentación](README.md) › Configuración

La configuración se guarda en un archivo TOML. Se debe copiar `config.ejemplo.toml` como `config.toml`
y ajustarlo a cada proyecto; se indica con `--config config.toml`. **Todas las claves son opcionales**:
si se omite una clave, se usa el valor por defecto.

```powershell
copy config.ejemplo.toml config.toml
```

> `config.toml` no se sube al repositorio (está en `.gitignore`): puede contener nombres de campos o de
> empresas del proyecto real.

---

## Claves

| Sección / clave | Por defecto | Uso |
|---|---|---|
| `[campos] edt` | `"WBS"` | Campo con el código EDT: `WBS` (campo EDT de MS Project), `OutlineNumber` (número de esquema), un campo de texto (`Text1`...) o su alias. Si un campo personalizado tiene el alias `EDT`, `edt = "EDT"` usa ese campo. El archivo de ejemplo indica el campo personalizado `"Item EDT"`. |
| `[campos] metodo` | vacío | Campo personalizado con el método de medición, p. ej. `"Text10"` o su alias `"Método EVM"`. |
| `[campos] peso_hito` | vacío | Campo numérico con el peso de cada hito, p. ej. `"Number1"` o `"Peso hito"`. |
| `[horizonte] mes_inicio` | vacío | Primer mes (`"AAAA-MM"`). Vacío = inicio de la línea base. El archivo de ejemplo indica `"2026-01"`. |
| `[horizonte] meses` | `120` | Cantidad de meses. |
| `[exportacion] planificacion` | sin filtro | Código EDT de alto nivel de la planificación a exportar, por ejemplo `"INTERNO"` o `"CLIENTE"`; el archivo de ejemplo indica `"INTERNO"`. Si se omite, se exportan todas las tareas. |
| `[exportacion] nivel_maximo` | `3` | Nivel máximo del código EDT exportado; los niveles inferiores se suman en su elemento de este nivel. |
| `[exportacion] linea_base` | la del proyecto | `0` = Línea base, `1`..`10` = Línea base N. |
| `[exportacion] rellenar_edt` | `true` | `1` → `1.0.0`, `1.2` → `1.2.0`. |
| `[exportacion] incluir_inactivas` | `false` | Incluir tareas inactivas. |
| `[exportacion] distribucion` | `"habiles"` | Para tareas sin distribución mensual en MS Project: reparto lineal por tiempo laborable (`habiles`) o por días corridos (`calendario`). |
| `[exportacion] feriados` | `[]` | Feriados adicionales. Con XML se agregan al calendario del proyecto y afectan todos los repartos por tiempo laborable; con `.mpp` solo el reparto lineal (MS Project usa sus propios calendarios). |
| `[exportacion] metodo_por_defecto` | `"Por definir"` | Método para las filas sin método. |
| `[exportacion] moneda` | `"EUR"` | Moneda que se muestra en el libro. Con `""` se usa el código de moneda del proyecto en MS Project. |
| `[exportacion] decimales` | `0` | Decimales de los montos en el libro (0 a 4). |
| `[metodos] lista` | ver ejemplo | Opciones de la lista desplegable de la columna C. |
| `[semaforo] rojo` / `amarillo` | `0.90` / `0.95` | Umbrales de SPI y CPI: rojo si es menor que `rojo`, amarillo si es menor que `amarillo`, verde en otro caso. |
| `[multiempresa] activar` | `false` | Activa el modo multi-empresa ([Multi-empresa](multiempresa.md)). |
| `[multiempresa] campo_empresa` | `"Empresa"` | Campo de texto de MS Project con la empresa de cada tarea (por ejemplo `"Text3"`, o su alias). |
| `[multiempresa] empresa_a` / `empresa_b` | `"EmpresaA"` / `"EmpresaB"` | Valores de la columna de empresa y sufijo de las hojas de cada empresa (`PV-EmpresaA`...). Máximo 15 caracteres, sin `[ ] : * ? / \ '`. |
| `[multiempresa] ambas` | `"Ambas"` | Valor que indica una tarea de las dos empresas. |
| `[multiempresa] campo_reparto` | `"% Reparto"` | Campo numérico con el % de la empresa A en las tareas `Ambas`. |
| `[multiempresa] reparto_por_defecto` | `50` | % de la empresa A cuando el campo de reparto está vacío, es 0 o está fuera de rango, y en las tareas sin empresa. |
| `[multiempresa] ac_desde_project` | `false` | `false`: el AC de cada empresa se carga a mano. `true`: `actualizar` también lo carga desde MS Project. |

Los nombres de campo aceptan el nombre interno en inglés (`Text10`, `Number1`), el nombre en español
(`Texto10`, `Número1`) o el alias definido en MS Project. Solo se buscan campos de tarea: un campo de
recurso con el mismo alias no se toma. Cómo crear esos campos se explica en
[Preparación del proyecto en MS Project](preparacion-ms-project.md#3-campos-personalizados).

Las opciones de la línea de comandos (`--planificacion`, `--linea-base`, `--nivel-max`, `--mes-inicio`,
`--meses`) reemplazan el valor del archivo de configuración.

El archivo se debe guardar con codificación **UTF-8**, con o sin BOM (en el Bloc de notas: *Guardar
como*, *Codificación: UTF-8*). Los valores enteros (`meses`, `nivel_maximo`, `decimales`, `linea_base`)
no admiten decimales.

---

## Ejemplos

**Configuración mínima** para un proyecto con el código EDT en el campo *EDT* de MS Project y los
campos de método y peso de hito con alias:

```toml
[campos]
edt = "WBS"
metodo = "Método EVM"
peso_hito = "Peso hito"
```

**Programa con dos planificaciones** y el código EDT en un campo personalizado:

```toml
[campos]
edt = "Item EDT"
metodo = "Método EVM"
peso_hito = "Peso hito"

[horizonte]
mes_inicio = "2026-01"
meses = 120

[exportacion]
planificacion = "INTERNO"
moneda = "EUR"
```

**Multi-empresa** con el AC cargado a mano en Excel:

```toml
[multiempresa]
activar = true
campo_empresa = "Empresa"
empresa_a = "Constructora"
empresa_b = "Ingenieria"
ambas = "Ambas"
campo_reparto = "% Reparto"
reparto_por_defecto = 50
ac_desde_project = false
```

**Semáforo más exigente** (rojo por debajo de 0,95 y amarillo por debajo de 0,98):

```toml
[semaforo]
rojo = 0.95
amarillo = 0.98
```
