# Códigos de errores y advertencias

[Inicio](../README.md) › [Documentación](README.md) › Códigos de errores y advertencias

Cada mensaje que muestra el programa, en la consola o en la hoja `Validación` del libro, lleva un
**código** como `FEC-003`. Este documento explica cada código: qué lo causa y qué hacer. También explica
cómo **silenciar** las observaciones que ya se revisaron y no se quieren volver a ver.

> Este documento se genera a partir del catálogo del programa (`evm/messages.py`) con
> `python docs\herramientas\codigos.py`: no se edita a mano. Para agregar un código, ver
> [Desarrollo](desarrollo.md#5-códigos-de-errores-y-advertencias).

---

## Contenido

1. [Cómo se lee un código](#1-cómo-se-lee-un-código)
2. [Silenciar observaciones](#2-silenciar-observaciones)
3. [Áreas](#3-áreas)
4. [Códigos](#4-códigos)
   1. [Configuración y opciones](#41-cfg--configuración-y-opciones)
   2. [Lectura de MS Project](#42-lec--lectura-de-ms-project)
   3. [Estructura y códigos EDT](#43-edt--estructura-y-códigos-edt)
   4. [Campo Visible](#44-vis--campo-visible)
   5. [Fechas y duración](#45-fec--fechas-y-duración)
   6. [Hitos y pesos](#46-hit--hitos-y-pesos)
   7. [Método de medición](#47-met--método-de-medición)
   8. [Empresas](#48-emp--empresas)
   9. [Avance a la fecha de estado](#49-ava--avance-a-la-fecha-de-estado)
   10. [Lógica del cronograma](#410-log--lógica-del-cronograma)
   11. [Libro](#411-lib--libro)
   12. [Migración (--migrar-desde)](#412-mig--migración---migrar-desde)
   13. [Depuración de costes](#413-dep--depuración-de-costes)
   14. [Actualizar y analizar](#414-act--actualizar-y-analizar)
   15. [Fichas de tareas](#415-fic--fichas-de-tareas)
   16. [Errores del sistema](#416-sis--errores-del-sistema)

---

## 1. Cómo se lee un código

El código tiene la forma **ÁREA-NNN**: tres letras del área (por ejemplo `FEC`, *Fechas y duración*) y un
número. El código de un mensaje no cambia aunque cambien su texto o su severidad, y un código que deja de
usarse no se reutiliza: queda en este documento como *retirado*.

| Severidad | En la consola | Qué indica |
|---|---|---|
| Error (detiene el programa) | `ERROR CFG-003: ...` (salida de error) | El comando no puede seguir: termina sin crear ni modificar archivos, con código de salida 2 (3 si es un error inesperado, `SIS-001`). |
| Error | `[X] EDT-009 ...` | Dato incorrecto que se debe corregir, o dato que no se pudo trasladar o calcular. `validar` termina con código de salida 1. |
| Advertencia | `[!] FEC-007 ...` | Se debe revisar: puede afectar el resultado. |
| Info | `[i] EDT-016 ...` | Información sobre lo que hizo el programa. |

Algunos códigos tienen una severidad u otra según el caso (por ejemplo, *Advertencia o Info*); el
documento lo indica en cada uno.

En la **consola**, cada observación muestra su código, el código EDT y el nombre de la fila o tarea (si
corresponde) y el mensaje:

```text
  [!] FEC-007 1.2.3 Montaje: Duración estimada («?»).
  [i] EDT-016 Se excluyeron 2 tareas inactivas: no forman parte del % planificado.
```

En la **hoja `Validación`** del libro, el código está en la columna `Código`, junto a la severidad; con el
filtro de la columna se pueden ver las observaciones de un código.

---

## 2. Silenciar observaciones

Las observaciones de un código se pueden **silenciar**: no se muestran en la consola ni en la hoja
`Validación`. Sirve para no volver a ver lo que ya se revisó y se acepta (por ejemplo, un control DCMA que
no aplica al proyecto). Se indica en el archivo de configuración, en la sección `[validacion]`:

```toml
[validacion]
silenciar = ["LOG-004", "MET-004:1.2.1", "FEC-007"]
```

o con la opción `--silenciar` de cualquier comando (se puede repetir, y los códigos se separan con comas);
se suma a los de la configuración:

```powershell
python evm_export.py validar --origen "C:\Programa\Programa.mpp" --silenciar LOG-004,FEC-007
```

| Forma | Qué silencia |
|---|---|
| `LOG-004` | Todas las observaciones de ese código. |
| `MET-004:1.2.1` | Solo las de la fila o tarea con ese código EDT (los ceros finales no cuentan: `1.2.1` es `1.2.1.0`). Las demás filas siguen mostrando el aviso. |

- Siempre se informa cuántas observaciones se silenciaron y de qué códigos (`CFG-049`), para que no
  desaparezcan sin dejar rastro.
- Los errores silenciados no cuentan para el código de salida de `validar`.
- Los errores que detienen el programa no se pueden silenciar (`CFG-048`).
- Un código que no existe en el catálogo se informa con `CFG-047` y se ignora.
- Un código retirado se acepta y no silencia nada: las configuraciones que lo usan siguen siendo válidas.

---

## 3. Áreas

| Área | Nombre | Tema en la hoja Validación | Dónde aparecen | Códigos |
|---|---|---|---|---|
| [CFG](#41-cfg--configuración-y-opciones) | Configuración y opciones | Lectura y configuración | Todos los comandos | 51 |
| [LEC](#42-lec--lectura-de-ms-project) | Lectura de MS Project | Lectura y configuración | generar, actualizar, validar y generar-fichas (lectura de MS Project) | 22 |
| [EDT](#43-edt--estructura-y-códigos-edt) | Estructura y códigos EDT | Estructura y códigos EDT | generar (hoja Validación) y validar | 22 |
| [VIS](#44-vis--campo-visible) | Campo Visible | Campo Visible | generar (hoja Validación) y validar | 10 |
| [FEC](#45-fec--fechas-y-duración) | Fechas y duración | Fechas y duración | generar (hoja Validación) y validar | 12 |
| [HIT](#46-hit--hitos-y-pesos) | Hitos y pesos | Hitos y pesos | generar (hoja Validación) y validar | 9 |
| [MET](#47-met--método-de-medición) | Método de medición | Método de medición | generar (hoja Validación) y validar | 5 |
| [EMP](#48-emp--empresas) | Empresas | Empresas | generar (hoja Validación) y validar | 6 |
| [AVA](#49-ava--avance-a-la-fecha-de-estado) | Avance a la fecha de estado | Avance a la fecha de estado | validar | 17 |
| [LOG](#410-log--lógica-del-cronograma) | Lógica del cronograma | Lógica del cronograma | validar | 16 |
| [LIB](#411-lib--libro) | Libro | Libro | actualizar, analizar y generar `--migrar-desde` (lectura del libro) | 15 |
| [MIG](#412-mig--migración---migrar-desde) | Migración (--migrar-desde) | Libro | generar `--migrar-desde` (hoja Validación) | 19 |
| [DEP](#413-dep--depuración-de-costes) | Depuración de costes | Libro | generar con depuracion_costes (hoja Validación) | 8 |
| [ACT](#414-act--actualizar-y-analizar) | Actualizar y analizar | Libro | actualizar | 49 |
| [FIC](#415-fic--fichas-de-tareas) | Fichas de tareas | Libro | generar-fichas | 4 |
| [SIS](#416-sis--errores-del-sistema) | Errores del sistema | Libro | Todos los comandos | 3 |

---

## 4. Códigos

### 4.1 CFG · Configuración y opciones

Archivo de configuración, opciones de la línea de comandos, horizonte y códigos silenciados.

| Código | Severidad | Descripción |
|---|---|---|
| [CFG-001](#cfg-001) | Error (detiene el programa) | No se encontró el archivo de configuración |
| [CFG-002](#cfg-002) | Error (detiene el programa) | El archivo de configuración no está en UTF-8 |
| [CFG-003](#cfg-003) | Error (detiene el programa) | Error de sintaxis en el archivo de configuración |
| [CFG-004](#cfg-004) | Error (detiene el programa) | Clave con un valor del tipo incorrecto |
| [CFG-005](#cfg-005) | Error (detiene el programa) | Sección escrita como clave |
| [CFG-006](#cfg-006) | Error (detiene el programa) | Archivos de MS Project mal indicados |
| [CFG-007](#cfg-007) | Error (detiene el programa) | Mes con formato no válido |
| [CFG-008](#cfg-008) | Error (detiene el programa) | Fecha con formato no válido |
| [CFG-009](#cfg-009) | Error (detiene el programa) | Nivel máximo fuera de rango |
| [CFG-010](#cfg-010) | Error (detiene el programa) | Cantidad de meses fuera de rango |
| [CFG-011](#cfg-011) | Error (detiene el programa) | Línea base fuera de rango |
| [CFG-012](#cfg-012) | Error (detiene el programa) | Distribución no válida |
| [CFG-013](#cfg-013) | Error (detiene el programa) | Decimales fuera de rango |
| [CFG-014](#cfg-014) | Error (detiene el programa) | Umbrales del semáforo incoherentes |
| [CFG-015](#cfg-015) | Error (detiene el programa) | Umbral del semáforo indicado dos veces |
| [CFG-016](#cfg-016) | Error (detiene el programa) | Umbral de `[validacion] fuera` de rango |
| [CFG-017](#cfg-017) | Error (detiene el programa) | Feriados mal indicados |
| [CFG-018](#cfg-018) | Error (detiene el programa) | Lista de métodos mal indicada |
| [CFG-019](#cfg-019) | Error (detiene el programa) | Nombre de empresa vacío |
| [CFG-020](#cfg-020) | Error (detiene el programa) | Nombre de empresa demasiado largo |
| [CFG-021](#cfg-021) | Error (detiene el programa) | Nombre de empresa con caracteres no admitidos |
| [CFG-022](#cfg-022) | Error (detiene el programa) | Las dos empresas tienen el mismo nombre |
| [CFG-023](#cfg-023) | Error (detiene el programa) | Valor de 'ambas' no válido |
| [CFG-024](#cfg-024) | Error (detiene el programa) | Reparto por defecto fuera de rango |
| [CFG-025](#cfg-025) | Error (detiene el programa) | Empresa de la depuración no reconocida |
| [CFG-026](#cfg-026) | Error (detiene el programa) | Umbral de la depuración fuera de rango |
| [CFG-027](#cfg-027) | Error (detiene el programa) | Archivo de `[[proyectos]]` sin nombre |
| [CFG-028](#cfg-028) | Error (detiene el programa) | Nombre de archivo de `[[proyectos]]` no válido |
| [CFG-029](#cfg-029) | Error (detiene el programa) | Nombre de archivo de `[[proyectos]]` repetido |
| [CFG-030](#cfg-030) | Error (detiene el programa) | Archivo de `[[proyectos]]` sin origen |
| [CFG-031](#cfg-031) | Error (detiene el programa) | 'activo' con varios archivos |
| [CFG-032](#cfg-032) | Error (detiene el programa) | No se encontró el archivo de la lista EDT |
| [CFG-033](#cfg-033) | Error (detiene el programa) | No se pudo leer el archivo de la lista EDT |
| [CFG-034](#cfg-034) | Error (detiene el programa) | La lista EDT no tiene códigos |
| [CFG-035](#cfg-035) | Error (detiene el programa) | Extensión del archivo de origen no reconocida |
| [CFG-036](#cfg-036) | Error (detiene el programa) | Falta el archivo de MS Project |
| [CFG-037](#cfg-037) | Error (detiene el programa) | Varios `--origen` sin `[[proyectos]]` |
| [CFG-038](#cfg-038) | Error (detiene el programa) | `--planificacion` o `--linea-base` con varios archivos |
| [CFG-039](#cfg-039) | Error (detiene el programa) | `--origen` no válido con varios archivos |
| [CFG-040](#cfg-040) | Error (detiene el programa) | Falta `--salida` con varios archivos |
| [CFG-041](#cfg-041) | Error (detiene el programa) | Extensión del archivo de salida no válida |
| [CFG-042](#cfg-042) | Error (detiene el programa) | El archivo de salida ya existe |
| [CFG-043](#cfg-043) | Error (detiene el programa) | `--migrar-desde` igual a `--salida` |
| [CFG-044](#cfg-044) | Advertencia | % planificado anterior al mes 1 |
| [CFG-045](#cfg-045) | Advertencia | % planificado posterior al último mes |
| [CFG-046](#cfg-046) | Advertencia | Monedas distintas en los archivos |
| [CFG-047](#cfg-047) | Advertencia | Código desconocido en silenciar |
| [CFG-048](#cfg-048) | Advertencia | Código que no se puede silenciar |
| [CFG-049](#cfg-049) | Info | Observaciones silenciadas |
| [CFG-050](#cfg-050) | Error (detiene el programa) | loe_maximo fuera de rango |
| [CFG-051](#cfg-051) | Error (detiene el programa) | No existe la carpeta de salida |

#### CFG-001

**No se encontró el archivo de configuración** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo indicado con `--config` no existe en esa ruta.

**Qué hacer:** Revisar la ruta (las rutas relativas se toman desde la carpeta en que se ejecuta el programa) o crear el archivo a partir de `config.ejemplo.toml`.

#### CFG-002

**El archivo de configuración no está en UTF-8** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo se guardó con otra codificación (por ejemplo ANSI) y tiene acentos o eñes.

**Qué hacer:** Guardarlo con codificación UTF-8 (en el Bloc de notas: Guardar como > Codificación: UTF-8).

#### CFG-003

**Error de sintaxis en el archivo de configuración** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo no es TOML válido: comillas sin cerrar, una clave repetida, una sección mal escrita... El mensaje indica la línea y la columna.

**Qué hacer:** Corregir la línea indicada. Los textos van entre comillas, los números y true/false sin comillas.

#### CFG-004

**Clave con un valor del tipo incorrecto** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Una clave espera true o false, un número, un número entero o un texto, y tiene otro tipo de valor (por ejemplo, meses = "24" con comillas).

**Qué hacer:** Corregir el valor de la clave indicada (ver `config.ejemplo.toml`).

#### CFG-005

**Sección escrita como clave** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Una sección (por ejemplo `[exportacion]`) se escribió como clave: 'exportacion = ...'.

**Qué hacer:** Escribir la sección en una línea propia, [nombre], seguida de sus claves.

#### CFG-006

**Archivos de MS Project mal indicados** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

La clave proyectos no es una lista de secciones `[[proyectos]]`.

**Qué hacer:** Indicar cada archivo en una sección `[[proyectos]]` (con dos corchetes), con 'nombre' y 'origen'.

#### CFG-007

**Mes con formato no válido** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Un mes (mes_inicio o `--mes-inicio`) no tiene el formato AAAA-MM o el mes no existe.

**Qué hacer:** Indicar el mes como AAAA-MM, por ejemplo 2026-01.

#### CFG-008

**Fecha con formato no válido** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Una fecha (`--fecha-corte`, feriados...) no tiene el formato AAAA-MM-DD o la fecha no existe.

**Qué hacer:** Indicar la fecha como AAAA-MM-DD, por ejemplo 2026-10-31.

#### CFG-009

**Nivel máximo fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

nivel_maximo (o `--nivel-max`) no está entre 1 y 10.

**Qué hacer:** Indicar un nivel entre 1 y 10 (por defecto, 3).

#### CFG-010

**Cantidad de meses fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

meses (o `--meses`) no está entre 1 y 600.

**Qué hacer:** Indicar una cantidad de meses entre 1 y 600 (por defecto, 120).

#### CFG-011

**Línea base fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

linea_base (general, de un archivo de `[[proyectos]]` o `--linea-base`) no está entre 0 y 10.

**Qué hacer:** Indicar 0 para «Línea base» o de 1 a 10 para «Línea base N».

#### CFG-012

**Distribución no válida** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

distribucion no es 'habiles' ni 'calendario'.

**Qué hacer:** Indicar distribucion = "habiles" (por defecto) o "calendario".

#### CFG-013

**Decimales fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

decimales no está entre 0 y 4.

**Qué hacer:** Indicar una cantidad de decimales entre 0 y 4.

#### CFG-014

**Umbrales del semáforo incoherentes** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Los umbrales de `[semaforo] no` cumplen 0 < rojo_inferior <= amarillo_inferior < amarillo_superior <= rojo_superior.

**Qué hacer:** Corregir los umbrales (por defecto 0,90 / 0,95 / 1,05 / 1,10).

#### CFG-015

**Umbral del semáforo indicado dos veces** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

En `[semaforo] se` indicó una clave y también su nombre anterior, que es la misma clave.

**Qué hacer:** Dejar solo el nombre actual de la clave.

#### CFG-016

**Umbral de `[validacion] fuera` de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

meses_0_100 o meses_formula_fija no está entre 0 y 120 meses.

**Qué hacer:** Indicar un valor entre 0 y 120 (0 = no se controla).

#### CFG-017

**Feriados mal indicados** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

feriados no es una lista de fechas.

**Qué hacer:** Indicar los feriados como lista: feriados = ["2026-12-25", "2027-01-01"].

#### CFG-018

**Lista de métodos mal indicada** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

`[metodos] lista` no es una lista de textos.

**Qué hacer:** Indicar los métodos como lista de textos entre comillas.

#### CFG-019

**Nombre de empresa vacío** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

En `[multiempresa]`, empresa_a o empresa_b está vacío.

**Qué hacer:** Indicar el nombre de las dos empresas.

#### CFG-020

**Nombre de empresa demasiado largo** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El nombre de una empresa supera los 15 caracteres: no cabría en el nombre de sus hojas (PV-EmpresaA...).

**Qué hacer:** Usar un nombre o una abreviatura de hasta 15 caracteres.

#### CFG-021

**Nombre de empresa con caracteres no admitidos** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El nombre de una empresa tiene caracteres que Excel no admite en el nombre de una hoja ([ ] : * ? / \ o comillas simples).

**Qué hacer:** Quitar esos caracteres del nombre.

#### CFG-022

**Las dos empresas tienen el mismo nombre** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

empresa_a y empresa_b son iguales (sin distinguir mayúsculas).

**Qué hacer:** Indicar dos nombres distintos.

#### CFG-023

**Valor de 'ambas' no válido** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El valor que indica las tareas de las dos empresas (ambas) está vacío o es el nombre de una empresa.

**Qué hacer:** Indicar un valor distinto de los nombres de las empresas (por defecto, "Ambas").

#### CFG-024

**Reparto por defecto fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

reparto_por_defecto no está entre 0 y 100.

**Qué hacer:** Indicar el % de la empresa A en las tareas compartidas, entre 0 y 100 (por defecto, 50).

#### CFG-025

**Empresa de la depuración no reconocida** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

depuracion_empresa no es empresa_a ni empresa_b.

**Qué hacer:** Indicar el nombre de una de las dos empresas (o dejarlo vacío para depurar empresa_a).

#### CFG-026

**Umbral de la depuración fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

depuracion_umbral no es mayor que 0 y como máximo 100.

**Qué hacer:** Indicar un % mayor que 0 y como máximo 100 (por defecto, 10).

#### CFG-027

**Archivo de `[[proyectos]]` sin nombre** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Una sección `[[proyectos]]` no tiene 'nombre'.

**Qué hacer:** Indicar un nombre breve para cada archivo (p. ej. "P1").

#### CFG-028

**Nombre de archivo de `[[proyectos]]` no válido** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El nombre de un archivo de `[[proyectos]]` es demasiado largo o tiene ':' o ';'.

**Qué hacer:** Usar un nombre breve, sin ':' ni ';'.

#### CFG-029

**Nombre de archivo de `[[proyectos]]` repetido** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Dos secciones `[[proyectos]]` tienen el mismo nombre.

**Qué hacer:** Indicar un nombre distinto para cada archivo.

#### CFG-030

**Archivo de `[[proyectos]]` sin origen** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Una sección `[[proyectos]]` no indica 'origen'.

**Qué hacer:** Indicar la ruta del archivo .mpp o .xml.

#### CFG-031

**'activo' con varios archivos** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Con varios archivos (`[[proyectos]]`), un origen es 'activo' (el proyecto abierto en MS Project).

**Qué hacer:** Indicar la ruta del archivo .mpp o .xml de cada archivo.

#### CFG-032

**No se encontró el archivo de la lista EDT** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo indicado en lista_edt no existe.

**Qué hacer:** Revisar la ruta (las rutas relativas se toman desde la carpeta del archivo de configuración).

#### CFG-033

**No se pudo leer el archivo de la lista EDT** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo de lista_edt existe, pero no se puede leer (permisos, archivo bloqueado...).

**Qué hacer:** Cerrar el archivo en otros programas y revisar sus permisos.

#### CFG-034

**La lista EDT no tiene códigos** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

El archivo de lista_edt está vacío o solo tiene comentarios.

**Qué hacer:** Escribir un código EDT por línea (por ejemplo 1.0.0, 1.1.0, 1.1.1).

#### CFG-035

**Extensión del archivo de origen no reconocida** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas · no se puede silenciar

El archivo de `--origen` no es .mpp, .mpt, .mpd ni .xml.

**Qué hacer:** Usar un archivo .mpp (requiere MS Project) o .xml (MS Project > Guardar como > XML), o indicar `--motor`.

#### CFG-036

**Falta el archivo de MS Project** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas · no se puede silenciar

No se indicó `--origen` ni hay secciones `[[proyectos]]` en la configuración.

**Qué hacer:** Indicar el archivo con `--origen`.

#### CFG-037

**Varios `--origen` sin `[[proyectos]]`** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas · no se puede silenciar

Se indicó `--origen` más de una vez sin secciones `[[proyectos]]`.

**Qué hacer:** Para combinar varios archivos en un libro, indicarlos en secciones `[[proyectos]]` de la configuración.

#### CFG-038

**`--planificacion` o `--linea-base` con varios archivos** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas · no se puede silenciar

Con varios archivos, la planificación y la línea base son de cada archivo.

**Qué hacer:** Indicarlas en la sección `[[proyectos]]` de cada archivo.

#### CFG-039

**`--origen` no válido con varios archivos** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas · no se puede silenciar

Con `[[proyectos]]`, `--origen` debe ser NOMBRE=RUTA, con el nombre de uno de los archivos.

**Qué hacer:** Indicar `--origen` P1=C:\ruta\archivo.mpp, con un nombre de `[[proyectos]]`.

#### CFG-040

**Falta `--salida` con varios archivos** · Error (detiene el programa) · generar y generar-fichas · no se puede silenciar

Con varios archivos no hay un nombre de salida por defecto.

**Qué hacer:** Indicar el archivo a crear con `--salida`.

#### CFG-041

**Extensión del archivo de salida no válida** · Error (detiene el programa) · generar y generar-fichas · no se puede silenciar

El libro se guarda como .xlsx y las fichas como .md.

**Qué hacer:** Indicar `--salida` con la extensión que corresponde.

#### CFG-042

**El archivo de salida ya existe** · Error (detiene el programa) · generar y generar-fichas · no se puede silenciar

El archivo de `--salida` existe y no se indicó `--sobrescribir`.

**Qué hacer:** Usar `--sobrescribir` para reemplazarlo o indicar otro nombre. Para conservar los EV/AC cargados en un libro, se genera un archivo nuevo con `--migrar-desde`.

#### CFG-043

**`--migrar-desde` igual a `--salida`** · Error (detiene el programa) · generar · no se puede silenciar

El libro anterior y el libro nuevo son el mismo archivo.

**Qué hacer:** Generar el libro nuevo con otro nombre.

#### CFG-044

**% planificado anterior al mes 1** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunas filas tienen % planificado antes del primer mes del horizonte (mes_inicio): se acumuló en la columna del mes 1.

**Qué hacer:** Si el horizonte debe empezar antes, indicar mes_inicio (o `--mes-inicio`) y volver a generar.

#### CFG-045

**% planificado posterior al último mes** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunas filas tienen % planificado después del último mes del horizonte: no aparece en las columnas mensuales y su PV no llega al BAC.

**Qué hacer:** Aumentar meses (o `--meses`) y volver a generar.

#### CFG-046

**Monedas distintas en los archivos** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Con varios archivos, los proyectos tienen monedas distintas en MS Project; el libro muestra una de ellas.

**Qué hacer:** Indicar la moneda del libro en `[exportacion] moneda`.

#### CFG-047

**Código desconocido en silenciar** · Advertencia · Todos los comandos · se puede silenciar

Un elemento de `[validacion] silenciar` o de `--silenciar` no es un código del catálogo o no tiene el formato ÁREA-NNN (o ÁREA-NNN:EDT).

**Qué hacer:** Corregir el código según este documento.

#### CFG-048

**Código que no se puede silenciar** · Advertencia · Todos los comandos · se puede silenciar

Se pidió silenciar un error que detiene el programa, el resumen de las observaciones silenciadas o un aviso del registro de diagnóstico ([LEC-022](#lec-022)).

**Qué hacer:** Quitarlo de silenciar: esos mensajes siempre se muestran.

#### CFG-049

**Observaciones silenciadas** · Info · Todos los comandos · no se puede silenciar

Resumen de las observaciones que no se muestran por `[validacion] silenciar` o `--silenciar`, con la cantidad de cada código.

**Qué hacer:** Ninguna. Para volver a verlas, quitar los códigos de silenciar.

#### CFG-050

**loe_maximo fuera de rango** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

`[validacion] loe_maximo` no es un % entre 0 y 100.

**Qué hacer:** Indicar un % entre 0 y 100 (por defecto 15; 0 = no se controla).

#### CFG-051

**No existe la carpeta de salida** · Error (detiene el programa) · generar y generar-fichas · no se puede silenciar

La carpeta del archivo indicado en `--salida` no existe.

**Qué hacer:** Crear la carpeta o indicar otra ruta en `--salida`.

---

### 4.2 LEC · Lectura de MS Project

Apertura del archivo .mpp o .xml y campos configurados que no se encuentran o están vacíos.

| Código | Severidad | Descripción |
|---|---|---|
| [LEC-001](#lec-001) | Error (detiene el programa) | No existe el archivo de origen |
| [LEC-002](#lec-002) | Error (detiene el programa) | El XML no es válido |
| [LEC-003](#lec-003) | Error (detiene el programa) | El XML no es de MS Project |
| [LEC-004](#lec-004) | Error (detiene el programa) | El XML no tiene tareas |
| [LEC-005](#lec-005) | Error (detiene el programa) | Falta pywin32 o MS Project para leer .mpp |
| [LEC-006](#lec-006) | Error (detiene el programa) | No se pudo iniciar MS Project |
| [LEC-007](#lec-007) | Error (detiene el programa) | MS Project no responde |
| [LEC-008](#lec-008) | Error (detiene el programa) | No hay proyectos abiertos en MS Project |
| [LEC-009](#lec-009) | Error (detiene el programa) | MS Project no pudo abrir el archivo |
| [LEC-010](#lec-010) | Advertencia | No se encontró el campo EDT |
| [LEC-011](#lec-011) | Advertencia | No se encontró el campo de método |
| [LEC-012](#lec-012) | Advertencia | No se encontró el campo de peso de hitos |
| [LEC-013](#lec-013) | Advertencia | No se encontró el campo Visible |
| [LEC-014](#lec-014) | Advertencia | No se encontró el campo de descripción |
| [LEC-015](#lec-015) | Advertencia | No se encontró el campo de empresa |
| [LEC-016](#lec-016) | Advertencia | No se encontró el campo de reparto |
| [LEC-017](#lec-017) | Advertencia | Tareas externas omitidas |
| [LEC-018](#lec-018) | Advertencia | Excepción de calendario no interpretada |
| [LEC-019](#lec-019) | Advertencia | Campo de método vacío |
| [LEC-020](#lec-020) | Advertencia | Ningún hito tiene peso |
| [LEC-021](#lec-021) | Advertencia | Campo de empresa vacío |
| [LEC-022](#lec-022) | Advertencia | No se pudo cerrar el proyecto en MS Project |

#### LEC-001

**No existe el archivo de origen** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

El archivo de `--origen` (o de `[[proyectos]]`) no existe en esa ruta.

**Qué hacer:** Revisar la ruta y el nombre del archivo.

#### LEC-002

**El XML no es válido** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

El archivo .xml está dañado o incompleto y no se puede leer.

**Qué hacer:** Volver a guardarlo desde MS Project (Archivo > Guardar como > XML).

#### LEC-003

**El XML no es de MS Project** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

El archivo .xml no tiene el elemento <Project> de MS Project.

**Qué hacer:** Usar un XML guardado desde MS Project (Archivo > Guardar como > XML).

#### LEC-004

**El XML no tiene tareas** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

El archivo .xml de MS Project no tiene la lista de tareas.

**Qué hacer:** Revisar que el proyecto tenga tareas y volver a guardarlo como XML.

#### LEC-005

**Falta pywin32 o MS Project para leer .mpp** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

Los archivos .mpp se leen con MS Project por automatización COM, que necesita el paquete pywin32.

**Qué hacer:** Instalar pywin32 (pip install pywin32) o guardar el proyecto como XML y procesar ese archivo.

#### LEC-006

**No se pudo iniciar MS Project** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

MS Project de escritorio no está instalado o no se puede iniciar por automatización COM.

**Qué hacer:** Verificar la instalación de MS Project, o exportar el proyecto a XML y usar ese archivo.

#### LEC-007

**MS Project no responde** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

MS Project no responde por automatización COM: puede estar cerrándose o mostrando un diálogo (inicio de sesión, recuperación de archivos...).

**Qué hacer:** Abrir MS Project y cerrar el diálogo, o esperar a que termine de cerrarse (o cerrarlo desde el Administrador de tareas), y volver a intentar.

#### LEC-008

**No hay proyectos abiertos en MS Project** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

Se indicó `--origen` activo y MS Project no tiene ningún proyecto abierto.

**Qué hacer:** Abrir el proyecto en MS Project o indicar la ruta del archivo.

#### LEC-009

**MS Project no pudo abrir el archivo** · Error (detiene el programa) · generar, actualizar, validar y generar-fichas (lectura de MS Project) · no se puede silenciar

MS Project no pudo abrir el .mpp (archivo dañado, de una versión más reciente, bloqueado o con un diálogo pendiente).

**Qué hacer:** Abrir el archivo a mano en MS Project para ver el problema, o guardarlo como XML y usar ese archivo.

#### LEC-010

**No se encontró el campo EDT** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El campo indicado en `[campos] edt` no existe en el proyecto: se usa el campo EDT (WBS) de MS Project.

**Qué hacer:** Revisar el nombre del campo (como se ve en MS Project o el nombre interno, p. ej. Texto1).

#### LEC-011

**No se encontró el campo de método** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El campo indicado en `[campos] metodo` no existe: todas las filas tienen el método por defecto.

**Qué hacer:** Revisar el nombre del campo o dejar la clave vacía si no se usa.

#### LEC-012

**No se encontró el campo de peso de hitos** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El campo indicado en `[campos] peso_hito` no existe: los hitos no tienen peso.

**Qué hacer:** Revisar el nombre del campo o dejar la clave vacía si no se usa.

#### LEC-013

**No se encontró el campo Visible** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El campo indicado en `[campos] visible` no existe: las filas del libro siguen el nivel máximo.

**Qué hacer:** Revisar el nombre del campo, o indicar visible = "" si no se usa.

#### LEC-014

**No se encontró el campo de descripción** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El campo indicado en `[campos] descripcion` no existe: las fichas de tareas no tienen descripción.

**Qué hacer:** Revisar el nombre del campo o dejar la clave vacía si no se usa.

#### LEC-015

**No se encontró el campo de empresa** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

En multi-empresa, el campo indicado en campo_empresa no existe: todas las tareas son de las dos empresas.

**Qué hacer:** Revisar `[multiempresa] campo_empresa`.

#### LEC-016

**No se encontró el campo de reparto** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

En multi-empresa, el campo indicado en campo_reparto no existe: las tareas compartidas se reparten con reparto_por_defecto.

**Qué hacer:** Revisar `[multiempresa] campo_reparto`.

#### LEC-017

**Tareas externas omitidas** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

El proyecto tiene tareas externas (vínculos con otros proyectos): no forman parte del cálculo.

**Qué hacer:** Ninguna, si es lo esperado.

#### LEC-018

**Excepción de calendario no interpretada** · Advertencia · generar, actualizar, validar y generar-fichas (lectura de MS Project) · se puede silenciar

Un calendario del XML tiene una excepción con una periodicidad que el programa no interpreta: se omite. Solo afecta el reparto de la duración de las tareas entre los meses.

**Qué hacer:** Si la excepción es importante (p. ej. un cierre anual), indicarla como fechas en feriados.

#### LEC-019

**Campo de método vacío** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El campo de método existe, pero está vacío en todas las tareas: todas las filas tienen el método por defecto.

**Qué hacer:** Cargar el método en la tarea de cada paquete, o revisar `[campos] metodo`.

#### LEC-020

**Ningún hito tiene peso** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El campo de peso de hitos existe, pero ningún hito tiene peso: los hitos no tienen % planificado propio.

**Qué hacer:** Cargar el peso de los hitos (% del paquete), o revisar `[campos] peso_hito`.

#### LEC-021

**Campo de empresa vacío** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

En multi-empresa, el campo de empresa existe, pero está vacío en todas las tareas.

**Qué hacer:** Cargar la empresa de cada tarea, o revisar `[multiempresa] campo_empresa`.

#### LEC-022

**No se pudo cerrar el proyecto en MS Project** · Advertencia · generar, actualizar, validar y generar-fichas con archivos .mpp · no se puede silenciar

Al terminar la lectura de un .mpp, MS Project no cerró el proyecto (por ejemplo, por un diálogo abierto): queda abierto, sin guardar cambios. Se muestra en el registro de diagnóstico (WARNING evm.reader_com), por eso no se puede silenciar.

**Qué hacer:** Cerrar el proyecto a mano en MS Project, sin guardar.

---

### 4.3 EDT · Estructura y códigos EDT

Formato, dependencia, orden y duplicados de los códigos EDT, planificación, lista EDT y filas del libro.

| Código | Severidad | Descripción |
|---|---|---|
| [EDT-001](#edt-001) | Error | Código EDT con formato no válido |
| [EDT-002](#edt-002) | Advertencia | Código EDT no numérico |
| [EDT-003](#edt-003) | Advertencia | Código de nivel inferior sin tarea superior con código |
| [EDT-004](#edt-004) | Error | Mismo código EDT que la tarea superior |
| [EDT-005](#edt-005) | Error | El código EDT no depende del de la tarea superior |
| [EDT-006](#edt-006) | Advertencia | El código EDT salta un nivel |
| [EDT-007](#edt-007) | Advertencia | Código EDT fuera de orden |
| [EDT-008](#edt-008) | Info | Numeración con saltos |
| [EDT-009](#edt-009) | Error | Código EDT duplicado |
| [EDT-010](#edt-010) | Error (detiene el programa, salvo en validar) | No se encontró la planificación |
| [EDT-011](#edt-011) | Error (detiene el programa, salvo en validar) | Código EDT en dos archivos |
| [EDT-012](#edt-012) | Error (detiene el programa, salvo en validar) | Ningún código de la lista EDT existe |
| [EDT-013](#edt-013) | Advertencia | Códigos de la lista EDT sin tarea |
| [EDT-014](#edt-014) | Advertencia | Código de la lista EDT sin tarea |
| [EDT-015](#edt-015) | Info | Elemento de la lista EDT con su tarea superior fuera de la lista |
| [EDT-016](#edt-016) | Info | Tareas inactivas excluidas |
| [EDT-017](#edt-017) | Info | Planificación exportada |
| [EDT-018](#edt-018) | Info | Fila «Otras tareas» |
| [EDT-019](#edt-019) | Advertencia o Info | Tareas fuera de la lista EDT omitidas |
| [EDT-020](#edt-020) | Advertencia o Info | Tareas sin código EDT omitidas |
| [EDT-021](#edt-021) | Advertencia | Ninguna tarea tiene código EDT |
| [EDT-022](#edt-022) | Info | Tareas agregadas en su elemento superior |

#### EDT-001

**Código EDT con formato no válido** · Error · generar (hoja Validación) y validar · se puede silenciar

El código tiene segmentos vacíos (p. ej. 1..2 o 1.2.).

**Qué hacer:** Corregir el código en MS Project: números separados por un punto, como 1.2.3.

#### EDT-002

**Código EDT no numérico** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El código no está formado por números separados por puntos, así que no se puede comprobar con el de su tarea superior.

**Qué hacer:** Usar códigos numéricos (1.2.3). Los códigos sin puntos, como INTERNO, se admiten como planificación.

#### EDT-003

**Código de nivel inferior sin tarea superior con código** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea tiene un código de nivel 2 o más, pero ninguna tarea superior tiene código EDT.

**Qué hacer:** Asignar a la tarea superior el código indicado en el mensaje, o revisar el código de la tarea.

#### EDT-004

**Mismo código EDT que la tarea superior** · Error · generar (hoja Validación) y validar · se puede silenciar

La tarea repite el código de su tarea superior.

**Qué hacer:** Asignarle un código de un nivel más (p. ej. 1.2.1 bajo 1.2.0).

#### EDT-005

**El código EDT no depende del de la tarea superior** · Error · generar (hoja Validación) y validar · se puede silenciar

El código no empieza por el de su tarea superior (p. ej. 1.3.1 bajo 1.2.0).

**Qué hacer:** Corregir el código según el patrón indicado en el mensaje.

#### EDT-006

**El código EDT salta un nivel** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El código tiene dos niveles más que su tarea superior: probablemente falta el código de una tarea intermedia.

**Qué hacer:** Asignar código EDT a la tarea intermedia, o corregir el código.

#### EDT-007

**Código EDT fuera de orden** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El código es menor que el de la tarea anterior del mismo nivel.

**Qué hacer:** Ordenar las tareas en MS Project o corregir la numeración.

#### EDT-008

**Numeración con saltos** · Info · generar (hoja Validación) y validar · se puede silenciar

Entre tareas del mismo nivel faltan números (p. ej. 1.2.1 → 1.2.3).

**Qué hacer:** Ninguna, si es intencional (p. ej. un paquete eliminado); si no, corregir los códigos.

#### EDT-009

**Código EDT duplicado** · Error · generar (hoja Validación) y validar · se puede silenciar

Varias tareas tienen el mismo código EDT: el libro no puede distinguir sus filas.

**Qué hacer:** Corregir los códigos. Si el proyecto tiene varias planificaciones (p. ej. INTERNO y CLIENTE), elegir una con planificacion.

#### EDT-010

**No se encontró la planificación** · Error (detiene el programa, salvo en validar) · generar, actualizar, generar-fichas y validar · no se puede silenciar

Ninguna tarea tiene el código EDT (o el nombre) de la planificación indicada. validar lo informa como error y revisa el proyecto completo.

**Qué hacer:** Revisar planificacion (o `--planificacion`): el mensaje lista las tareas de nivel superior.

#### EDT-011

**Código EDT en dos archivos** · Error (detiene el programa, salvo en validar) · generar, generar-fichas y validar · no se puede silenciar

Con varios archivos de MS Project, el mismo código EDT está en dos de ellos: el libro identifica cada fila por su código. validar lo informa como error.

**Qué hacer:** Revisar los códigos EDT o la planificación de cada archivo (p. ej. 1.0.0 a 8.0.0 en uno y 9.0.0 en otro).

#### EDT-012

**Ningún código de la lista EDT existe** · Error (detiene el programa, salvo en validar) · generar, generar-fichas y validar · no se puede silenciar

Ningún código de lista_edt corresponde a una tarea: el libro no tendría filas. validar lo informa como error.

**Qué hacer:** Revisar la lista, el campo EDT (`[campos] edt`) y la planificación elegida.

#### EDT-013

**Códigos de la lista EDT sin tarea** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunos códigos de lista_edt no existen en MS Project, o su tarea está inactiva o fuera de la planificación: no tienen fila en el libro.

**Qué hacer:** Corregir la lista o el código de la tarea.

#### EDT-014

**Código de la lista EDT sin tarea** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Detalle de [EDT-013](#edt-013): el código de la lista no tiene tarea en MS Project.

**Qué hacer:** Corregir la lista o el código de la tarea.

#### EDT-015

**Elemento de la lista EDT con su tarea superior fuera de la lista** · Info · generar (hoja Validación) y validar · se puede silenciar

El código está en la lista EDT y el de su tarea superior no: la fila queda debajo del elemento listado más cercano.

**Qué hacer:** Ninguna, si es lo esperado; si no, agregar la tarea superior a la lista.

#### EDT-016

**Tareas inactivas excluidas** · Info · generar (hoja Validación) y validar · se puede silenciar

Las tareas inactivas (y su rama) no forman parte del % planificado.

**Qué hacer:** Ninguna. Para incluirlas, activarlas en MS Project.

#### EDT-017

**Planificación exportada** · Info · generar (hoja Validación) y validar · se puede silenciar

Se exportan solo las tareas que dependen de la planificación indicada (p. ej. INTERNO).

**Qué hacer:** Ninguna.

#### EDT-018

**Fila «Otras tareas»** · Info · generar (hoja Validación) y validar · se puede silenciar

Las tareas sin código EDT (o de nivel mayor al máximo, o fuera de la lista EDT) que dependen de una tarea resumen forman su fila -D, con su propio BAC.

**Qué hacer:** Ingresar el BAC de la fila -D, o asignar códigos EDT a esas tareas.

#### EDT-019

**Tareas fuera de la lista EDT omitidas** · Advertencia o Info · generar (hoja Validación) y validar · se puede silenciar

Algunas tareas no están en la lista EDT ni dependen de un elemento de la lista: no forman parte del libro ni del PV y el EV. Es una advertencia si alguna tiene duración o es un hito con peso.

**Qué hacer:** Agregar a la lista los elementos que faltan, o revisar sus códigos.

#### EDT-020

**Tareas sin código EDT omitidas** · Advertencia o Info · generar (hoja Validación) y validar · se puede silenciar

Algunas tareas sin código EDT no dependen de ningún elemento EDT: no forman parte de ninguna fila. Es una advertencia si alguna tiene duración o es un hito con peso.

**Qué hacer:** Asignarles código EDT o moverlas debajo de un elemento EDT.

#### EDT-021

**Ninguna tarea tiene código EDT** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El libro no tiene filas.

**Qué hacer:** Revisar el campo EDT configurado (`[campos] edt`) y la planificación elegida.

#### EDT-022

**Tareas agregadas en su elemento superior** · Info · generar (hoja Validación) y validar · se puede silenciar

Las tareas sin código EDT o de nivel mayor al máximo (o fuera de la lista EDT) se suman en la fila de su elemento EDT superior.

**Qué hacer:** Ninguna.

---

### 4.4 VIS · Campo Visible

Valores del campo Visible y ramas que quedan fuera del cálculo.

| Código | Severidad | Descripción |
|---|---|---|
| [VIS-001](#vis-001) | Advertencia | Valor de Visible no reconocido |
| [VIS-002](#vis-002) | Advertencia | Visible = Sí dentro de una rama excluida |
| [VIS-003](#vis-003) | Advertencia | Visible = Sí bajo una tarea que representa a sus subtareas |
| [VIS-004](#vis-004) | Advertencia | Visible = Sí en una tarea sin código EDT |
| [VIS-005](#vis-005) | Advertencia | Todos los elementos de nivel superior con Visible = No |
| [VIS-006](#vis-006) | Info | Campo Visible sin usar con lista EDT |
| [VIS-007](#vis-007) | Info | Tareas con Visible = No |
| [VIS-008](#vis-008) | Info | Tarea con Visible = No |
| [VIS-009](#vis-009) | Info | Todas las subtareas con Visible = No |
| [VIS-010](#vis-010) | Info | Visible = Sí por debajo del nivel máximo |

#### VIS-001

**Valor de Visible no reconocido** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El campo Visible tiene un valor distinto de Sí o No: se usa el nivel máximo.

**Qué hacer:** Escribir Sí, No o dejarlo vacío.

#### VIS-002

**Visible = Sí dentro de una rama excluida** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea tiene Visible = Sí, pero depende de una tarea con Visible = No: queda fuera del cálculo con esa rama.

**Qué hacer:** Quitar el No de la tarea superior o el Sí de la tarea.

#### VIS-003

**Visible = Sí bajo una tarea que representa a sus subtareas** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea tiene Visible = Sí, pero depende de una tarea con Visible = No cuyas hermanas también lo tienen: se suma con su rama.

**Qué hacer:** Revisar los valores de Visible de esa rama.

#### VIS-004

**Visible = Sí en una tarea sin código EDT** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Una tarea sin código EDT no puede tener fila propia: se suma en su fila superior.

**Qué hacer:** Asignarle código EDT o quitar el Sí.

#### VIS-005

**Todos los elementos de nivel superior con Visible = No** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Si se aplicara, el libro no tendría filas: se ignora en esos elementos.

**Qué hacer:** Revisar el campo Visible de los elementos de nivel superior.

#### VIS-006

**Campo Visible sin usar con lista EDT** · Info · generar (hoja Validación) y validar · se puede silenciar

Con lista_edt, las filas del libro las decide la lista y el campo Visible no se usa.

**Qué hacer:** Ninguna.

#### VIS-007

**Tareas con Visible = No** · Info · generar (hoja Validación) y validar · se puede silenciar

Esas tareas (con su rama) no forman parte del % planificado ni del % ganado.

**Qué hacer:** El BAC y el AC que se ingresan en el libro no las deben incluir.

#### VIS-008

**Tarea con Visible = No** · Info · generar (hoja Validación) y validar · se puede silenciar

Detalle de [VIS-007](#vis-007): la tarea queda fuera del PV, el EV y el AC.

**Qué hacer:** Ninguna, si es lo esperado.

#### VIS-009

**Todas las subtareas con Visible = No** · Info · generar (hoja Validación) y validar · se puede silenciar

La tarea representa a sus subtareas: su fila es el paquete de trabajo, con toda su rama (o se suman en su fila superior).

**Qué hacer:** Ninguna, si es lo esperado.

#### VIS-010

**Visible = Sí por debajo del nivel máximo** · Info · generar (hoja Validación) y validar · se puede silenciar

Esas tareas tienen fila propia aunque su código sea más detallado que el nivel máximo.

**Qué hacer:** Ninguna.

---

### 4.5 FEC · Fechas y duración

Fechas de la línea base y del programa actual, hitos y duración de las tareas.

| Código | Severidad | Descripción |
|---|---|---|
| [FEC-001](#fec-001) | Error | Fin previsto anterior al comienzo previsto |
| [FEC-002](#fec-002) | Error | Fin anterior al comienzo |
| [FEC-003](#fec-003) | Advertencia | Tarea sin Comienzo o sin Fin |
| [FEC-004](#fec-004) | Advertencia | Hito con duración |
| [FEC-005](#fec-005) | Advertencia | Duración 0 sin marca de hito |
| [FEC-006](#fec-006) | Advertencia | Tareas con duración estimada |
| [FEC-007](#fec-007) | Advertencia | Duración estimada |
| [FEC-008](#fec-008) | Advertencia | Tareas programadas manualmente |
| [FEC-009](#fec-009) | Advertencia | El proyecto no tiene la línea base seleccionada |
| [FEC-010](#fec-010) | Advertencia | Tareas sin fechas de línea base |
| [FEC-011](#fec-011) | Advertencia | Tarea sin fechas de línea base |
| [FEC-012](#fec-012) | Advertencia | Tarea sin fechas |

#### FEC-001

**Fin previsto anterior al comienzo previsto** · Error · generar (hoja Validación) y validar · se puede silenciar

En la línea base, la tarea termina antes de comenzar.

**Qué hacer:** Corregir la tarea y volver a guardar su línea base.

#### FEC-002

**Fin anterior al comienzo** · Error · generar (hoja Validación) y validar · se puede silenciar

En el programa actual, la tarea termina antes de comenzar.

**Qué hacer:** Corregir las fechas de la tarea en MS Project.

#### FEC-003

**Tarea sin Comienzo o sin Fin** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea no tiene fechas en el programa actual (p. ej. una tarea programada manualmente sin fechas).

**Qué hacer:** Indicar sus fechas o su duración.

#### FEC-004

**Hito con duración** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea está marcada como hito y tiene duración: se calcula como hito, sin su duración (sin Peso hito no pesa en el % planificado; con Peso hito, pesa entero en el mes de su fecha de fin).

**Qué hacer:** Si no es un hito, quitar la marca de hito (Información de la tarea > Avanzado).

#### FEC-005

**Duración 0 sin marca de hito** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea dura 0 y no es un hito: no pesa en el % planificado de su fila.

**Qué hacer:** Si es un hito, marcarla como hito (y, si corresponde, indicar su peso).

#### FEC-006

**Tareas con duración estimada** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunas tareas tienen duración estimada (con «?», p. ej. «5 días?»): la duración define el % planificado.

**Qué hacer:** Confirmar o corregir la duración de esas tareas.

#### FEC-007

**Duración estimada** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Detalle de [FEC-006](#fec-006): la duración de la tarea es estimada.

**Qué hacer:** Confirmar o corregir la duración.

#### FEC-008

**Tareas programadas manualmente** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

MS Project no recalcula las fechas de esas tareas cuando cambian sus predecesoras o el avance: la lógica del cronograma deja de funcionar en ese punto. En un cronograma integrado no debería haber ninguna.

**Qué hacer:** Programarlas automáticamente. Si es intencional (p. ej. fechas impuestas por terceros), silenciar [FEC-008](#fec-008).

#### FEC-009

**El proyecto no tiene la línea base seleccionada** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El % planificado se calculó con las fechas actuales de las tareas.

**Qué hacer:** Guardar la línea base en MS Project (Proyecto > Establecer línea base) o elegir otra con linea_base, y volver a exportar.

#### FEC-010

**Tareas sin fechas de línea base** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunas tareas no tienen fechas de línea base (p. ej. se agregaron después de guardarla): su % planificado se calculó con sus fechas actuales.

**Qué hacer:** Si corresponde, actualizar la línea base de esas tareas.

#### FEC-011

**Tarea sin fechas de línea base** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Detalle de [FEC-010](#fec-010): se usaron las fechas actuales de la tarea.

**Qué hacer:** Si corresponde, actualizar la línea base de la tarea.

#### FEC-012

**Tarea sin fechas** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea no tiene fechas de línea base ni fechas actuales: no aporta % planificado.

**Qué hacer:** Indicar sus fechas o su duración.

---

### 4.6 HIT · Hitos y pesos

Peso de los hitos, método «Hitos ponderados» y filas sin % planificado.

| Código | Severidad | Descripción |
|---|---|---|
| [HIT-001](#hit-001) | Advertencia | Peso en una tarea que no es un hito de detalle |
| [HIT-002](#hit-002) | Advertencia | «Hitos ponderados» sin hitos |
| [HIT-003](#hit-003) | Info | «Hitos ponderados» con hitos sin peso |
| [HIT-004](#hit-004) | Info | «Hitos ponderados» con algunos hitos sin peso |
| [HIT-005](#hit-005) | Advertencia | Fila sin % planificado |
| [HIT-006](#hit-006) | Advertencia | Hitos con peso negativo |
| [HIT-007](#hit-007) | Advertencia | Los pesos de los hitos suman más de 100 % |
| [HIT-008](#hit-008) | Advertencia | Pesos de hitos menores a 100 % sin tareas con duración |
| [HIT-009](#hit-009) | Advertencia | Tareas sin duración ni hitos con peso |

#### HIT-001

**Peso en una tarea que no es un hito de detalle** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea tiene peso de hito, pero no es un hito o es una tarea resumen: el peso no se usa (las tareas pesan por su duración).

**Qué hacer:** Quitar el peso, o marcar la tarea como hito.

#### HIT-002

**«Hitos ponderados» sin hitos** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La fila usa el método «Hitos ponderados» y no tiene hitos: actualizar calculará el % ganado con el % físico completado (o el % completado).

**Qué hacer:** Agregar los hitos del paquete o cambiar el método.

#### HIT-003

**«Hitos ponderados» con hitos sin peso** · Info · generar (hoja Validación) y validar · se puede silenciar

Ningún hito de la fila tiene peso: los hitos se ponderan por igual.

**Qué hacer:** Indicar el peso de los hitos, si no deben pesar igual.

#### HIT-004

**«Hitos ponderados» con algunos hitos sin peso** · Info · generar (hoja Validación) y validar · se puede silenciar

Los hitos sin peso no suman % ganado al cumplirse.

**Qué hacer:** Indicar el peso de esos hitos, si deben sumar.

#### HIT-005

**Fila sin % planificado** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Ninguna tarea de la fila tiene fechas ni es un hito con peso: su BAC no se puede repartir en el tiempo.

**Qué hacer:** Indicar fechas o duración a sus tareas, o el peso de sus hitos.

#### HIT-006

**Hitos con peso negativo** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunos hitos tienen peso negativo: se toma como 0.

**Qué hacer:** Corregir el peso de esos hitos.

#### HIT-007

**Los pesos de los hitos suman más de 100 %** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Los pesos se escalaron a 100 % y las tareas sin hito de la fila no reciben % planificado.

**Qué hacer:** Corregir los pesos para que sumen como máximo 100 %.

#### HIT-008

**Pesos de hitos menores a 100 % sin tareas con duración** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La fila solo tiene hitos y sus pesos suman menos de 100 %: se escalaron a 100 %.

**Qué hacer:** Corregir los pesos para que sumen 100 %.

#### HIT-009

**Tareas sin duración ni hitos con peso** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Ninguna tarea de la fila tiene duración ni es un hito con peso: todas pesan igual.

**Qué hacer:** Indicar la duración de las tareas o el peso de los hitos.

---

### 4.7 MET · Método de medición

Método de medición del EV de cada fila.

| Código | Severidad | Descripción |
|---|---|---|
| [MET-001](#met-001) | Advertencia | Método no reconocido |
| [MET-002](#met-002) | Info | Subtareas con otro método |
| [MET-003](#met-003) | Advertencia | Fila sin método con subtareas con método |
| [MET-004](#met-004) | Advertencia | Método de fórmula fija en un paquete largo |
| [MET-005](#met-005) | Info | Filas sin método definido |

#### MET-001

**Método no reconocido** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El método de la fila no es uno de los admitidos: actualizar calculará el % ganado con el % completado.

**Qué hacer:** Corregir el método en MS Project o en la columna C del libro.

#### MET-002

**Subtareas con otro método** · Info · generar (hoja Validación) y validar · se puede silenciar

Algunas subtareas de la fila tienen otro método en MS Project: se usa el de la tarea de la fila.

**Qué hacer:** Ninguna, si es lo esperado.

#### MET-003

**Fila sin método con subtareas con método** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

La tarea de la fila no tiene método y sus subtareas sí: se usa el método por defecto.

**Qué hacer:** Indicar el método en la tarea de la fila.

#### MET-004

**Método de fórmula fija en un paquete largo** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Con 0/100 (o 50/50, 25/75...) el EV queda fijo hasta que el paquete termina: en un paquete largo, el EV no refleja el avance.

**Qué hacer:** Dividir el paquete o usar un método de avance (% completado, hitos ponderados). Los umbrales son `[validacion] meses_0_100` y meses_formula_fija.

#### MET-005

**Filas sin método definido** · Info · generar (hoja Validación) y validar · se puede silenciar

Algunas filas no tienen método de medición (columna C).

**Qué hacer:** Indicar el método en MS Project o en la columna C del libro.

---

### 4.8 EMP · Empresas

Multi-empresa: empresa y % de reparto de cada tarea.

| Código | Severidad | Descripción |
|---|---|---|
| [EMP-001](#emp-001) | Info | % de reparto en una tarea que no es compartida |
| [EMP-002](#emp-002) | Info | Empresa sin tareas en la fila |
| [EMP-003](#emp-003) | Advertencia | Tareas sin empresa |
| [EMP-004](#emp-004) | Advertencia | Tarea sin empresa |
| [EMP-005](#emp-005) | Advertencia | Empresa no reconocida |
| [EMP-006](#emp-006) | Advertencia | % de reparto fuera de rango |

#### EMP-001

**% de reparto en una tarea que no es compartida** · Info · generar (hoja Validación) y validar · se puede silenciar

La tarea tiene % de reparto, pero su empresa no es la de las tareas compartidas (ambas): el reparto no se usa.

**Qué hacer:** Quitar el reparto o cambiar la empresa de la tarea.

#### EMP-002

**Empresa sin tareas en la fila** · Info · generar (hoja Validación) y validar · se puede silenciar

La empresa no participa en las tareas de algunas filas: si se le asigna BAC, su PV se reparte con el % planificado de la fila completa.

**Qué hacer:** Ninguna, si no se le asigna BAC en esas filas.

#### EMP-003

**Tareas sin empresa** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Algunas tareas no tienen empresa (ni la heredan de sus tareas superiores): se tratan como compartidas, con el reparto por defecto.

**Qué hacer:** Indicar la empresa de esas tareas o de su tarea superior.

#### EMP-004

**Tarea sin empresa** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

Detalle de [EMP-003](#emp-003): la tarea (y sus subtareas sin empresa) se trata como compartida.

**Qué hacer:** Indicar la empresa de la tarea.

#### EMP-005

**Empresa no reconocida** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El valor del campo de empresa no es empresa_a, empresa_b ni ambas: la tarea se trata como compartida.

**Qué hacer:** Corregir el valor en MS Project.

#### EMP-006

**% de reparto fuera de rango** · Advertencia · generar (hoja Validación) y validar · se puede silenciar

El % de reparto de la tarea no es mayor que 0 y como máximo 100: se usa el reparto por defecto.

**Qué hacer:** Corregir el % de reparto.

---

### 4.9 AVA · Avance a la fecha de estado

Fechas reales, % completado, avance actualizado a la fecha de estado y fuera de secuencia, e índices DCMA del avance: BEI, tareas incumplidas y CPLI (solo validar).

| Código | Severidad | Descripción |
|---|---|---|
| [AVA-001](#ava-001) | Error | Fin real sin Comienzo real |
| [AVA-002](#ava-002) | Error | Fin real anterior al Comienzo real |
| [AVA-003](#ava-003) | Advertencia | % completado sin Comienzo real |
| [AVA-004](#ava-004) | Advertencia | 100 % completado sin Fin real |
| [AVA-005](#ava-005) | Advertencia | Fin real con % completado menor a 100 |
| [AVA-006](#ava-006) | Advertencia | Avance sin fecha de estado |
| [AVA-007](#ava-007) | Info | Sin fecha de estado ni avance |
| [AVA-008](#ava-008) | Error | Avance registrado en el futuro |
| [AVA-009](#ava-009) | Advertencia | Tarea atrasada sin terminar |
| [AVA-010](#ava-010) | Advertencia | Tarea sin comenzar con Comienzo pasado |
| [AVA-011](#ava-011) | Advertencia | Trabajo pendiente antes de la fecha de estado |
| [AVA-012](#ava-012) | Info | Tareas que debían comenzar y no comenzaron |
| [AVA-013](#ava-013) | Advertencia o Info | Índice de ejecución de la línea base (BEI) |
| [AVA-014](#ava-014) | Advertencia o Info | Tareas incumplidas respecto de la línea base |
| [AVA-015](#ava-015) | Advertencia o Info | Índice de longitud del camino crítico (CPLI) |
| [AVA-016](#ava-016) | Advertencia | Avance fuera de secuencia |
| [AVA-017](#ava-017) | Advertencia | Tarea con avance fuera de secuencia |

#### AVA-001

**Fin real sin Comienzo real** · Error · validar · se puede silenciar

La tarea tiene Fin real y no tiene Comienzo real.

**Qué hacer:** Registrar el Comienzo real.

#### AVA-002

**Fin real anterior al Comienzo real** · Error · validar · se puede silenciar

Las fechas reales de la tarea están invertidas.

**Qué hacer:** Corregir las fechas reales.

#### AVA-003

**% completado sin Comienzo real** · Advertencia · validar · se puede silenciar

La tarea tiene % completado y no tiene Comienzo real.

**Qué hacer:** Registrar el Comienzo real (MS Project lo completa al registrar el avance).

#### AVA-004

**100 % completado sin Fin real** · Advertencia · validar · se puede silenciar

La tarea está completa y no tiene Fin real.

**Qué hacer:** Registrar el Fin real.

#### AVA-005

**Fin real con % completado menor a 100** · Advertencia · validar · se puede silenciar

La tarea tiene Fin real, pero el programa la considera terminada solo con el 100 %.

**Qué hacer:** Completar el % completado o quitar el Fin real.

#### AVA-006

**Avance sin fecha de estado** · Advertencia · validar · se puede silenciar

El proyecto tiene avance cargado y no tiene fecha de estado: no se revisó el avance a la fecha de estado.

**Qué hacer:** Indicar la fecha de estado en MS Project (Proyecto > Información del proyecto) o con `--fecha-corte`; actualizar la necesita.

#### AVA-007

**Sin fecha de estado ni avance** · Info · validar · se puede silenciar

El proyecto no tiene fecha de estado ni avance cargado: no se revisó el avance.

**Qué hacer:** Ninguna; la fecha se puede indicar con `--fecha-corte`.

#### AVA-008

**Avance registrado en el futuro** · Error · validar · se puede silenciar

El Comienzo real o el Fin real de la tarea es posterior a la fecha de estado.

**Qué hacer:** Corregir la fecha real o la fecha de estado.

#### AVA-009

**Tarea atrasada sin terminar** · Advertencia · validar · se puede silenciar

El Fin de la tarea es anterior a la fecha de estado y la tarea no está terminada.

**Qué hacer:** Registrar el avance o reprogramar el trabajo pendiente después de la fecha de estado.

#### AVA-010

**Tarea sin comenzar con Comienzo pasado** · Advertencia · validar · se puede silenciar

El Comienzo de la tarea es anterior a la fecha de estado y no tiene avance ni Comienzo real.

**Qué hacer:** Registrar el avance o reprogramar la tarea después de la fecha de estado.

#### AVA-011

**Trabajo pendiente antes de la fecha de estado** · Advertencia · validar · se puede silenciar

Algunas tareas tienen trabajo pendiente programado antes de la fecha de estado: el cronograma no está actualizado.

**Qué hacer:** Registrar el avance o reprogramar el trabajo pendiente (Proyecto > Actualizar proyecto).

#### AVA-012

**Tareas que debían comenzar y no comenzaron** · Info · validar · se puede silenciar

Según la línea base, algunas tareas debían comenzar a la fecha de estado y no comenzaron.

**Qué hacer:** Revisar el atraso de esas tareas.

#### AVA-013

**Índice de ejecución de la línea base (BEI)** · Advertencia o Info · validar · se puede silenciar

Tareas terminadas frente a las que debían terminar a la fecha de estado. Es una advertencia si es menor que 0,95 (criterio DCMA).

**Qué hacer:** Revisar las tareas sin terminar que lista el mensaje.

#### AVA-014

**Tareas incumplidas respecto de la línea base** · Advertencia o Info · validar · se puede silenciar

De las tareas que debían terminar a la fecha de estado según la línea base, las que no terminaron o terminaron después de su fin de línea base. Es una advertencia si superan el 5 % (criterio DCMA).

**Qué hacer:** Revisar el atraso de las tareas que lista el mensaje.

#### AVA-015

**Índice de longitud del camino crítico (CPLI)** · Advertencia o Info · validar · se puede silenciar

Días laborables que quedan hasta el fin previsto del proyecto, más la holgura frente al fin de la línea base, divididos por los días que quedan. Menor que 1 indica que el fin previsto es posterior al de la línea base; es una advertencia si es menor que 0,95 (criterio DCMA).

**Qué hacer:** Revisar las tareas críticas: con un CPLI bajo, cumplir el fin de la línea base exige recuperar el atraso en el trabajo que queda.

#### AVA-016

**Avance fuera de secuencia** · Advertencia · validar · se puede silenciar

Resumen de [AVA-017](#ava-017): tareas con avance que su vínculo con la predecesora no permite todavía (p. ej. la sucesora comenzó y la predecesora fin a comienzo no terminó). Las fechas y la holgura del trabajo que queda pueden no ser realistas.

**Qué hacer:** Corregir el avance registrado o el vínculo, si ya no representa la secuencia real del trabajo.

#### AVA-017

**Tarea con avance fuera de secuencia** · Advertencia · validar · se puede silenciar

Detalle de [AVA-016](#ava-016): la tarea tiene avance y su predecesora no terminó (vínculo fin a comienzo o fin a fin) o no comenzó (comienzo a comienzo o comienzo a fin). Los vínculos con adelanto no se revisan ([LOG-006](#log-006)).

**Qué hacer:** Corregir el avance de la tarea o de su predecesora, o el vínculo.

---

### 4.10 LOG · Lógica del cronograma

Vínculos, delimitaciones, holgura y duración de las tareas, según los criterios de la evaluación DCMA de 14 puntos (solo validar).

| Código | Severidad | Descripción |
|---|---|---|
| [LOG-001](#log-001) | Advertencia | Cronograma sin vínculos |
| [LOG-002](#log-002) | Advertencia | Tareas sin predecesoras o sin sucesoras |
| [LOG-003](#log-003) | Advertencia | Tarea sin predecesoras ni sucesoras |
| [LOG-004](#log-004) | Advertencia | Tarea sin predecesoras |
| [LOG-005](#log-005) | Advertencia | Tarea sin sucesoras |
| [LOG-006](#log-006) | Advertencia | Vínculo con adelanto |
| [LOG-007](#log-007) | Advertencia | Vínculo comienzo a fin (CF) |
| [LOG-008](#log-008) | Advertencia o Info | Vínculos con posición positiva |
| [LOG-009](#log-009) | Advertencia | Pocos vínculos fin a comienzo |
| [LOG-010](#log-010) | Advertencia | Tarea resumen con vínculos |
| [LOG-011](#log-011) | Advertencia | Tareas con delimitaciones |
| [LOG-012](#log-012) | Advertencia | Tarea con delimitación |
| [LOG-013](#log-013) | Advertencia o Info | Tareas con holgura alta |
| [LOG-014](#log-014) | Advertencia | Tareas con holgura negativa |
| [LOG-015](#log-015) | Advertencia | Tarea con holgura negativa |
| [LOG-016](#log-016) | Advertencia o Info | Tareas de duración alta |

#### LOG-001

**Cronograma sin vínculos** · Advertencia · validar · se puede silenciar

Las fechas no se recalculan cuando cambian las duraciones o el avance.

**Qué hacer:** Vincular las tareas (cada tarea con al menos una predecesora y una sucesora).

#### LOG-002

**Tareas sin predecesoras o sin sucesoras** · Advertencia · validar · se puede silenciar

Resumen de [LOG-003](#log-003) a [LOG-005](#log-005), con el % de tareas sin terminar (DCMA: hasta 5 %).

**Qué hacer:** Vincular esas tareas.

#### LOG-003

**Tarea sin predecesoras ni sucesoras** · Advertencia · validar · se puede silenciar

Las fechas de la tarea no dependen de la lógica del cronograma.

**Qué hacer:** Vincular la tarea.

#### LOG-004

**Tarea sin predecesoras** · Advertencia · validar · se puede silenciar

El comienzo de la tarea no depende de otra tarea.

**Qué hacer:** Agregar una predecesora.

#### LOG-005

**Tarea sin sucesoras** · Advertencia · validar · se puede silenciar

El fin de la tarea no afecta a ninguna otra tarea.

**Qué hacer:** Agregar una sucesora.

#### LOG-006

**Vínculo con adelanto** · Advertencia · validar · se puede silenciar

El vínculo tiene posición negativa (adelanto); la evaluación DCMA recomienda no usarlos.

**Qué hacer:** Reemplazar el adelanto por tareas o vínculos sin adelanto.

#### LOG-007

**Vínculo comienzo a fin (CF)** · Advertencia · validar · se puede silenciar

Es un tipo de vínculo poco habitual.

**Qué hacer:** Usar un vínculo fin a comienzo (FC), salvo que sea necesario.

#### LOG-008

**Vínculos con posición positiva** · Advertencia o Info · validar · se puede silenciar

Una espera fija oculta trabajo que conviene representar como tarea. Es una advertencia si superan el 5 % de los vínculos (criterio DCMA).

**Qué hacer:** Reemplazar las esperas por tareas, si representan trabajo.

#### LOG-009

**Pocos vínculos fin a comienzo** · Advertencia · validar · se puede silenciar

Menos del 90 % de los vínculos son fin a comienzo (FC) (criterio DCMA).

**Qué hacer:** Revisar los vínculos de otro tipo.

#### LOG-010

**Tarea resumen con vínculos** · Advertencia · validar · se puede silenciar

Los vínculos de una tarea resumen ocultan la lógica de sus tareas de detalle.

**Qué hacer:** Vincular las tareas de detalle.

#### LOG-011

**Tareas con delimitaciones** · Advertencia · validar · se puede silenciar

Las delimitaciones que fijan fechas (Debe comenzar el, No finalizar después del...) impiden que la lógica mueva esas tareas (DCMA: hasta 5 %).

**Qué hacer:** Usar «Lo antes posible» y, si hace falta, una fecha límite.

#### LOG-012

**Tarea con delimitación** · Advertencia · validar · se puede silenciar

Detalle de [LOG-011](#log-011): la delimitación de la tarea.

**Qué hacer:** Usar «Lo antes posible» y, si hace falta, una fecha límite.

#### LOG-013

**Tareas con holgura alta** · Advertencia o Info · validar · se puede silenciar

Tareas sin terminar con más de 44 días laborables de holgura total: suele indicar vínculos que faltan. Es una advertencia si superan el 5 % de las tareas sin terminar (criterio DCMA).

**Qué hacer:** Revisar las sucesoras de esas tareas: la holgura alta desaparece al vincularlas con el trabajo que depende de ellas.

#### LOG-014

**Tareas con holgura negativa** · Advertencia · validar · se puede silenciar

Resumen de [LOG-015](#log-015): tareas sin terminar cuyo fin previsto supera una fecha límite o una delimitación. La evaluación DCMA no admite holgura negativa.

**Qué hacer:** Recuperar el atraso (cambiar la secuencia o la duración del trabajo que queda) o revisar la fecha límite.

#### LOG-015

**Tarea con holgura negativa** · Advertencia · validar · se puede silenciar

Detalle de [LOG-014](#log-014): la holgura total de la tarea.

**Qué hacer:** Revisar las predecesoras de la tarea y su fecha límite o delimitación.

#### LOG-016

**Tareas de duración alta** · Advertencia o Info · validar · se puede silenciar

Tareas sin terminar con más de 44 días laborables de duración en la línea base (sin los hitos ni las filas LOE): son difíciles de seguir y de medir. Es una advertencia si superan el 5 % de las tareas sin terminar (criterio DCMA).

**Qué hacer:** Dividir esas tareas en tareas más cortas, con entregables intermedios.

---

### 4.11 LIB · Libro

Libro Excel: recordatorio del BAC y libros dañados, modificados o abiertos en Excel.

| Código | Severidad | Descripción |
|---|---|---|
| [LIB-001](#lib-001) | Info | Cálculo del PV y el EV |
| [LIB-002](#lib-002) | Error (detiene el programa) | No existe el libro |
| [LIB-003](#lib-003) | Error (detiene el programa) | El libro no es .xlsx |
| [LIB-004](#lib-004) | Error (detiene el programa) | No se pudo abrir el libro |
| [LIB-005](#lib-005) | Error (detiene el programa) | El libro no fue generado por este programa |
| [LIB-006](#lib-006) | Error (detiene el programa) | Falta un dato de la hoja `_meta` |
| [LIB-007](#lib-007) | Error (detiene el programa) | Dato de la hoja `_meta` no válido |
| [LIB-008](#lib-008) | Error (detiene el programa) | Hoja `_estructura` no válida |
| [LIB-009](#lib-009) | Error (detiene el programa) | Libro de una versión anterior |
| [LIB-010](#lib-010) | Error (detiene el programa) | Falta una hoja del libro |
| [LIB-011](#lib-011) | Error (detiene el programa) | Filas del libro movidas |
| [LIB-012](#lib-012) | Error (detiene el programa) | Fórmula sin valor calculado |
| [LIB-013](#lib-013) | Error (detiene el programa) | Valor no numérico |
| [LIB-014](#lib-014) | Error (detiene el programa) | Número guardado como texto |
| [LIB-015](#lib-015) | Error (detiene el programa) | No se puede guardar el archivo |

#### LIB-001

**Cálculo del PV y el EV** · Info · generar (hoja Validación) · se puede silenciar

Recordatorio: el PV y el EV de cada fila son el BAC × el % planificado y el % ganado.

**Qué hacer:** Ingresar el BAC de cada fila en las celdas amarillas de la columna D.

#### LIB-002

**No existe el libro** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

El libro indicado (`--libro` o `--migrar-desde`) no existe en esa ruta.

**Qué hacer:** Revisar la ruta y el nombre del libro.

#### LIB-003

**El libro no es .xlsx** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

El archivo indicado no es un libro .xlsx.

**Qué hacer:** Indicar el libro .xlsx generado con generar.

#### LIB-004

**No se pudo abrir el libro** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

El archivo no se puede abrir como libro de Excel (dañado, cifrado o de otro formato).

**Qué hacer:** Recuperar el libro desde un respaldo.

#### LIB-005

**El libro no fue generado por este programa** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Faltan las hojas ocultas `_meta` y `_estructura`.

**Qué hacer:** Indicar un libro generado con generar.

#### LIB-006

**Falta un dato de la hoja `_meta`** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

La hoja oculta `_meta` del libro no tiene un dato necesario.

**Qué hacer:** Recuperar el libro desde un respaldo (la hoja `_meta` no se debe modificar).

#### LIB-007

**Dato de la hoja `_meta` no válido** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Un dato de la hoja oculta `_meta` fue modificado.

**Qué hacer:** Recuperar el libro desde un respaldo (la hoja `_meta` no se debe modificar).

#### LIB-008

**Hoja `_estructura` no válida** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

La hoja oculta `_estructura` fue modificada.

**Qué hacer:** Recuperar el libro desde un respaldo (la hoja `_estructura` no se debe modificar).

#### LIB-009

**Libro de una versión anterior** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

El libro se generó con la versión 1.x, que calculaba el PV y el EV con los costos de MS Project.

**Qué hacer:** Usar la versión 1.x para ese libro, o generar un libro nuevo.

#### LIB-010

**Falta una hoja del libro** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Una hoja que el programa lee o escribe se eliminó o se cambió de nombre.

**Qué hacer:** Restaurar su nombre original o recuperar el libro desde un respaldo.

#### LIB-011

**Filas del libro movidas** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Una hoja mensual no tiene las filas con las que se generó el libro: se insertaron, eliminaron u ordenaron filas. actualizar ubica cada fila por su posición.

**Qué hacer:** Deshacer el cambio o recuperar el libro desde un respaldo. Para cambiar la EDT, se modifica MS Project y se regenera el libro con `--migrar-desde`.

#### LIB-012

**Fórmula sin valor calculado** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Una celda de ingreso tiene una fórmula que Excel no calculó (el libro se guardó sin calcular).

**Qué hacer:** Abrir y guardar el libro en Excel, o reemplazar la fórmula por un número.

#### LIB-013

**Valor no numérico** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Una celda de ingreso tiene un texto o un valor que no es un número.

**Qué hacer:** Corregir el valor de la celda indicada.

#### LIB-014

**Número guardado como texto** · Error (detiene el programa) · actualizar, analizar y generar `--migrar-desde` (lectura del libro) · no se puede silenciar

Una celda tiene un número guardado como texto: Excel no lo suma.

**Qué hacer:** Ingresarlo como número.

#### LIB-015

**No se puede guardar el archivo** · Error (detiene el programa) · generar, actualizar y analizar · no se puede silenciar

El libro (o el archivo de salida) está abierto en Excel o en otro programa, o no hay permiso para escribirlo.

**Qué hacer:** Cerrar el archivo y volver a ejecutar el comando.

---

### 4.12 MIG · Migración (--migrar-desde)

Traslado del BAC, el % ganado, el AC y los métodos de un libro anterior.

| Código | Severidad | Descripción |
|---|---|---|
| [MIG-001](#mig-001) | Error (detiene el programa) | Migración entre una y dos empresas |
| [MIG-002](#mig-002) | Advertencia | Libro anterior de un archivo y nuevo de varios |
| [MIG-003](#mig-003) | Advertencia | Libro anterior de varios archivos y nuevo de uno |
| [MIG-004](#mig-004) | Advertencia | Archivo del libro anterior que no está en el nuevo |
| [MIG-005](#mig-005) | Advertencia | Proyecto distinto |
| [MIG-006](#mig-006) | Advertencia | Nivel máximo distinto |
| [MIG-007](#mig-007) | Info | Historial sumado en el paquete |
| [MIG-008](#mig-008) | Advertencia | Historial asociado por código EDT |
| [MIG-009](#mig-009) | Advertencia | % ganado de varias filas sin BAC |
| [MIG-010](#mig-010) | Info | Métodos de medición conservados |
| [MIG-011](#mig-011) | Advertencia | EV/AC antes del primer mes del nuevo horizonte |
| [MIG-012](#mig-012) | Advertencia | EV/AC después del último mes del nuevo horizonte |
| [MIG-013](#mig-013) | Info | BAC conservado |
| [MIG-014](#mig-014) | Error | Historial que no se pudo trasladar |
| [MIG-015](#mig-015) | Advertencia | Análisis de desviaciones sin fila en el libro nuevo |
| [MIG-016](#mig-016) | Info | Análisis de desviaciones trasladado |
| [MIG-017](#mig-017) | Advertencia | Filas sin trasladar |
| [MIG-018](#mig-018) | Info | Registro de cambios trasladado |
| [MIG-019](#mig-019) | Error | Valores escritos sobre fórmulas sin trasladar |

#### MIG-001

**Migración entre una y dos empresas** · Error (detiene el programa) · generar `--migrar-desde` · no se puede silenciar

El historial de EV y AC de un libro de una sola empresa no se puede repartir entre dos empresas (ni al revés).

**Qué hacer:** Generar el libro nuevo sin `--migrar-desde` y cargar los datos a mano.

#### MIG-002

**Libro anterior de un archivo y nuevo de varios** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Los EV/AC se asocian por código EDT.

**Qué hacer:** Verificar el traslado en el libro nuevo.

#### MIG-003

**Libro anterior de varios archivos y nuevo de uno** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Los EV/AC se asocian por código EDT.

**Qué hacer:** Verificar el traslado en el libro nuevo.

#### MIG-004

**Archivo del libro anterior que no está en el nuevo** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Un archivo de `[[proyectos]]` del libro anterior no forma parte del nuevo: sus filas se asocian por código EDT.

**Qué hacer:** Verificar el traslado de esas filas.

#### MIG-005

**Proyecto distinto** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

El libro anterior corresponde a otro proyecto de MS Project: los EV/AC se asocian por código EDT.

**Qué hacer:** Verificar el traslado en el libro nuevo.

#### MIG-006

**Nivel máximo distinto** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Las filas que ahora se dividen en varias no pueden repartir su historial.

**Qué hacer:** Cargar a mano el historial de las filas que informa [MIG-014](#mig-014).

#### MIG-007

**Historial sumado en el paquete** · Info · generar `--migrar-desde` (hoja Validación) · se puede silenciar

El historial de una fila anterior que quedó dentro de un paquete nuevo se sumó en ese paquete.

**Qué hacer:** Ninguna.

#### MIG-008

**Historial asociado por código EDT** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

La fila recibió el historial de una fila anterior con otro nombre, asociada por código EDT.

**Qué hacer:** Verificar que sea el mismo paquete.

#### MIG-009

**% ganado de varias filas sin BAC** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

La fila recibió el % ganado de varias filas anteriores sin BAC: no se pudo combinar.

**Qué hacer:** Revisar la hoja % EV de esa fila.

#### MIG-010

**Métodos de medición conservados** · Info · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se conservaron los métodos editados en la columna C del libro anterior.

**Qué hacer:** Revisarlos si el libro anterior no indica si se editaron en Excel.

#### MIG-011

**EV/AC antes del primer mes del nuevo horizonte** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se acumularon en el mes 1.

**Qué hacer:** Revisar mes_inicio del libro nuevo.

#### MIG-012

**EV/AC después del último mes del nuevo horizonte** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se acumularon en el último mes.

**Qué hacer:** Aumentar meses del libro nuevo.

#### MIG-013

**BAC conservado** · Info · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se conservó el BAC ingresado en el libro anterior.

**Qué hacer:** Ninguna.

#### MIG-014

**Historial que no se pudo trasladar** · Error · generar `--migrar-desde` (hoja Validación) · se puede silenciar

La tarea de la fila anterior no existe en el proyecto nuevo o se dividió en varias filas.

**Qué hacer:** Cargar a mano el BAC, el EV y el AC en las filas que correspondan.

#### MIG-015

**Análisis de desviaciones sin fila en el libro nuevo** · Advertencia · generar `--migrar-desde` (hoja Validación) · se puede silenciar

El paquete ya no tiene fila: el texto del análisis se conservó, sin valores.

**Qué hacer:** Revisar el texto en la hoja Análisis de desviaciones.

#### MIG-016

**Análisis de desviaciones trasladado** · Info · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se trasladaron las filas de la hoja Análisis de desviaciones.

**Qué hacer:** Ninguna.

#### MIG-017

**Filas sin trasladar** · Advertencia · generar `--migrar-desde` · se puede silenciar

Resumen en la consola de las filas con historial que no se pudieron trasladar (ver [MIG-014](#mig-014)).

**Qué hacer:** Ver la hoja Validación del libro nuevo.

#### MIG-018

**Registro de cambios trasladado** · Info · generar `--migrar-desde` (hoja Validación) · se puede silenciar

Se trasladaron las filas de la hoja Registro de cambios y la referencia de los valores informados, con la que actualizar detecta los cambios siguientes.

**Qué hacer:** Ninguna.

#### MIG-019

**Valores escritos sobre fórmulas sin trasladar** · Error · generar `--migrar-desde` · se puede silenciar

El libro anterior tiene valores escritos sobre celdas que deben tener fórmula (subtotales, fila TOTAL o suma de las empresas; ver [ACT-024](#act-024)). El libro nuevo tiene la fórmula en esas celdas: los valores no se trasladan.

**Qué hacer:** Revisar esas celdas en el libro anterior: si son el AC o el BAC de un paquete, cargarlos en sus filas de ingreso del libro nuevo.

---

### 4.13 DEP · Depuración de costes

Hoja PV-depuración y su traslado con `--migrar-desde`.

| Código | Severidad | Descripción |
|---|---|---|
| [DEP-001](#dep-001) | Info | Depuración de costes activa |
| [DEP-002](#dep-002) | Advertencia | Filas con código EDT no numérico |
| [DEP-003](#dep-003) | Advertencia | Códigos EDT repetidos en el libro |
| [DEP-004](#dep-004) | Info | Previsión no trasladada: libro sin depuración |
| [DEP-005](#dep-005) | Advertencia | Previsión no trasladada: otra empresa |
| [DEP-006](#dep-006) | Info | Previsión trasladada |
| [DEP-007](#dep-007) | Info | Previsión pasada al formato con EDT interna |
| [DEP-008](#dep-008) | Advertencia o Info | Meses de la previsión desplazados |

#### DEP-001

**Depuración de costes activa** · Info · generar con depuracion_costes (hoja Validación) · se puede silenciar

Recordatorio: la previsión se pega en la hoja PV-depuración y el libro de entrega se genera sin depuración.

**Qué hacer:** Ninguna.

#### DEP-002

**Filas con código EDT no numérico** · Advertencia · generar (hoja Validación) y validar, con depuracion_costes · se puede silenciar

Esas filas no pueden recibir la previsión de la hoja PV-depuración, que se asigna por código numérico.

**Qué hacer:** Usar códigos EDT numéricos.

#### DEP-003

**Códigos EDT repetidos en el libro** · Advertencia · generar (hoja Validación) y validar, con depuracion_costes · se puede silenciar

La previsión de un código repetido se asigna a la primera fila con ese código.

**Qué hacer:** Corregir los códigos repetidos (ver [EDT-009](#edt-009)).

#### DEP-004

**Previsión no trasladada: libro sin depuración** · Info · generar `--migrar-desde` · se puede silenciar

El libro nuevo no tiene depuración de costes: la hoja PV-depuración no se trasladó.

**Qué hacer:** Ninguna (es lo habitual en el libro de entrega).

#### DEP-005

**Previsión no trasladada: otra empresa** · Advertencia · generar con depuracion_costes (hoja Validación) · se puede silenciar

El libro anterior tiene la previsión de una empresa y el nuevo depura otra.

**Qué hacer:** Pegar la previsión de nuevo.

#### DEP-006

**Previsión trasladada** · Info · generar con depuracion_costes (hoja Validación) · se puede silenciar

Se conservó el contenido de la hoja PV-depuración.

**Qué hacer:** Ninguna.

#### DEP-007

**Previsión pasada al formato con EDT interna** · Info · generar con depuracion_costes (hoja Validación) · se puede silenciar

La hoja PV-depuración del libro anterior no tenía las columnas de EDT interna: los códigos numéricos quedaron como EDT cliente y los demás, como EDT interna.

**Qué hacer:** Revisar la hoja PV-depuración.

#### DEP-008

**Meses de la previsión desplazados** · Advertencia o Info · generar con depuracion_costes (hoja Validación) · se puede silenciar

Cambió el mes 1 del horizonte: los meses de la hoja PV-depuración se desplazaron para conservar su fecha. Es una advertencia si quedaron importes fuera de las columnas.

**Qué hacer:** Revisar los importes que quedaron fuera, si los hay.

---

### 4.14 ACT · Actualizar y analizar

Carga del % ganado del mes: fecha de corte, filas del libro, BAC y observaciones por fila.

| Código | Severidad | Descripción |
|---|---|---|
| [ACT-001](#act-001) | Error (detiene el programa) | El libro es de un solo archivo |
| [ACT-002](#act-002) | Error (detiene el programa) | El libro es de varios archivos |
| [ACT-003](#act-003) | Error (detiene el programa) | Los archivos no coinciden con los del libro |
| [ACT-004](#act-004) | Error (detiene el programa) | Sin fecha de estado |
| [ACT-005](#act-005) | Error (detiene el programa) | Fechas de estado distintas |
| [ACT-006](#act-006) | Error (detiene el programa) | Fecha de corte fuera del horizonte |
| [ACT-007](#act-007) | Error (detiene el programa) | Valores en meses posteriores a la fecha de corte |
| [ACT-008](#act-008) | Error (detiene el programa) | El libro no tiene fecha de corte |
| [ACT-009](#act-009) | Advertencia | Planificación distinta de la del libro |
| [ACT-010](#act-010) | Advertencia | Proyecto distinto del libro |
| [ACT-011](#act-011) | Advertencia | La tarea de la fila ya no existe |
| [ACT-012](#act-012) | Advertencia | La tarea de la fila está inactiva |
| [ACT-013](#act-013) | Advertencia | No se encontró la tarea de la fila |
| [ACT-014](#act-014) | Advertencia | El campo Visible cambió |
| [ACT-015](#act-015) | Advertencia | Tareas nuevas sin fila en el libro |
| [ACT-016](#act-016) | Advertencia | lista_edt en un libro generado sin lista |
| [ACT-017](#act-017) | Advertencia | La lista EDT cambió |
| [ACT-018](#act-018) | Advertencia | Fecha de corte que no es fin de mes |
| [ACT-019](#act-019) | Advertencia | Fecha de estado distinta de la fecha de corte |
| [ACT-020](#act-020) | Advertencia | Fecha de corte anterior a la última cargada |
| [ACT-021](#act-021) | Advertencia | `--forzar` con meses posteriores cargados |
| [ACT-022](#act-022) | Advertencia | Ninguna fila tiene BAC |
| [ACT-023](#act-023) | Advertencia | Filas con % ganado sin BAC |
| [ACT-024](#act-024) | Advertencia | Celdas sin la fórmula del libro |
| [ACT-025](#act-025) | Advertencia | Celdas de ingreso convertidas en número |
| [ACT-026](#act-026) | Info | El AC se carga a mano |
| [ACT-027](#act-027) | Info | % ganado del mes negativo |
| [ACT-028](#act-028) | Info | La empresa no participa en las tareas de la fila |
| [ACT-029](#act-029) | Info | Hitos con pesos negativos ponderados por igual |
| [ACT-030](#act-030) | Info | Hitos sin peso ponderados por igual |
| [ACT-031](#act-031) | Info | Hitos con peso negativo tomados con peso 0 |
| [ACT-032](#act-032) | Info | Sin hitos: % físico completado |
| [ACT-033](#act-033) | Info | Sin hitos ni % físico: % completado |
| [ACT-034](#act-034) | Info | Método no definido: % completado |
| [ACT-035](#act-035) | Info | Método no reconocido: % completado |
| [ACT-036](#act-036) | Advertencia o Info | Cambio de BAC |
| [ACT-037](#act-037) | Advertencia | Valores de meses ya informados modificados |
| [ACT-038](#act-038) | Info | Referencia del registro de cambios guardada |
| [ACT-039](#act-039) | Advertencia | Avance sin costo |
| [ACT-040](#act-040) | Advertencia | Costo sin avance |
| [ACT-041](#act-041) | Advertencia | Costo después de terminar el paquete |
| [ACT-042](#act-042) | Info | AC del mes de corte sin cargar |
| [ACT-043](#act-043) | Advertencia | Análisis del mes anterior incompleto |
| [ACT-044](#act-044) | Advertencia | Desviaciones del mes anterior sin analizar |
| [ACT-045](#act-045) | Advertencia | Acción vencida con el paquete fuera de umbral |
| [ACT-046](#act-046) | Advertencia | Meses sin informe |
| [ACT-047](#act-047) | Advertencia | Línea base de MS Project distinta de la del libro |
| [ACT-048](#act-048) | Advertencia | EAC poco realista |
| [ACT-049](#act-049) | Advertencia | Mucho BAC en filas LOE |

#### ACT-001

**El libro es de un solo archivo** · Error (detiene el programa) · actualizar · no se puede silenciar

El libro se generó con un solo archivo de MS Project y la configuración tiene `[[proyectos]]`.

**Qué hacer:** Indicar ese archivo con `--origen`, con una configuración sin `[[proyectos]]`.

#### ACT-002

**El libro es de varios archivos** · Error (detiene el programa) · actualizar · no se puede silenciar

El libro se generó con varios archivos de MS Project y se indicó uno solo.

**Qué hacer:** Indicar los archivos en secciones `[[proyectos]]`, con los mismos nombres que al generar el libro.

#### ACT-003

**Los archivos no coinciden con los del libro** · Error (detiene el programa) · actualizar · no se puede silenciar

Los nombres de `[[proyectos]]` no son los del libro.

**Qué hacer:** Indicar todos los archivos del libro, con los mismos nombres. Para agregar o quitar un archivo, generar un libro nuevo con `--migrar-desde`.

#### ACT-004

**Sin fecha de estado** · Error (detiene el programa) · actualizar · no se puede silenciar

El proyecto (o algún archivo) no tiene fecha de estado y no se indicó `--fecha-corte`.

**Qué hacer:** Indicar `--fecha-corte` AAAA-MM-DD o definir Proyecto > Fecha de estado en MS Project.

#### ACT-005

**Fechas de estado distintas** · Error (detiene el programa) · actualizar · no se puede silenciar

Con varios archivos, las fechas de estado no coinciden.

**Qué hacer:** Indicar `--fecha-corte` o definir la misma fecha de estado en todos los archivos.

#### ACT-006

**Fecha de corte fuera del horizonte** · Error (detiene el programa) · actualizar y analizar · no se puede silenciar

La fecha de corte no está entre el primer y el último mes del libro.

**Qué hacer:** Revisar la fecha, o regenerar el libro con un horizonte mayor (`--migrar-desde`).

#### ACT-007

**Valores en meses posteriores a la fecha de corte** · Error (detiene el programa) · actualizar · no se puede silenciar

El libro tiene % ganado cargado después del mes de la fecha de corte.

**Qué hacer:** Borrar esos valores, o usar `--forzar` si se quiere corregir un mes anterior (ver [ACT-021](#act-021)).

#### ACT-008

**El libro no tiene fecha de corte** · Error (detiene el programa) · analizar · no se puede silenciar

analizar necesita la fecha de corte de la hoja Resumen, que escribe actualizar.

**Qué hacer:** Ejecutar actualizar antes de analizar.

#### ACT-009

**Planificación distinta de la del libro** · Advertencia · actualizar · se puede silenciar

El libro se generó para otra planificación: se usa la del libro.

**Qué hacer:** Revisar planificacion en la configuración.

#### ACT-010

**Proyecto distinto del libro** · Advertencia · actualizar · se puede silenciar

El libro se generó para otro proyecto de MS Project: las filas se asocian por código EDT.

**Qué hacer:** Verificar los resultados.

#### ACT-011

**La tarea de la fila ya no existe** · Advertencia · actualizar · se puede silenciar

La fila se omite.

**Qué hacer:** Si cambió la EDT, regenerar el libro con `--migrar-desde`.

#### ACT-012

**La tarea de la fila está inactiva** · Advertencia · actualizar · se puede silenciar

La fila se omite y se conservan los valores ya cargados.

**Qué hacer:** Si quedó fuera del alcance, regenerar el libro con `--migrar-desde`.

#### ACT-013

**No se encontró la tarea de la fila** · Advertencia · actualizar · se puede silenciar

Con un proyecto distinto del libro, ninguna tarea tiene el código EDT de la fila: se omite.

**Qué hacer:** Revisar el código EDT de la tarea.

#### ACT-014

**El campo Visible cambió** · Advertencia · actualizar · se puede silenciar

El campo Visible de algunas tareas no coincide con las filas del libro: el libro conserva sus filas.

**Qué hacer:** Para aplicar el cambio, regenerar el libro con `--migrar-desde`.

#### ACT-015

**Tareas nuevas sin fila en el libro** · Advertencia · actualizar · se puede silenciar

Se agregaron tareas en MS Project después de generar el libro: su avance no se incluye en el % ganado.

**Qué hacer:** Para incluirlas, regenerar el libro con `--migrar-desde`.

#### ACT-016

**lista_edt en un libro generado sin lista** · Advertencia · actualizar · se puede silenciar

La configuración indica lista_edt, pero el libro se generó sin lista EDT: se mantienen sus filas.

**Qué hacer:** Para aplicar la lista, regenerar el libro con `--migrar-desde`.

#### ACT-017

**La lista EDT cambió** · Advertencia · actualizar · se puede silenciar

La lista EDT de la configuración cambió después de generar el libro: se usa la lista con la que se generó.

**Qué hacer:** Para aplicar los cambios, regenerar el libro con `--migrar-desde`.

#### ACT-018

**Fecha de corte que no es fin de mes** · Advertencia · actualizar · se puede silenciar

El PV acumulado incluye el mes completo.

**Qué hacer:** Usar el último día del mes como fecha de corte.

#### ACT-019

**Fecha de estado distinta de la fecha de corte** · Advertencia · actualizar · se puede silenciar

El avance registrado en MS Project debe corresponder a la fecha de corte.

**Qué hacer:** Usar la fecha de estado de MS Project como fecha de corte, o actualizar el avance a esa fecha.

#### ACT-020

**Fecha de corte anterior a la última cargada** · Advertencia · actualizar · se puede silenciar

Se carga un mes anterior al último cargado en el libro.

**Qué hacer:** Verificar que sea intencional (p. ej. una corrección).

#### ACT-021

**`--forzar` con meses posteriores cargados** · Advertencia · actualizar · se puede silenciar

Los valores de los meses posteriores se calcularon con el acumulado anterior: ya no coinciden con MS Project.

**Qué hacer:** Revisar esos meses o volver a cargarlos con actualizar, en orden.

#### ACT-022

**Ninguna fila tiene BAC** · Advertencia · actualizar · se puede silenciar

El PV, el EV y los indicadores quedan en 0 hasta ingresar el BAC.

**Qué hacer:** Ingresar el BAC de cada fila en la columna D de la hoja PV.

#### ACT-023

**Filas con % ganado sin BAC** · Advertencia · actualizar · se puede silenciar

El EV de esas filas no se puede valorar.

**Qué hacer:** Ingresar su BAC en la columna D de la hoja PV.

#### ACT-024

**Celdas sin la fórmula del libro** · Advertencia · actualizar · se puede silenciar

Subtotales, la fila TOTAL o la suma de las empresas tienen un valor en lugar de su fórmula: los totales del libro no coinciden con los que muestra la consola.

**Qué hacer:** Restaurar las fórmulas (p. ej. copiándolas de una celda vecina de la misma fila) o recuperar el libro desde un respaldo.

#### ACT-025

**Celdas de ingreso convertidas en número** · Advertencia · actualizar y analizar · se puede silenciar

Celdas de ingreso con un número guardado como texto o con una fórmula se reemplazaron por su valor.

**Qué hacer:** Ninguna; se recomienda ingresar números.

#### ACT-026

**El AC se carga a mano** · Info · actualizar · se puede silenciar

Recordatorio: los indicadores de costo usan el AC cargado en las hojas AC.

**Qué hacer:** Cargar el AC del mes.

#### ACT-027

**% ganado del mes negativo** · Info · actualizar (observaciones por fila) · se puede silenciar

El % ganado acumulado bajó respecto del mes anterior: el mes tiene un valor negativo (corrección).

**Qué hacer:** Verificar el avance registrado en MS Project.

#### ACT-028

**La empresa no participa en las tareas de la fila** · Info · actualizar (observaciones por fila) · se puede silenciar

Se usó el avance de la fila completa para el % ganado de esa empresa.

**Qué hacer:** Ninguna.

#### ACT-029

**Hitos con pesos negativos ponderados por igual** · Info · actualizar (observaciones por fila) · se puede silenciar

Método «Hitos ponderados»: ningún hito tiene un peso válido (pesos negativos): se ponderaron por igual.

**Qué hacer:** Corregir el peso de los hitos.

#### ACT-030

**Hitos sin peso ponderados por igual** · Info · actualizar (observaciones por fila) · se puede silenciar

Método «Hitos ponderados»: ningún hito tiene peso: se ponderaron por igual.

**Qué hacer:** Indicar el peso de los hitos, si no deben pesar igual.

#### ACT-031

**Hitos con peso negativo tomados con peso 0** · Info · actualizar (observaciones por fila) · se puede silenciar

Método «Hitos ponderados»: algunos hitos tienen peso negativo.

**Qué hacer:** Corregir el peso de esos hitos.

#### ACT-032

**Sin hitos: % físico completado** · Info · actualizar (observaciones por fila) · se puede silenciar

Método «Hitos ponderados» sin hitos: se usó el % físico completado.

**Qué hacer:** Agregar los hitos del paquete o cambiar el método.

#### ACT-033

**Sin hitos ni % físico: % completado** · Info · actualizar (observaciones por fila) · se puede silenciar

Método «Hitos ponderados» sin hitos ni % físico completado: se usó el % completado.

**Qué hacer:** Agregar los hitos del paquete o cambiar el método.

#### ACT-034

**Método no definido: % completado** · Info · actualizar (observaciones por fila) · se puede silenciar

La fila no tiene método de medición: se usó el % completado.

**Qué hacer:** Indicar el método en la columna C del libro o en MS Project.

#### ACT-035

**Método no reconocido: % completado** · Info · actualizar (observaciones por fila) · se puede silenciar

El método de la fila no es uno de los admitidos: se usó el % completado (ver [MET-001](#met-001)).

**Qué hacer:** Corregir el método en la columna C del libro o en MS Project.

#### ACT-036

**Cambio de BAC** · Advertencia o Info · actualizar y analizar · se puede silenciar

El BAC de la fila cambió desde la ejecución anterior: el cambio se agregó a la hoja Registro de cambios. Es una advertencia si la fila tiene meses ya informados: su PV y su EV se recalcularon con el BAC nuevo (cambio retroactivo).

**Qué hacer:** Indicar el motivo y el número de cambio aprobado en la hoja Registro de cambios.

#### ACT-037

**Valores de meses ya informados modificados** · Advertencia · actualizar y analizar · se puede silenciar

El % ganado o el AC de meses anteriores al mes de corte cambió desde la ejecución anterior, o se cargó el AC de un mes anterior al último informado: el cambio se agregó a la hoja Registro de cambios.

**Qué hacer:** Si es una corrección, indicar el motivo en la hoja Registro de cambios; si no, restaurar el valor anterior.

#### ACT-038

**Referencia del registro de cambios guardada** · Info · actualizar y analizar · se puede silenciar

Primera ejecución con el registro de cambios: se guardaron el BAC y los valores de los meses informados de cada fila. Los cambios se registran a partir de la ejecución siguiente.

**Qué hacer:** Ninguna.

#### ACT-039

**Avance sin costo** · Advertencia · actualizar y analizar · se puede silenciar

La fila tiene EV acumulado y no tiene AC acumulado: el costo real no se cargó o se imputó a otro paquete, y el CPI no se puede calcular.

**Qué hacer:** Cargar el AC del paquete en la hoja AC o revisar el avance informado.

#### ACT-040

**Costo sin avance** · Advertencia · actualizar y analizar · se puede silenciar

La fila tiene AC acumulado y no tiene EV acumulado: costo sin trabajo medido. No se revisa en los paquetes con método 0/100 o «Hitos ponderados», en los que el EV llega al terminar o al cumplirse cada hito.

**Qué hacer:** Revisar el avance en MS Project, el método de medición o la imputación del AC.

#### ACT-041

**Costo después de terminar el paquete** · Advertencia · actualizar y analizar · se puede silenciar

La fila tiene AC en meses posteriores al mes en que su % ganado llegó al 100 %.

**Qué hacer:** Revisar si es un costo tardío del paquete (y explicarlo) o un costo de otro paquete.

#### ACT-042

**AC del mes de corte sin cargar** · Info · actualizar y analizar · se puede silenciar

Los controles del EV y el AC se revisan hasta el último mes con AC cargado (o no se revisan, si ningún mes tiene AC).

**Qué hacer:** Cargar el AC del mes en la hoja AC y ejecutar analizar.

#### ACT-043

**Análisis del mes anterior incompleto** · Advertencia · actualizar y analizar · se puede silenciar

La fila del último análisis anterior al mes de corte tiene texto, pero no tiene Causa o no tiene Acción correctiva.

**Qué hacer:** Completar la causa y la acción correctiva en la hoja Análisis de desviaciones.

#### ACT-044

**Desviaciones del mes anterior sin analizar** · Advertencia · actualizar y analizar · se puede silenciar

Paquetes que estaban fuera de umbral en el último análisis anterior al mes de corte y cuya fila no tiene ningún texto: se quitan de la hoja. Los que siguen fuera de umbral vuelven a listarse en el mes de corte.

**Qué hacer:** Analizar las desviaciones de cada mes antes del cierre siguiente.

#### ACT-045

**Acción vencida con el paquete fuera de umbral** · Advertencia · actualizar y analizar · se puede silenciar

El paquete sigue fuera de umbral en el mes de corte y la Fecha compromiso de su acción correctiva es anterior a la fecha de corte: la acción no se cumplió o no fue eficaz.

**Qué hacer:** Revisar la acción correctiva y su fecha compromiso en la hoja Análisis de desviaciones.

#### ACT-046

**Meses sin informe** · Advertencia · actualizar y analizar · se puede silenciar

Entre el primer mes con % ganado y el mes de corte hay meses sin % ganado cargado: su avance quedó en el mes cargado siguiente. La norma EIA-748-E pide la información al menos cada mes.

**Qué hacer:** Ejecutar actualizar al cierre de cada mes, con la fecha de estado de ese mes, aunque el informe sea bimestral. Si la frecuencia está acordada con el cliente, silenciar [ACT-046](#act-046).

#### ACT-047

**Línea base de MS Project distinta de la del libro** · Advertencia · actualizar · se puede silenciar

El % planificado que resulta hoy de MS Project no coincide con el de las hojas % PV, que calculó generar: se guardó de nuevo la línea base, o cambiaron las tareas o el calendario de la fila. El libro sigue con la línea base anterior. No se comparan las filas con tareas sin fechas de línea base.

**Qué hacer:** Si el cambio de línea base está aprobado, generar un libro nuevo con `--migrar-desde` (conserva el BAC, el % ganado y el AC); si no, restaurar la línea base en MS Project.

#### ACT-048

**EAC poco realista** · Advertencia · actualizar y analizar · se puede silenciar

El TCPI(BAC) supera al CPI en más de 0,10, o el AC ya alcanzó el BAC con trabajo pendiente: terminar dentro del BAC exige una eficiencia de costo muy superior a la lograda. Se calcula con el proyecto (cada empresa en multi-empresa) en el último mes con AC cargado, desde el 20 % completado: antes, el CPI acumulado todavía no es representativo.

**Qué hacer:** Revisar el EAC y, si el desvío se confirma, evaluar una línea base por encima del objetivo (OTB) con el cliente.

#### ACT-049

**Mucho BAC en filas LOE** · Advertencia · actualizar y analizar · se puede silenciar

El % del BAC total en filas con método LOE supera `[validacion] loe_maximo`: su EV es igual a su PV y acerca el SPI a 1 aunque el trabajo medible se atrase. La norma EIA-748-E pide mantener el LOE al mínimo.

**Qué hacer:** Medir con hitos o % completado el trabajo que lo permita, o ajustar loe_maximo si el contrato lo admite.

---

### 4.15 FIC · Fichas de tareas

Comando generar-fichas.

| Código | Severidad | Descripción |
|---|---|---|
| [FIC-001](#fic-001) | Advertencia | No hay fichas |
| [FIC-002](#fic-002) | Info | Vínculos con tareas sin ficha |
| [FIC-003](#fic-003) | Advertencia | Fichas sin línea base |
| [FIC-004](#fic-004) | Advertencia | El archivo tiene errores |

#### FIC-001

**No hay fichas** · Advertencia · generar-fichas · se puede silenciar

Ninguna tarea tiene código EDT.

**Qué hacer:** Revisar el campo EDT configurado (`[campos] edt`) y la planificación elegida.

#### FIC-002

**Vínculos con tareas sin ficha** · Info · generar-fichas · se puede silenciar

Los vínculos con tareas de otra planificación, inactivas, con Visible = No, fuera de la lista EDT o de otro proyecto no se muestran.

**Qué hacer:** Ninguna.

#### FIC-003

**Fichas sin línea base** · Advertencia · generar-fichas · se puede silenciar

Se usaron las fechas y la duración del programa actual (se indica en la ficha).

**Qué hacer:** Guardar la línea base de esas tareas.

#### FIC-004

**El archivo tiene errores** · Advertencia · generar-fichas · se puede silenciar

El cronograma tiene errores (por ejemplo, códigos EDT duplicados) que pueden afectar las fichas.

**Qué hacer:** Revisarlo con validar.

---

### 4.16 SIS · Errores del sistema

Errores de archivos del sistema y errores inesperados del programa.

| Código | Severidad | Descripción |
|---|---|---|
| [SIS-001](#sis-001) | Error (detiene el programa) | Error inesperado |
| [SIS-002](#sis-002) | Error (detiene el programa) | No se pudo leer o escribir un archivo |
| [SIS-003](#sis-003) | Error (detiene el programa) | Error sin código específico |

#### SIS-001

**Error inesperado** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Un error no previsto del programa (código de salida 3).

**Qué hacer:** Volver a ejecutar con -v para ver el detalle técnico e informar el problema.

#### SIS-002

**No se pudo leer o escribir un archivo** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Error del sistema de archivos: disco lleno, unidad de red no disponible, ruta no válida...

**Qué hacer:** Revisar el espacio en disco, la ruta y los permisos.

#### SIS-003

**Error sin código específico** · Error (detiene el programa) · Todos los comandos · no se puede silenciar

Un error previsto (archivo, permisos o datos) sin un código propio en este catálogo.

**Qué hacer:** Leer el mensaje; si no es claro, volver a ejecutar con -v e informar el problema.
