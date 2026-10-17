# Validaciones, limitaciones y solución de problemas

[Inicio](../README.md) › [Documentación](README.md) › Validaciones y solución de problemas

---

## Contenido

1. [Validaciones](#1-validaciones)
2. [Limitaciones conocidas](#2-limitaciones-conocidas)
3. [Solución de problemas](#3-solución-de-problemas)

---

## 1. Validaciones

`generar` informa en la hoja **Validación** (y en la consola) las observaciones de la exportación, con
su severidad: `Info` (informativo), `Advertencia` (se debe revisar) o `Error` (dato que no se pudo
trasladar o calcular).

![Hoja Validación del proyecto de ejemplo](img/excel/validacion.png)

| Observación | Qué revisar |
|---|---|
| Planificación exportada y cantidad de tareas que dependen de ella | Que sea la planificación correcta. |
| Tareas sin código EDT omitidas (con su costo) | Tareas que no dependen de ningún elemento EDT: si tienen costo, asignarles código EDT o moverlas bajo un elemento. |
| Tareas sin código EDT o de nivel mayor al máximo sumadas en su elemento superior | Informativo: es el comportamiento esperado para subtareas y hitos de un paquete. |
| Proyecto sin línea base: se usó el programa actual | Guardar la línea base antes de informar ([Presupuesto](presupuesto-bac.md#6-guardar-la-línea-base)). |
| Tareas con costo pero sin línea base | Tareas agregadas después de guardar la línea base: no aportan PV. Si están aprobadas, guardar la línea base de esas tareas. |
| Tareas resumen con costo propio | Se agrega una fila `-D Otros costos`. Se recomienda mover el costo a una tarea de detalle. |
| Línea base de una tarea resumen distinta de la suma de sus subtareas | La línea base no se acumuló en las tareas resumen: guardar la línea base con *Resumir líneas base: en todas las tareas de resumen*. |
| Distribución mensual de MS Project que no suma el costo de línea base | Costo de línea base editado a mano: se ajusta en forma proporcional. |
| Tareas repartidas en forma lineal | Tareas sin distribución mensual en MS Project. |
| PV anterior al mes 1 o posterior al último mes | Revisar `mes_inicio` y `meses`. |
| BAC total distinto del costo de línea base del proyecto o de la planificación | Costos fuera de la EDT exportada, tareas inactivas o línea base no acumulada. |
| Códigos EDT duplicados | Por ejemplo, si no se eligió una planificación y ambas usan los mismos códigos. |
| Filas sin método de medición | Completar el método en MS Project o en la columna C de la hoja `PV`. |
| Tareas inactivas excluidas | Informativo (con `incluir_inactivas = false`). |
| Multi-empresa: BAC de cada empresa, tareas sin empresa, valores no reconocidos y `% Reparto` fuera de rango | Completar la columna `Empresa` y revisar los repartos ([Multi-empresa](multiempresa.md#7-validaciones-y-controles)). |
| `--migrar-desde`: historial no trasladado (`Error`) | Cargar a mano el EV y el AC de esas filas en el libro nuevo. |

`actualizar` informa sus controles en pantalla ([Cálculo de EV y AC](calculo-ev-ac.md#7-controles-de-consistencia)).

---

## 2. Limitaciones conocidas

- El modo `.mpp` requiere MS Project de escritorio en Windows. Sin él, se debe usar el XML.
- El programa no guarda cambios en el archivo de MS Project: lo abre en modo solo lectura.
- **Proyectos maestros con subproyectos insertados:** el grupo de recursos se abre en solo lectura; se
  debe verificar en `Validación` que el BAC total coincida con el del proyecto maestro.
- **Motor XML:** los tramos que cruzan de mes se reparten con el calendario del proyecto. No se usan los
  calendarios de recursos ni de tareas, y se ignoran las excepciones recurrentes relativas (por ejemplo,
  "primer lunes del mes" o "último viernes de marzo"). Con `.mpp` no hay esta limitación.
- Si la curva PV no alcanza el BAC dentro del horizonte, la duración planificada y el Cronograma Ganado
  quedan vacíos; en ese caso se debe aumentar `meses`.
- `% Completado` aplica el porcentaje sobre el BAC, a diferencia del CPTR de MS Project
  ([Cálculo de EV y AC](calculo-ev-ac.md#33-diferencia-con-el-valor-acumulado-de-ms-project)).
- MS Project no tiene métodos nativos 0/100, 50/50 ni de hitos ponderados: el programa los calcula a
  partir del avance registrado.
- El libro está pensado para Excel. LibreOffice abre las fórmulas correctamente, pero dibuja como cero
  los meses sin datos de EV y AC en los gráficos.
- El lector `.mpp` se probó con objetos que simulan MS Project. En el primer uso con un proyecto real se
  recomienda comparar el BAC total y el PV de algunos meses con la vista *Uso de tareas*
  ([Presupuesto](presupuesto-bac.md#7-verificar-el-bac-y-el-pv-en-ms-project)).

---

## 3. Solución de problemas

| Mensaje o síntoma | Solución |
|---|---|
| `No se pudo iniciar MS Project por automatización COM` | Se debe verificar que MS Project de escritorio esté instalado. Alternativa: usar el XML. |
| `Para leer archivos .mpp se necesita pywin32` | Se debe instalar con `python -m pip install pywin32`. |
| Error de atributos de `win32com` o `gen_py` | Se debe borrar la carpeta `%TEMP%\gen_py` y volver a ejecutar. |
| MS Project muestra un diálogo y el programa no avanza | Se recomienda ejecutar con `--visible` para ver el diálogo (inicio de sesión, archivo protegido con contraseña, etc.). |
| Con *Project Online Desktop Client*, MS Project pide elegir una cuenta o un perfil al iniciarse | Se debe abrir MS Project manualmente, elegir la cuenta (por ejemplo *Equipo* para trabajar con archivos locales) y luego ejecutar el programa, que usa esa instancia abierta y no la cierra. |
| El proyecto está en Project Online, SharePoint u OneDrive (se abre desde una dirección web) | Se abre en MS Project y se usa `--origen activo`, o se guarda una copia local `.mpp` y se usa esa ruta. |
| El `.mpp` se descargó de Internet o del correo y MS Project no lo abre | Windows puede haberlo bloqueado: clic derecho en el archivo > *Propiedades* > *Desbloquear*. |
| `No se pudo guardar ...` | Si el libro está abierto en Excel, se debe cerrar y repetir el comando. |
| `Hay valores cargados en meses posteriores a la fecha de corte` | Se deben borrar esos meses en `EV`/`AC` o usar `--forzar`. |
| `El proyecto no tiene fecha de estado` | Se debe definir la fecha de estado en MS Project o indicar `--fecha-corte`. |
| `contiene el número ... guardado como texto` | Se debe escribir ese valor como número en Excel, o ejecutar `actualizar`, que lo convierte. |
| `La celda ... tiene una fórmula sin valor calculado` | Se debe abrir y guardar el libro en Excel, o reemplazar la fórmula de esa celda de ingreso por un número. |
| La curva no muestra EV/AC | Se debe completar la fecha de corte en `Resumen!B5`. |
| `No se encontró la planificación ...` | El mensaje lista las tareas de nivel superior con su código EDT; se debe indicar en `planificacion` uno de esos códigos. |
| El libro solo muestra los elementos de primer nivel (`1.0.0`, `2.0.0`...) | Se debe revisar que el campo indicado en `[campos] edt` contenga los códigos EDT de todas las tareas. El nivel se toma del código EDT. |
| Todas las filas tienen método `Por definir` | Se debe revisar `[campos] metodo` (nombre del campo o su alias) y que el campo esté completo en las tareas de nivel 3. |
| El EV de una fila con `Hitos ponderados` no avanza | Los hitos deben ser subtareas de la fila con duración 0 y marcarse al 100 % al cumplirse ([Métodos de medición](metodos-de-medicion.md#33-hitos-ponderados)). |
| El AC del libro difiere del costo real de MS Project | Revisar las advertencias de tareas con costo real sin fila y de costos en tareas resumen. |
| `En modo multi-empresa el AC se carga a mano ...` | El AC se carga en las hojas de cada empresa. Para cargarlo desde MS Project se debe indicar `ac_desde_project = true` en `[multiempresa]`. |
| `--migrar-desde no se admite entre un libro de una sola empresa y uno multi-empresa` | Se debe generar el libro nuevo sin `--migrar-desde` y cargar a mano el historial que se necesite. |
| `No se encontró el campo de empresa ...` | Se debe revisar `campo_empresa` en `[multiempresa]`: nombre del campo (`Text3`) o su alias en MS Project. |
| Los indicadores aparecen vacíos al abrir el libro | El libro se abrió en vista protegida o en una vista previa: se debe habilitar la edición para que Excel calcule las fórmulas. |
| Las tildes se ven mal en la consola | Solo afecta a la consola, no al libro. Se recomienda usar Windows Terminal o ejecutar `chcp 65001`. |
