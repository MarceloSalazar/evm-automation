# Registro de cambios

Cambios de cada versión del Exportador EVM, de la más reciente a la más antigua. Cada versión tiene su
etiqueta `vX.Y.Z` en el repositorio y su publicación (*release*) en
[GitHub](https://github.com/MarceloSalazar/evm-automation/releases) con estas mismas notas.

Las versiones siguen el formato MAYOR.MENOR.CORRECCIÓN ([versionado semántico](https://semver.org/lang/es/)):
MENOR para funciones nuevas que no impiden usar los libros ya generados, CORRECCIÓN para correcciones, y
MAYOR para cambios que exigen regenerar los libros o cambiar la configuración. La versión instalada se
consulta con `python evm_export.py --version` (con el instalador de Windows, `evm_export --version`).

Los cambios incorporados a `main` que todavía no forman parte de una versión se anotan en
*Sin publicar*; al publicar la versión siguiente, esa sección pasa a ser su entrada.

---

## [Sin publicar]

### Nuevas funciones

- **Depuración de costes con EDT cliente y EDT interna**: la hoja `PV-depuración` sigue el plan de
  costes de la empresa en dos niveles: A `EDT cliente`, B `Actividad`, C `EDT interna` (conceptos de
  coste internos, con códigos no numéricos como `ABC01`), D `Actividad interna`, E `BAC` y los meses
  desde la columna F. Cada línea de EDT interna se suma en su EDT cliente, la de la línea de arriba o la
  repetida en la columna A; la línea de la EDT cliente es un subtotal y no se suma. Los controles pasan a
  la columna E, y `Asignado a` y `Observación` a las columnas DW y DX, con una observación nueva para las
  líneas de EDT interna sin EDT cliente encima
  ([Depuración de costes](docs/depuracion-costes.md#4-hoja-pv-depuración-la-previsión-pegada)).
- `--migrar-desde` pasa al formato nuevo la previsión pegada en un libro de depuración anterior: los
  códigos numéricos quedan como EDT cliente y los demás (`XRF01`), como EDT interna. `actualizar` y
  `analizar` siguen funcionando con los libros de depuración anteriores.
- **Códigos de errores y advertencias**: cada mensaje del programa tiene un código `ÁREA-NNN` (por
  ejemplo `FEC-003`), en la consola (`[!] FEC-007 1.2.3 Montaje: ...`, `ERROR CFG-001: ...`) y en la
  nueva columna `Código` de la hoja `Validación`. El documento nuevo
  [Códigos de errores y advertencias](docs/codigos.md) explica cada código (severidad, causa, qué hacer y
  comandos en que aparece).
- **Silenciar observaciones**: las observaciones ya revisadas se ocultan en la consola y en la hoja
  `Validación` con `[validacion] silenciar = ["LOG-004", "MET-004:1.2.1"]` o con la opción
  `--silenciar` de cualquier comando (un código, o un código para una sola fila o tarea). Siempre se
  informa cuántas se silenciaron (`CFG-049`); los errores silenciados no cuentan para el código de salida
  de `validar` y los errores que detienen el programa no se pueden silenciar
  ([Silenciar observaciones](docs/codigos.md#2-silenciar-observaciones)).
- **Evaluación DCMA en `validar`**: nuevos controles del cronograma, con los criterios de la evaluación
  DCMA de 14 puntos. `validar` revisa ahora 12 de los 14 puntos:
  - holgura alta: tareas sin terminar con más de 44 días laborables de holgura total (`LOG-013`;
    advertencia si superan el 5 %);
  - holgura negativa (`LOG-014` y `LOG-015`, con la holgura de cada tarea);
  - duración alta: tareas sin terminar con más de 44 días laborables de duración en la línea base
    (`LOG-016`; advertencia si superan el 5 %). La holgura alta y la duración alta no cuentan las filas
    LOE;
  - tareas incumplidas: las que debían terminar según la línea base y no terminaron, o terminaron
    después de su fin de línea base (`AVA-014`; advertencia si superan el 5 %);
  - índice de longitud del camino crítico (CPLI): días laborables que quedan hasta el fin previsto, más la
    holgura frente al fin de la línea base, divididos por los días que quedan (`AVA-015`; advertencia si
    es menor que 0,95).
- **Avance fuera de secuencia** en `validar` (`AVA-016` y `AVA-017`): tareas que comenzaron (o
  terminaron) y cuya predecesora no terminó (vínculo FC o FF) o no comenzó (CC o CF).
- Documento nuevo [Evaluación DCMA de 14 puntos](docs/evaluacion-dcma.md): qué mide cada punto, cómo lo
  calcula `validar`, qué hacer en MS Project, la prueba manual del camino crítico (punto 12) y lo que el
  programa no verifica (los recursos, punto 10).
- **Registro de cambios** (#24, #31): `actualizar` y `analizar` comparan el libro con los valores de la
  ejecución anterior y agregan a la hoja nueva `Registro de cambios` cada cambio del BAC y del % ganado o
  el AC de un mes ya informado, con el valor anterior, el nuevo y columnas amarillas para el motivo y el
  número de cambio aprobado. Un cambio del BAC con meses informados es una advertencia (`ACT-036`): el PV
  y el EV de esos meses se recalculan (cambio retroactivo, guía 26 de la norma EIA-748-E). Los cambios de
  meses ya informados se avisan con `ACT-037`; cargar el AC del último mes informado no es un cambio. La
  primera ejecución solo guarda la referencia (`ACT-038`), `--simular` no registra y `--migrar-desde`
  traslada el registro y la referencia (`MIG-018`)
  ([El libro Excel](docs/libro-excel.md#registro-de-cambios)).
- **Coherencia del EV y el AC** (#32): `actualizar` y `analizar` avisan de los paquetes con avance y sin
  costo (`ACT-039`), con costo y sin avance (`ACT-040`, salvo los métodos 0/100 e hitos ponderados) y con
  costo después de llegar al 100 % (`ACT-041`), hasta el último mes con AC cargado (`ACT-042`).
- **Análisis de desviaciones incompleto y acciones vencidas** (#33): en el cierre siguiente se avisa de
  las filas del mes anterior sin causa o sin acción correctiva (`ACT-043`), de los paquetes que quedaron
  sin analizar (`ACT-044`) y de las acciones con la fecha compromiso vencida de los paquetes que siguen
  fuera de umbral (`ACT-045`)
  ([Análisis de desviaciones](docs/analisis-desviaciones.md#3-cierre-mensual)).
- **Meses sin informe** (#34): `actualizar` avisa de los meses sin % ganado entre el último mes cargado y
  el mes de corte, y `analizar`, de todos los meses sin % ganado desde el primero cargado (`ACT-046`).

### Correcciones

- Ninguna.

### Otros

- Documento nuevo [Norma EIA-748-E](docs/evaluacion-eia-748-e.md): las 27 guías de la norma de EE. UU.
  para sistemas EVM (revisión E, 2026), qué exige cada una, qué cubre el programa y con qué controles, y
  qué debe aportar la organización.
- El ayudante del ejemplo (`generar_ejemplo.py --libro`) escribe la previsión ficticia con líneas de EDT
  interna, y la captura de la hoja `PV-depuración` muestra el formato nuevo.
- Hoja `Validación`: la columna `Código` es la B; el código EDT, la tarea, el detalle y el tema pasan a
  las columnas C a F. En la consola, las observaciones informativas de `generar` y las de `actualizar`,
  `analizar` y `generar-fichas` se muestran con el mismo formato que las de `validar` (código, fila o
  tarea y mensaje).
- Metodología para agregar un mensaje con su código en
  [Desarrollo](docs/desarrollo.md#5-códigos-de-errores-y-advertencias): catálogo en `evm/messages.py`,
  `docs/codigos.md` generado con `docs/herramientas/codigos.py` y pruebas que comprueban que todo mensaje
  tenga un código del catálogo y que el documento esté al día.

## [2.3.0] - 2026-10-26 · Plazo según MS Project, tendencia de los índices y lista EDT

### Nuevas funciones

- **Plazo según el cronograma de MS Project** (#22): `generar` y `actualizar` leen la holgura total y la
  marca de tarea crítica de cada tarea, y el libro muestra si un atraso afecta la fecha de fin:
  - `Indicadores EDT`, columnas Q a U: fin de línea base, fin previsto, desvío de fin (días laborables;
    > 0 = atraso), holgura total y crítica de cada elemento EDT, con sus tareas de detalle; en rojo el
    desvío de los paquetes críticos o con holgura negativa
    ([El libro Excel](docs/libro-excel.md#indicadores-edt)).
  - `Resumen` y `Cuadro de Mando`: fin de línea base y fin previsto del proyecto, desvío de fin, paquetes
    críticos sin terminar (y cuántos con el fin atrasado) y término estimado con el SPI(t).
  - `Análisis de desviaciones`: desvío de fin, holgura y crítica de cada paquete, con los valores del mes
    del análisis.
  - `actualizar` muestra en pantalla el fin previsto, el desvío y los paquetes críticos atrasados.
- **Tendencia de los índices** (#23):
  - `Cuadro de Mando`: CPI y SPI de los últimos 3 meses y variación del CPI, el SPI y el EAC respecto del
    mes anterior con datos (con informes bimestrales, el del informe anterior)
    ([Cuadro de Mando](docs/cuadro-de-mando.md#4-tendencia-y-plazo)).
  - `Resumen` y salida de `actualizar`: los mismos indicadores a la fecha de corte.
  - `Análisis de desviaciones`: SPI y CPI del mes anterior de cada paquete y columna `Tendencia`
    (`Mejora`, `Estable` o `Empeora`).
- **Lista EDT** (#27): nueva clave opcional `[exportacion] lista_edt` con un archivo de texto de códigos
  EDT, uno por línea, que decide las filas del libro en lugar de `nivel_maximo` y del campo `Visible`. Las
  tareas que no están en la lista se suman en su elemento listado más cercano; los códigos de la lista
  que no existen en MS Project se advierten. El libro guarda la lista y `actualizar` la usa
  ([Cálculo del PV y el EV](docs/calculo-pv-ev.md#34-lista-edt-filas-definidas-en-un-archivo)).
- **Versión y barra de progreso:** todos los comandos muestran la versión del programa en la primera
  línea y, en una consola, una barra de progreso con la tarea en curso, el % de esa tarea y del total y
  el tiempo restante estimado ([Comandos](docs/comandos.md)).

### Correcciones

- Los códigos EDT con un 0 intermedio, como `1.2.0.1` debajo de `1.2.0.0` (un elemento insertado antes de
  `1.2.1.0` sin renumerar a sus hermanos), son válidos y de nivel 3. Antes se informaban como error
  (#28) ([Validaciones](docs/validaciones-y-problemas.md#11-datos-del-libro-generar-y-validar)).

### Otros

- El `Cuadro de Mando` tiene dos filas de tarjetas nuevas sobre los gráficos. Los libros de versiones
  anteriores conservan su disposición; `actualizar` les agrega las columnas de plazo, las secciones del
  `Resumen` y las columnas nuevas del `Análisis de desviaciones`, y las tarjetas nuevas se agregan al
  regenerarlos con `--migrar-desde`.
- El proyecto de ejemplo simula el cronograma vigente (reprogramación del trabajo pendiente), la holgura
  total y las tareas críticas.
- Instalador de Windows de esta versión.
- Portada de la presentación con la versión 2.3.0.

## [2.2.1] - 2026-10-25 · Formato de números, códigos de salida y revisión de la documentación

### Nuevas funciones

- Ninguna.

### Correcciones

- La consola muestra todos los números con punto de miles y coma decimal, como Excel en español:
  indicadores de `actualizar` (`44.210.000`, `SPI = 0,84`, `27,4 %`), paquetes fuera de umbral, tabla de
  `--simular` e importes de `--migrar-desde`. Antes se mezclaban `0,95 a 1,05` con `1.00` y
  `44,210,000`. Los mensajes de la hoja `Validación` con importes o porcentajes usan el mismo formato.
- Códigos de salida: los errores inesperados del programa terminan con código 3 (antes, 1), de modo que
  el código 1 queda solo para `validar` con errores. Los errores de lectura o escritura de archivos (disco
  lleno, unidad de red no disponible) terminan con código 2 y un mensaje claro, en lugar de informarse
  como error inesperado. Tabla de códigos de todos los comandos en
  [docs/comandos.md](docs/comandos.md#códigos-de-salida).

### Otros

- Revisión de coherencia de la documentación y de la presentación con el programa:
  - `validar` y `generar-fichas` aparecen donde corresponde: README, índice de la documentación, proceso
    de reporte (cierre mensual, diagramas y roles), tutorial, varios archivos, multi-empresa y
    presentación (diapositivas 3, 11, 17, 25, 26, 27, 36 y 37, y sus notas).
  - El cierre mensual sigue el mismo orden en todos los documentos: fecha de estado, registro del
    avance, reprogramación del trabajo pendiente y `validar`.
  - Se corrigieron detalles que no coincidían con el programa: el peso de las tareas con
    `distribucion = "calendario"`, los feriados con XML, el mes 1 por defecto, las filas `-D`, el campo
    `Visible` modificado después de generar, `--nivel-max` en `actualizar`, los códigos de salida de
    `validar`, el Cronograma Ganado cuando la curva PV no llega al BAC, la consulta de un mes anterior en
    el análisis de desviaciones y el ejemplo de `% Completado` de la presentación (con hitos que pesan 100).
  - Presentación: pestañas `% PV` y `% EV` en el libro (diapositiva 30) y extracto actual de la hoja `AC`.
  - Plantillas de tickets sin la opción `--solo ev`, que se eliminó en la versión 2.0.0, y con
    `generar-fichas` (parte del ticket #18); paso manual de los enlaces de las notas de versión en
    [Desarrollo](docs/desarrollo.md#31-versiones-y-publicaciones).
- Instalador de Windows de esta versión.
- Portada de la presentación con la versión 2.2.1.

## [2.2.0] - 2026-10-24 · Comando generar-fichas: fichas de tareas para el cliente

### Nuevas funciones

- **Comando `generar-fichas`:** crea un documento Markdown (`.md`) para el cliente con una ficha por
  cada elemento EDT del libro: código EDT, nombre, inicio y fin planificados, duración en días
  laborables (de la línea base, como el % planificado), y predecesoras y sucesoras (`código - nombre`).
  Los vínculos entre tareas de detalle se trasladan a los elementos con ficha que las contienen. Los
  vínculos con tareas sin ficha (otra planificación, inactivas, con `Visible = No`) no se muestran
  ([docs/comandos.md](docs/comandos.md#generar-fichas-fichas-de-tareas-para-el-cliente)).
- Nueva clave `[campos] descripcion`: agrega a cada ficha las notas de la tarea (`"Notas"`) o un campo de
  texto. Por defecto no se lee.
- Los lectores `.mpp` y XML leen la duración de la línea base y las horas por día del proyecto.

### Correcciones

- Ninguna.

### Otros

- Instalador de Windows de esta versión, que incluye también el comando `validar` de la versión 2.1.0.
  El instalador prueba también `generar-fichas` con el proyecto de ejemplo.
- `.gitignore` ignora las fichas generadas (`*_fichas.md`).
- Portada de la presentación con la versión 2.2.0.

## [2.1.0] - 2026-10-24 · Comando validar: revisión del archivo de MS Project

### Nuevas funciones

- **Comando `validar`:** revisa el archivo de MS Project con la configuración y muestra el resultado en
  pantalla, agrupado por tema, sin crear ni modificar ningún archivo. Termina con código 1 si encuentra
  errores. Además de las observaciones de la hoja `Validación`, revisa el avance a la fecha de estado
  (fechas reales, trabajo pendiente antes de la fecha de estado, índice BEI) y la lógica del cronograma
  (vínculos, adelantos, posiciones y delimitaciones, con los criterios DCMA de 14 puntos)
  ([docs/comandos.md](docs/comandos.md#validar-revisa-el-archivo-de-ms-project)).
- **Controles nuevos en la hoja `Validación` (y en `validar`):** códigos EDT que no dependen del de su
  tarea superior (por ejemplo `1.14.2` bajo `1.2.0`), con formato no válido, que saltan un nivel, fuera
  de orden o con saltos en la numeración; códigos duplicados en todas las tareas (ahora como `Error`);
  fechas invertidas, hitos con duración, tareas con duración 0 que no son hitos, duraciones estimadas y
  tareas manuales; `Peso hito` en tareas que no son hitos; filas con `Hitos ponderados` sin hitos o con
  hitos sin peso; métodos no reconocidos o cargados en subtareas; `% Reparto` en tareas que no son
  `Ambas`; campos configurados que existen pero están vacíos
  ([docs/validaciones-y-problemas.md](docs/validaciones-y-problemas.md#1-validaciones)).
- **Método frente a la duración del paquete:** advertencia si un paquete con `0/100` dura más de 1 mes o
  uno con `50/50`, `25/75`, `20/80` más de 2 meses (línea base). Los umbrales se configuran en la nueva
  sección `[validacion]` (`meses_0_100`, `meses_formula_fija`).
- La hoja `Validación` tiene la columna `Tema` para filtrar las observaciones, y las tareas sin código
  EDT se identifican con su Id de MS Project.
- Los lectores `.mpp` y XML leen la duración, la duración estimada, la programación manual, las
  delimitaciones y los vínculos con las predecesoras.

### Correcciones

- Ninguna.

### Otros

- El instalador prueba también `validar` con el proyecto de ejemplo.
- Esta versión no publica instalador de Windows: se publicará con la próxima versión. Hasta entonces, el
  comando `validar` se usa con la instalación con Python; el instalador de la versión 2.0.0 sigue
  disponible, sin este comando.
- Portada de la presentación con la versión 2.1.0.

## [2.0.0] - 2026-10-24 · PV y EV con el BAC del libro: cronograma sin importes

### Nuevas funciones

- **PV y EV con el BAC ingresado en el libro:** el cronograma de MS Project no necesita importes.
  `generar` calcula el % planificado de cada fila y mes con la duración de la línea base de sus tareas y
  el `Peso hito` de sus hitos (el % del BAC que se imputa en el mes del hito), y `actualizar` calcula el
  % ganado según el método de medición. El BAC de cada fila se ingresa en el libro, y el PV y el EV son
  el BAC × el % (fórmulas de Excel). El AC se carga a mano ([docs/calculo-pv-ev.md](docs/calculo-pv-ev.md)).
- Hojas nuevas `% PV` y `% EV` (en multi-empresa, una por empresa) con el % de cada mes; las filas de
  nivel 1 y 2 y la fila TOTAL muestran el % ponderado por el BAC.
- Multi-empresa: el BAC de cada empresa se ingresa en su hoja `PV`, y cada empresa tiene su %
  planificado y su % ganado con las tareas en las que participa (`Empresa` y `% Reparto`). Los métodos
  `0/100`, `X/Y` y `Hitos ponderados` también se miden por empresa.
- Cambio de BAC: al escribir un BAC nuevo, el PV y el EV de todos los meses de la fila se recalculan.
- `--migrar-desde` traslada el BAC ingresado, el % ganado, el AC, los métodos y el análisis de
  desviaciones.
- Validaciones nuevas: tareas sin fechas de línea base, pesos de hitos que suman más de 100, filas sin
  pesos o sin fechas, empresas que no participan en una fila y, en `actualizar`, filas con avance y sin
  BAC, y tareas con código EDT sin fila en el libro (agregadas en MS Project después de generarlo).
- `ejemplos/generar_ejemplo.py --libro`: escribe en el libro el BAC y el AC ficticios del proyecto de
  ejemplo, que en un proyecto real se ingresan a mano.
- **Depuración de costes** (opcional; `[exportacion] depuracion_costes = true`): una empresa pega en la
  hoja nueva `PV-depuración` la previsión mensual de su plan de costes (código EDT, descripción, BAC y
  meses 1 a 120). Con fórmulas de Excel, cada línea se asigna a una fila del libro: las subtareas sin
  código numérico (`XRF01`) se agrupan bajo su código EDT padre, los códigos más detallados van a su fila
  y las líneas con subtareas son subtotales. En la hoja PV de la empresa (`depuracion_empresa`), debajo de
  cada fila va su fila «(coste)», y la hoja nueva `Depuración costes` compara las dos curvas acumuladas
  del elemento EDT elegido. Indicador de alineamiento opcional (`depuracion_alineamiento`,
  `depuracion_umbral`). `--migrar-desde` conserva la previsión pegada; `actualizar` y `analizar`
  funcionan igual ([docs/depuracion-costes.md](docs/depuracion-costes.md)).
- `ejemplos/generar_ejemplo.py --libro` en un libro con depuración de costes: escribe también una
  previsión ficticia en `PV-depuración`, con meses desplazados, hitos de pago y subtareas.
- `docs/herramientas/capturas_excel.py --solo GRUPO`: regenera solo un grupo de capturas.
- **Campo `Visible`** de MS Project (`[campos] visible`, por defecto el campo con alias `Visible`): decide
  tarea por tarea qué información muestra el libro. `Sí` da fila propia, también por debajo de
  `nivel_maximo`; `No` deja la tarea y su rama fuera del PV, el EV y el AC, sin crear filas `-D` (si
  todas las subtareas de una tarea tienen `No`, la tarea las representa y, si tiene fila, pasa a ser el
  paquete de trabajo); vacío sigue `nivel_maximo`. `actualizar` usa el campo con que se generó el libro y
  avisa si cambió ([docs/calculo-pv-ev.md](docs/calculo-pv-ev.md#33-campo-visible-más-o-menos-detalle-en-cada-rama)).

### Correcciones

- Ninguna.

### Otros

- **Se eliminó el cálculo con los costos de MS Project** (el modo costo de las versiones 1.x): el PV y
  el EV se calculan siempre con el BAC ingresado en el libro. Para calcular el PV, el BAC y el AC con los
  costos de MS Project se debe usar la **versión 1.5.1**.
- Los libros generados con la versión 1.x no se pueden actualizar ni migrar con la 2.0.0: se informa que
  se deben usar con la versión 1.x, o se genera un libro nuevo.
- Se eliminaron la clave `[exportacion] modo` (si está en la configuración, se ignora), la clave
  `[multiempresa] ac_desde_project`, la opción `actualizar --solo ev|ac` y el método `CPTR de MS Project`
  (si una fila lo tiene, se informa como método no reconocido y se usa `% Completado`). El programa ya no
  lee costos de MS Project (costo de línea base y su distribución mensual, costo real, CPTR y CRTR).
- Documentación: [docs/calculo-pv-ev.md](docs/calculo-pv-ev.md) reúne el cálculo del PV, el EV y el AC y
  reemplaza a `modo-porcentaje.md`, `calculo-pv.md` y `calculo-ev-ac.md`; se eliminó `presupuesto-bac.md`.
  Tutorial, capturas del libro y presentación actualizados.
- Pruebas: `test_calculo_pv_ev.py`; se eliminaron las pruebas de los cálculos con costos y la marca
  `modo`.

## [1.5.1] - 2026-10-19 · Método de medición documentado e instalador para Windows

### Nuevas funciones

- Ninguna.

### Correcciones

- Ninguna.

### Otros

- **Instalador para Windows** (`ExportadorEVM-1.5.1-instalador.exe`), adjunto a la publicación de
  GitHub: instala el programa sin Python y sin permisos de administrador, con los comandos `evm_export` y
  `generar_ejemplo`, la configuración y el proyecto de ejemplo, un acceso a la consola en el menú Inicio
  y la opción de agregar el programa al PATH del usuario. Se genera con `python instalador\construir.py`
  (PyInstaller e Inno Setup), que además prueba los ejecutables con el proyecto de ejemplo.
- Registro de cambios (este archivo), etiquetas `v1.0.0` a `v1.5.0` y publicaciones en GitHub; versión
  en la portada de la presentación.
- Sección *Contribuciones* en el README: cómo informar problemas, proponer mejoras y enviar Pull
  Requests.
- Plantillas de GitHub: formularios de tickets para informar un problema y para proponer una mejora, y
  plantilla de Pull Request con lista de control.
- Documentación de los métodos de medición: dónde se define el método y dónde se calcula el EV
  (MS Project, programa y Excel), y proceso para definir y aprobar el método de cada paquete (roles,
  pasos, criterio de medición, cambios de método y registro de los métodos aprobados).
- Presentación (37 diapositivas): nueva diapositiva *Cómo se define el método de cada paquete* y, en *El
  libro Excel*, el método se define en MS Project y la columna C solo se edita para cambios aprobados.

## [1.5.0] - 2026-10-17 · Análisis de desviaciones y semáforo con umbrales superiores

### Nuevas funciones

- Hoja **Análisis de desviaciones**: lista los paquetes de trabajo con SPI o CPI fuera del rango verde
  del semáforo en el mes de corte, con columnas para la causa, el impacto, la acción correctiva, el
  responsable y la fecha compromiso. El texto se conserva al actualizar y con `--migrar-desde`, y se copia
  al mes siguiente si el paquete sigue fuera de umbral.
- Comando **`analizar`**: vuelve a calcular esa lista sin leer MS Project, por ejemplo después de cargar a
  mano el AC del ERP.
- **Semáforo con umbrales inferiores y superiores**, comunes para el CPI y el SPI: `rojo_inferior`,
  `amarillo_inferior`, `amarillo_superior` y `rojo_superior` (0,90 / 0,95 / 1,05 / 1,10 por defecto).
  Se aplica en `Resumen`, `Indicadores EDT`, `Cuadro de Mando` y `Análisis de desviaciones`.
- El `Cuadro de Mando` muestra la escala del semáforo con los umbrales del libro, y el gráfico de CPI y SPI
  tiene también las bandas superiores.
- `actualizar` muestra en pantalla los paquetes fuera de umbral, también con `--simular`.

### Correcciones

- Ninguna.

### Otros

- Los índices por encima de 1,05 ahora se marcan en amarillo o rojo. Para mantener el comportamiento
  anterior se indican umbrales superiores altos (por ejemplo, 99).
- Los nombres anteriores `rojo` y `amarillo` siguen siendo válidos para los umbrales inferiores.
- Los libros de versiones anteriores reciben la hoja de análisis en el siguiente `actualizar`. Para
  aplicar los umbrales superiores en todas las hojas, se regeneran con `--migrar-desde`.
- Presentación actualizada (36 diapositivas): semáforo, comando `analizar` y análisis de desviaciones.
- Documento nuevo `docs/analisis-desviaciones.md`; 446 pruebas automáticas.

## [1.4.0] - 2026-10-17 · Varios archivos de MS Project en un mismo libro

### Nuevas funciones

- Secciones **`[[proyectos]]`** en la configuración: varios archivos `.mpp` o `.xml` en un mismo libro,
  cada uno con sus códigos EDT (por ejemplo, `1.0.0` a `8.0.0` en uno y `9.0.0` en otro), su
  planificación y su línea base. Las filas se ordenan por código EDT y los totales son los del programa.
- `actualizar` lee todos los archivos con una sola fecha de corte, con los controles por archivo.
  `--origen NOMBRE=RUTA` reemplaza la ruta de un archivo.
- `--migrar-desde` entre libros de varios archivos, al agregar o quitar archivos y desde un libro de un
  solo archivo.
- `[exportacion] titulo`: título del informe en el `Cuadro de Mando` y en `Resumen`.

### Correcciones

- Ninguna.

### Otros

- Con varios archivos, un código EDT repetido entre archivos detiene `generar`, y la fila TOTAL se llama
  "Total programa".
- Segundo proyecto de ejemplo (`ejemplos/proyecto2_ejemplo.xml`, código EDT `9.0.0`).
- Documento nuevo `docs/varios-archivos.md`, con capturas.

## [1.3.2] - 2026-10-17 · Conexión y cierre de MS Project más robustos

### Nuevas funciones

- Pruebas opcionales con MS Project real (marca `msproject`): comparan la lectura del `.mpp` por
  automatización COM con la del XML del mismo proyecto.

### Correcciones

- Conexión con MS Project: si se está cerrando o muestra un diálogo, el programa espera y reintenta
  durante 30 segundos como máximo, y muestra un mensaje claro si no responde.
- Cierre de MS Project: las referencias a tareas y proyectos se liberan antes de cerrarlo (error de RPC
  `0x800706BA`).

### Otros

- Documentación reorganizada: el README queda con la presentación y la instalación, y el detalle pasa a
  documentos por tema en `docs/`, con capturas reales del libro y diagramas.
- Presentación de 30 minutos para el equipo de proyecto (PowerPoint y PDF).
- Lectura de `.mpp` por COM validada con un proyecto real.

## [1.3.1] - 2026-10-17 · Pruebas de robustez y correcciones

### Nuevas funciones

- `actualizar` verifica el libro antes de leer MS Project (hojas, filas y hojas ocultas), se detiene sin
  modificarlo si encuentra un problema y avisa si está abierto en Excel.
- Guardado en un archivo temporal: un error al guardar no daña el libro anterior.
- Avisos nuevos: fórmulas reemplazadas por valores, paquetes nuevos bajo una fila `-D`, tareas
  desactivadas después de generar el libro y meses posteriores cargados con `--forzar`.
- Archivo de configuración en UTF-8 con BOM (Bloc de notas).

### Correcciones

- EV de las filas `-D Otros costos`: el avance se mide solo con las tareas de la fila.
- Hitos ponderados: un peso negativo se toma como 0 y el EV ya no supera el BAC.
- Distribución mensual de MS Project con signo opuesto al costo de línea base: se reparte en forma
  lineal y se avisa.
- El mes 1 del horizonte ya no se adelanta por tareas omitidas, y el Cronograma Ganado admite un EV
  negativo.
- Los caracteres de control en los nombres y los horizontes de más de 336 meses impedían generar el libro.
- `--migrar-desde` conserva los meses cargados con 0.
- XML: un campo de recurso con el mismo alias ya no oculta el campo de tarea, solo la tarea de UID 0 es
  el resumen del proyecto y las excepciones de calendario no interpretadas se informan.
- COM: el MS Project que inicia el programa se cierra aunque quede abierto el fondo de recursos.
- Mensajes claros para un archivo de configuración que no está en UTF-8 o con valores no válidos, y para
  `--meses 0` y `--nivel-max 0`.

### Otros

- 383 pruebas automáticas, con proyectos aleatorios de punta a punta y recálculo de las fórmulas con
  LibreOffice (marca `libreoffice`).

## [1.3.0] - 2026-10-16 · Cuadro de Mando con mes de reporte y gráficos nativos

### Nuevas funciones

- Hoja **Cuadro de Mando**, la primera del libro: valores acumulados, desviaciones, índices con semáforo
  y pronósticos (ETC, EAC, VAC y TCPI) del mes de reporte elegido, y tres gráficos nativos de Excel:
  evolución EVM, evolución de CPI y SPI sobre zonas de alarma, y CPI vs SPI por elemento EDT con su tabla.
- Tres métodos de EAC seleccionables: `BAC / CPI`, `AC + (BAC - EV)` y `AC + (BAC - EV) / (CPI × SPI)`.
- Se imprime en una página, lista para el informe mensual.

### Correcciones

- Ninguna.

### Otros

- La hoja `Curva S` pasa a llamarse `Evolución mensual` y conserva la tabla; los datos de los gráficos
  van en la hoja oculta `_graficos`.
- Los libros anteriores siguen funcionando con `actualizar`; para obtener el tablero se regeneran con
  `--migrar-desde`.

## [1.2.0] - 2026-10-16 · Multi-empresa: PV, EV y AC por empresa

### Nuevas funciones

- Modo **multi-empresa** opcional (`[multiempresa]`): el PV, el EV y el AC se reparten entre dos
  empresas según la columna `Empresa` de cada tarea (`EmpresaA`, `EmpresaB` o `Ambas` con `% Reparto`),
  con herencia de la tarea superior.
- Hojas `PV-`, `EV-` y `AC-` de cada empresa; las hojas `PV`, `EV` y `AC` son la suma de las dos.
- `actualizar` calcula el EV de cada empresa según el método de medición. El AC se carga a mano o, con
  `ac_desde_project = true`, desde MS Project.
- `Validación`: BAC de cada empresa, tareas sin empresa, valores no reconocidos y `% Reparto` fuera de
  rango.

### Correcciones

- Ninguna.

### Otros

- `config.ejemplo.toml` con los valores del programa (campo `Item EDT`, horizonte desde enero de 2026).
- El XML de ejemplo incluye los campos `Item EDT`, `Empresa` y `% Reparto`.

## [1.1.0] - 2026-10-16 · Planificaciones, niveles por código EDT y curva S en Excel

### Nuevas funciones

- Elección de la **planificación** a exportar (`INTERNO`, `CLIENTE`...) con `planificacion` en
  `[exportacion]` o con `--planificacion`. `actualizar` usa la planificación con que se generó el libro.
- El nivel de cada fila se toma del **código EDT** (`2.0.0` nivel 1, `2.1.0` nivel 2, `2.1.1` nivel 3),
  no de su posición en el esquema de MS Project.
- Las tareas sin código EDT no tienen fila propia: su costo se suma en su elemento EDT y las que no
  dependen de ningún elemento se informan en `Validación`.
- Curva S calculada solo en Excel, con fórmulas y un gráfico nativo sobre todo el horizonte.
- Moneda por defecto: euro (`EUR`); `moneda = ""` usa la moneda del proyecto.

### Correcciones

- Un código EDT vacío en un campo personalizado se reemplazaba por el EDT automático de MS Project.
- Un campo personalizado llamado "EDT" se confundía con el campo EDT de MS Project.

### Otros

- Se eliminan el comando `curva` (imagen PNG) y la dependencia `matplotlib`.
- Licencia MIT.
- Redacción impersonal en la documentación y en los mensajes.
- El proyecto de ejemplo tiene dos planificaciones (`INTERNO` y `CLIENTE`) e hitos sin código EDT.

## [1.0.0] - 2026-10-16 · Primera versión: exportador EVM desde MS Project a Excel

### Nuevas funciones

- Lectura del cronograma desde `.mpp` (MS Project por automatización COM) o desde el XML de MS Project,
  con el costo de línea base mensualizado.
- `generar`: libro Excel con el PV mensual de cada elemento de la EDT hasta el nivel 3 (120 meses), hojas
  EV y AC para la carga mensual, curva S, `Resumen` con los indicadores (SV, CV, SPI, CPI, EAC, ETC, VAC,
  TCPI y Cronograma Ganado), `Indicadores EDT` y `Validación`.
- `actualizar`: EV según el método de medición (0/100, fórmulas fijas, hitos ponderados, % completado,
  % físico completado, unidades, LOE y CPTR) y AC a la fecha de corte, cargados en su mes con una copia
  de respaldo.
- `curva`: imagen de la curva S e indicadores en pantalla.
- `--migrar-desde`: conserva EV, AC y métodos al regenerar el libro.

### Correcciones

- Ninguna (primera versión).

### Otros

- 73 pruebas automáticas en Python 3.13 y 3.14.

[Sin publicar]: https://github.com/MarceloSalazar/evm-automation/compare/v2.3.0...HEAD
[2.3.0]: https://github.com/MarceloSalazar/evm-automation/compare/v2.2.1...v2.3.0
[2.2.1]: https://github.com/MarceloSalazar/evm-automation/compare/v2.2.0...v2.2.1
[2.2.0]: https://github.com/MarceloSalazar/evm-automation/compare/v2.1.0...v2.2.0
[2.1.0]: https://github.com/MarceloSalazar/evm-automation/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.5.1...v2.0.0
[1.5.1]: https://github.com/MarceloSalazar/evm-automation/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.2...v1.4.0
[1.3.2]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.1...v1.3.2
[1.3.1]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.0...v1.3.1
[1.3.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/MarceloSalazar/evm-automation/releases/tag/v1.0.0
