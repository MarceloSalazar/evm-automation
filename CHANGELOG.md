# Registro de cambios

Cambios de cada versión del Exportador EVM, de la más reciente a la más antigua. Cada versión tiene su
etiqueta `vX.Y.Z` en el repositorio y su publicación (*release*) en
[GitHub](https://github.com/MarceloSalazar/evm-automation/releases) con estas mismas notas.

Las versiones siguen el formato MAYOR.MENOR.CORRECCIÓN ([versionado semántico](https://semver.org/lang/es/)):
MENOR para funciones nuevas que no impiden usar los libros ya generados, CORRECCIÓN para correcciones, y
MAYOR para cambios que exigen regenerar los libros o cambiar la configuración. La versión instalada se
consulta con `python evm_export.py --version`.

Los cambios incorporados a `main` que todavía no forman parte de una versión se anotan en
*Sin publicar*; al publicar la versión siguiente, esa sección pasa a ser su entrada.

---

## [Sin publicar]

### Nuevas funciones

- Ninguna.

### Correcciones

- Ninguna.

### Otros

- Ninguna.

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

[Sin publicar]: https://github.com/MarceloSalazar/evm-automation/compare/v1.5.1...HEAD
[1.5.1]: https://github.com/MarceloSalazar/evm-automation/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.2...v1.4.0
[1.3.2]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.1...v1.3.2
[1.3.1]: https://github.com/MarceloSalazar/evm-automation/compare/v1.3.0...v1.3.1
[1.3.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/MarceloSalazar/evm-automation/releases/tag/v1.0.0
