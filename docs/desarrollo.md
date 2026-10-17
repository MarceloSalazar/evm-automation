# Desarrollo, pruebas y repositorio

[Inicio](../README.md) › [Documentación](README.md) › Desarrollo

Información para quien modifica el programa o la documentación.

---

## Contenido

1. [Pruebas automáticas](#1-pruebas-automáticas)
2. [Estructura del código](#2-estructura-del-código)
3. [Flujo de trabajo del repositorio](#3-flujo-de-trabajo-del-repositorio)
4. [Documentación e imágenes](#4-documentación-e-imágenes)

---

## 1. Pruebas automáticas

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest                          # todas las pruebas
python -m pytest -m "not libreoffice"     # sin las que recalculan con LibreOffice (más rápido)
```

| Archivo de pruebas | Contenido |
|---|---|
| `test_robustez_calculo.py` | Estructura EDT y PV, métodos de medición, multi-empresa, reparto en el tiempo, indicadores y configuración. Incluye reglas que se verifican sobre árboles EDT aleatorios. |
| `test_robustez_lectores.py` | Lectura del XML (campos, líneas base, calendarios y excepciones) y del `.mpp` con objetos que simulan MS Project. |
| `test_robustez_libro.py` | `actualizar`, `--migrar-desde`, libros modificados a mano, errores de la línea de comandos y del archivo de configuración. |
| `test_proyectos_aleatorios.py` | Proyectos aleatorios de punta a punta (`generar`, varias fechas de corte, migración y recálculo). |
| `test_analisis_desviaciones.py` | Hoja `Análisis de desviaciones`: lista calculada aparte, texto conservado y copiado al mes siguiente, `analizar` con AC cargado a mano, libros antiguos, `--migrar-desde` y recálculo de sus fórmulas con LibreOffice. |
| `test_varios_archivos.py` | Varios archivos de MS Project (`[[proyectos]]`): el libro combinado coincide, fila por fila, con el libro de cada archivo (también en multi-empresa, con UniqueID repetidos entre archivos); errores de configuración y de la línea de comandos, fechas de estado y `--migrar-desde`. |
| Resto | Casos del programa de ejemplo, Cuadro de Mando, multi-empresa y planificación. |

- **LibreOffice:** si está instalado, las pruebas marcadas `libreoffice` recalculan las fórmulas del
  libro y las comparan con el cálculo en Python. Se usa un perfil propio de LibreOffice, por lo que
  funcionan aunque LibreOffice esté abierto. Si no está instalado, se omiten.
- **Proyectos aleatorios:** la cantidad de proyectos (y de árboles EDT aleatorios) se aumenta con la
  variable de entorno `EVM_SEMILLAS`, por ejemplo `$env:EVM_SEMILLAS = 50` antes de `python -m pytest`.
  Cada caso indica su semilla, de modo que un fallo se puede repetir.
- **Problemas conocidos:** las pruebas marcadas `xfail` documentan problemas detectados que aún no se
  corrigieron; aparecen como `xfailed` y no hacen fallar la ejecución. Cuando se corrige el problema,
  la prueba aparece como fallida (`XPASS(strict)`) para que se quite la marca.
- **Cobertura (opcional):** `python -m pip install coverage`, luego
  `python -m coverage run --branch --source=evm -m pytest` y `python -m coverage report -m`.

### 1.1 Pruebas con MS Project real (opcionales)

El lector `.mpp` se prueba con objetos que simulan MS Project. En un PC con MS Project, las pruebas
opcionales de `test_msproject_real.py` (marca `msproject`) leen un proyecto real por automatización COM
y comparan el resultado con el XML que MS Project guarda del mismo proyecto: estructura EDT, métodos,
BAC, PV mensual, EV y AC a la fecha de corte, y que MS Project quede como estaba (abierto o cerrado).

1. En MS Project se abre el proyecto, se define *Proyecto > Fecha de estado* y se guarda el `.mpp`.
2. Se guarda una copia en XML: *Archivo > Guardar como*, tipo XML.
3. Se ejecutan las pruebas:

```powershell
$env:EVM_MPP = "C:\Proyectos\Programa.mpp"
$env:EVM_MPP_XML = "C:\Proyectos\Programa.xml"
python -m pytest -m msproject -v
```

Opcionales: `EVM_CONFIG` (archivo de configuración; por defecto `config.toml`), `EVM_LINEA_BASE`
(línea base N) y `EVM_TOLERANCIA_PV` (diferencia admitida en el PV de cada mes, como fracción del BAC
de la fila; por defecto 0.005). Los mensajes de las pruebas solo muestran códigos EDT e importes.
Además, se recomienda comparar el BAC total y el PV de algunos meses con la vista *Uso de tareas*
(campo *Costo previsto* en la escala temporal mensual).

El proyecto de ejemplo (`ejemplos/programa_ejemplo.xml`) y su generador se describen en el
[tutorial](ejemplo-paso-a-paso.md). El segundo proyecto de ejemplo (`ejemplos/proyecto2_ejemplo.xml`,
para varios archivos) se genera con `python ejemplos\generar_ejemplo.py --proyecto 2`.

---

## 2. Estructura del código

| Archivo | Función |
|---|---|
| `evm_export.py` | Punto de entrada. |
| `evm/cli.py` | Comandos y opciones. |
| `evm/config.py` | Configuración TOML. |
| `evm/model.py` | Datos comunes: tareas y proyecto leídos de MS Project. |
| `evm/reader_com.py` | Lectura de `.mpp` con MS Project (COM). |
| `evm/reader_xml.py`, `evm/workcalendar.py` | Lectura del XML de MS Project y de su calendario. |
| `evm/fields.py` | Identificación de campos personalizados (Texto1..30, Número1..20 y alias). |
| `evm/wbs.py` | Estructura EDT hasta el nivel máximo y PV por fila; combinación de varios archivos y horizonte. |
| `evm/companies.py` | Multi-empresa: parte de cada empresa por tarea y reparto del EV y el AC. |
| `evm/timeline.py` | Horizonte mensual y repartos en el tiempo. |
| `evm/workbook.py` | Escritura del libro Excel y de sus indicadores. |
| `evm/dashboard.py` | Hoja `Cuadro de Mando` y sus tres gráficos. |
| `evm/methods.py` | Métodos de medición del EV. |
| `evm/analysis.py` | Hoja `Análisis de desviaciones`: paquetes fuera de umbral y texto conservado. |
| `evm/update.py` | Comando `actualizar` y lectura del libro. |
| `evm/migrate.py` | Traslado de EV, AC y métodos al regenerar el libro (`--migrar-desde`). |
| `evm/indicators.py` | Indicadores que se muestran en pantalla (mismos cálculos que la hoja `Resumen`). |

```mermaid
flowchart LR
    CLI["cli.py"] --> RC["reader_com.py"] & RX["reader_xml.py"]
    RC & RX --> M["model.py<br/>ProjectData / Task"]
    M --> W["wbs.py<br/>filas EDT y PV"]
    W --> CO["companies.py"]
    W --> WB["workbook.py<br/>+ dashboard.py"]
    CLI --> U["update.py<br/>actualizar"]
    U --> ME["methods.py"] & CO
    U --> WB
    CLI --> MI["migrate.py"]
```

---

## 3. Flujo de trabajo del repositorio

- La rama `main` contiene solo **versiones estables**. No se hacen commits directos en `main`.
- Todo cambio se hace en una rama propia y se incorpora mediante un **Pull Request**:

```powershell
git switch main
git pull
git switch -c mejora/descripcion-breve
# ... cambios ...
python -m pytest
git add -A
git commit -m "Descripción del cambio"
git push -u origin mejora/descripcion-breve
gh pr create --base main --fill
```

- Antes de abrir el Pull Request, todas las pruebas deben pasar (`python -m pytest`).
- No se deben subir al repositorio libros generados, configuraciones personales (`config.toml`) ni
  archivos de proyectos reales: el repositorio es público.
- Todos los textos (documentación, mensajes y notas del libro) se escriben en forma impersonal, por
  ejemplo "se debe habilitar" en lugar de "habilítelo".

---

## 4. Documentación e imágenes

- La documentación está en `README.md` (presentación, instalación y uso rápido) y en la carpeta
  [`docs`](README.md). Los diagramas se escriben en [Mermaid](https://mermaid.js.org/), que GitHub
  muestra como imágenes.
- `docs/img/excel/` contiene **capturas del libro** generado con el proyecto de ejemplo. Si cambia el
  formato del libro, se regeneran con:

  ```powershell
  python docs\herramientas\capturas_excel.py
  ```

  El script genera el libro de ejemplo en una carpeta temporal, simula 24 meses de avance, abre el libro
  con Excel (sin mostrarlo y sin guardar cambios), exporta cada zona a PDF y la convierte a PNG con
  `pdftoppm` (Poppler). Requiere Windows, Excel y Poppler (`winget install oschwartz10612.Poppler`, o la
  ruta de `pdftoppm.exe` en la variable de entorno `PDFTOPPM`).
- `docs/img/ms-project/` contiene **esquemas** de pantallas de MS Project en formato SVG (texto
  editable). No son capturas: muestran la ubicación de los campos y opciones. Se pueden reemplazar por
  capturas reales de MS Project en español (formato PNG en la misma carpeta, con el mismo nombre de
  archivo, actualizando el enlace en el documento).
- Las imágenes PNG de `docs/img/` están permitidas en `.gitignore` (el resto de los PNG del repositorio
  se ignora).
