"""Regenera las capturas del libro Excel que usa la documentación (docs/img/excel).

Pasos:
1. Genera el libro del proyecto de ejemplo (una empresa, multi-empresa y combinado con
   el segundo proyecto de ejemplo) en una carpeta temporal y simula 24 cierres
   mensuales (enero de 2026 a diciembre de 2027).
2. Abre cada libro con Excel por automatización COM, sin mostrarlo y en solo lectura,
   exporta cada zona a PDF y la convierte a PNG con pdftoppm (Poppler).
3. Agrega los números de las zonas a la captura del Cuadro de Mando.

Requisitos: Windows, Excel, pywin32, Pillow y Poppler (pdftoppm en el PATH, en la
carpeta de WinGet o en la variable de entorno PDFTOPPM).

Para que las capturas muestren el formato español (meses y separadores), la sesión de
Excel usa "," como separador decimal y "." como separador de miles; los valores
originales se restauran al terminar. El libro no se guarda.

Uso (desde la carpeta raíz del repositorio):
    python docs\\herramientas\\capturas_excel.py
"""
from __future__ import annotations

import calendar
import datetime as dt
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "docs" / "img" / "excel"
MONTHS = [(2026 + (m - 1) // 12, (m - 1) % 12 + 1) for m in range(1, 25)]

# (archivo, hoja, rango, opciones)
SINGLE = [
    ("cuadro-de-mando", "Cuadro de Mando", "$B$1:$U$73", {}),
    ("cdm-indicadores", "Cuadro de Mando", "$B$1:$U$23", {}),
    ("cdm-evolucion", "Cuadro de Mando", "$B$24:$U$45", {}),
    ("cdm-cpi-spi-edt", "Cuadro de Mando", "$B$46:$U$73", {}),
    ("resumen", "Resumen", "$A$1:$E$40", {}),
    ("pv", "PV", "$A$1:$P$41", {"landscape": True, "headings": True}),
    ("ev", "EV", "$A$1:$AD$41", {"landscape": True, "headings": True, "hide": "G:R"}),
    ("ac", "AC", "$A$1:$AD$41", {"landscape": True, "headings": True, "hide": "G:R"}),
    ("evolucion-mensual", "Evolución mensual", "$A$1:$O$28", {"landscape": True}),
    ("indicadores-edt", "Indicadores EDT", "$A$1:$P$41", {"landscape": True}),
    ("validacion", "Validación", "$A$1:$D$7", {"landscape": True}),
    ("listas", "Listas", "$A$1:$B$12", {"width": ("B", 150)}),
    ("cdm-mes-anterior", "Cuadro de Mando", "$B$1:$U$44", {"date": ("E5", dt.datetime(2027, 6, 1))}),
    ("cdm-eac-alternativo", "Cuadro de Mando", "$B$1:$U$23",
     {"date": ("E5", dt.datetime(2027, 12, 1)), "value": ("R5", "AC + (BAC - EV) / (CPI × SPI)")}),
    # Análisis de desviaciones de diciembre de 2027, con el texto de dos paquetes como ejemplo
    # (filas: 1.1.1 por encima del rango verde, 2.2.1 y 3.2.1 por debajo)
    ("analisis-desviaciones", "Análisis de desviaciones", "$A$1:$R$21",
     {"landscape": True, "values": [
         ("N5", "Dos puestos de la PMO sin cubrir desde octubre."),
         ("O5", "Ahorro de costo, con riesgo de menor control del programa."),
         ("P5", "Cubrir los puestos vacantes de la PMO."),
         ("Q5", "Director del Programa"), ("R5", dt.datetime(2028, 1, 15)),
         ("N15", "Planos de los proveedores de equipos recibidos con 6 semanas de atraso."),
         ("O15", "Atraso de 1 mes en la emisión para construcción; sin impacto en costo."),
         ("P15", "Reforzar el equipo de detalle con 2 proyectistas hasta marzo."),
         ("Q15", "Jefe de Ingeniería"), ("R15", dt.datetime(2028, 1, 31)),
         ("N21", "Retraso en la fabricación (FAT reprogramado)."),
         ("O21", "Riesgo de atraso del montaje mecánico; penalidad al proveedor."),
         ("P21", "Inspección semanal en fábrica y plan de recuperación del proveedor."),
         ("Q21", "Jefe de Compras"), ("R21", dt.datetime(2028, 2, 15))]}),
]
MULTI = [
    ("multi-pv-empresa-a", "PV-EmpresaA", "$A$1:$P$41", {"landscape": True, "headings": True}),
    ("multi-ev-empresa-b", "EV-EmpresaB", "$A$1:$AD$41", {"landscape": True, "headings": True, "hide": "G:R"}),
    ("multi-validacion", "Validación", "$A$1:$D$8", {"landscape": True}),
]
SEVERAL = [
    ("varios-cdm", "Cuadro de Mando", "$B$1:$U$23", {}),
    ("varios-pv", "PV", "$A$1:$P$49", {"landscape": True, "headings": True, "hide_rows": "5:33"}),
    ("varios-validacion", "Validación", "$A$1:$D$8", {"landscape": True}),
]
# Números de las zonas del Cuadro de Mando: (n, x, y) en píxeles de la captura a 130 ppp
MARKS = [(1, 58, 122), (2, 700, 122), (3, 262, 180), (4, 257, 291), (5, 104, 410),
         (6, 120, 545), (7, 612, 545), (8, 150, 916), (9, 960, 884)]
DPI = 130


def run(*args: str) -> None:
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    subprocess.run([sys.executable, *args], cwd=ROOT, check=True, env=env, stdout=subprocess.DEVNULL)


def build_books(work: Path) -> tuple[Path, Path, Path]:
    """Libros del ejemplo (una empresa, multi-empresa y dos archivos) con 24 meses de avance."""
    text = (ROOT / "config.ejemplo.toml").read_text(encoding="utf-8")
    text = re.sub(r'^metodo = ""', 'metodo = "Método EVM"', text, flags=re.M)
    text = re.sub(r'^peso_hito = ""', 'peso_hito = "Peso hito"', text, flags=re.M)
    single_cfg, multi_cfg = work / "config.toml", work / "config_multi.toml"
    several_cfg = work / "config_varios.toml"
    single_cfg.write_text(text, encoding="utf-8")
    examples = (ROOT / "ejemplos").as_posix()
    several_cfg.write_text(text + f'\n[[proyectos]]\nnombre = "P1"\norigen = "{examples}/programa_ejemplo.xml"\n'
                           f'\n[[proyectos]]\nnombre = "P2"\norigen = "{examples}/proyecto2_ejemplo.xml"\n'
                           'planificacion = ""\n', encoding="utf-8")
    text = re.sub(r"^activar = false", "activar = true", text, flags=re.M)
    text = re.sub(r"^ac_desde_project = false", "ac_desde_project = true", text, flags=re.M)
    multi_cfg.write_text(text, encoding="utf-8")
    origin = str(Path("ejemplos") / "programa_ejemplo.xml")     # relativa: es la que muestra la hoja Resumen
    books = (work / "EVM_Ejemplo.xlsx", work / "EVM_Multi.xlsx", work / "EVM_Varios.xlsx")
    for cfg, book in zip((single_cfg, multi_cfg), books):
        run("evm_export.py", "generar", "--origen", origin, "--config", str(cfg), "--salida", str(book))
    run("evm_export.py", "generar", "--config", str(several_cfg), "--salida", str(books[2]))
    for y, m in MONTHS:
        status = f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"
        progress = work / f"avance_{y:04d}-{m:02d}.xml"
        progress2 = work / f"avance2_{y:04d}-{m:02d}.xml"
        run(str(ROOT / "ejemplos" / "generar_ejemplo.py"), "--fecha-estado", status, "--salida", str(progress))
        run(str(ROOT / "ejemplos" / "generar_ejemplo.py"), "--proyecto", "2", "--fecha-estado", status,
            "--salida", str(progress2))
        for cfg, book in zip((single_cfg, multi_cfg), books):
            run("evm_export.py", "actualizar", "--origen", str(progress), "--config", str(cfg),
                "--libro", str(book), "--sin-respaldo")
        run("evm_export.py", "actualizar", "--origen", f"P1={progress}", "--origen", f"P2={progress2}",
            "--config", str(several_cfg), "--libro", str(books[2]), "--sin-respaldo")
    return books


def find_pdftoppm() -> str:
    candidates = [os.environ.get("PDFTOPPM", ""), shutil.which("pdftoppm") or ""]
    pattern = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WinGet\Packages\*Poppler*\**\pdftoppm.exe")
    candidates += sorted(glob.glob(pattern, recursive=True))
    for c in candidates:
        if c and Path(c).is_file():
            return c
    raise SystemExit("No se encontró pdftoppm (Poppler). Se puede instalar con 'winget install oschwartz10612.Poppler' "
                     "o indicar su ruta en la variable de entorno PDFTOPPM.")


def trim(img, pad: int = 12):
    from PIL import Image, ImageChops
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    box = ImageChops.difference(img, bg).convert("L").point(lambda v: 255 if v > 8 else 0).getbbox()
    if not box:
        return img
    left, top, right, bottom = box
    return img.crop((max(0, left - pad), max(0, top - pad), min(img.width, right + pad),
                     min(img.height, bottom + pad)))


def save_png(image, path: Path) -> None:
    """PNG con paleta de 256 colores: un tercio del tamaño, sin diferencia visible en capturas."""
    from PIL import Image
    image.quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).save(path, optimize=True)


def spanish_dates(ws) -> None:
    """Meses en español en las celdas y en los ejes de los gráficos ([$-C0A] = español)."""
    for cell in ws.UsedRange.Cells:
        fmt = cell.NumberFormat
        if isinstance(fmt, str) and "mmm" in fmt.lower() and not fmt.startswith("[$"):
            cell.NumberFormat = "[$-C0A]" + fmt
    for chart_object in ws.ChartObjects():
        for axis_type in (1, 2):
            try:
                axis = chart_object.Chart.Axes(axis_type)
                fmt = axis.TickLabels.NumberFormat
            except Exception:  # el gráfico no tiene ese eje
                continue
            if isinstance(fmt, str) and "mmm" in fmt.lower() and not fmt.startswith("[$"):
                axis.TickLabels.NumberFormat = "[$-C0A]" + fmt


def export(book: Path, items: list, pdftoppm: str, tmp: Path) -> None:
    import pythoncom
    import win32com.client
    from PIL import Image

    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    original = (xl.UseSystemSeparators, xl.DecimalSeparator, xl.ThousandsSeparator)
    try:
        xl.UseSystemSeparators = False
        xl.DecimalSeparator = ","
        xl.ThousandsSeparator = "."
        wb = xl.Workbooks.Open(str(book), UpdateLinks=0, ReadOnly=True)
        prepared = set()
        for name, sheet, area, opts in items:
            ws = wb.Worksheets(sheet)
            if sheet not in prepared:
                spanish_dates(ws)
                prepared.add(sheet)
            if "hide" in opts:
                ws.Range(opts["hide"]).EntireColumn.Hidden = True
            if "hide_rows" in opts:
                ws.Range(opts["hide_rows"]).EntireRow.Hidden = True
            if "width" in opts:
                ws.Columns(opts["width"][0]).ColumnWidth = opts["width"][1]
            if "date" in opts:
                cell, value = opts["date"]
                ws.Range(cell).Value = value.replace(tzinfo=dt.timezone.utc)
            if "value" in opts:
                cell, value = opts["value"]
                ws.Range(cell).Value = value
            for cell, value in opts.get("values", []):
                ws.Range(cell).Value = value.replace(tzinfo=dt.timezone.utc) if isinstance(value, dt.datetime) else value
            xl.CalculateFull()
            xl.PrintCommunication = False
            ps = ws.PageSetup
            ps.PrintArea = area
            ps.Orientation = 2 if opts.get("landscape") else 1
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.PrintHeadings = bool(opts.get("headings"))
            ps.LeftHeader = ps.CenterHeader = ps.RightHeader = ""
            ps.LeftFooter = ps.CenterFooter = ps.RightFooter = ""
            xl.PrintCommunication = True
            pdf = tmp / f"{name}.pdf"
            ws.ExportAsFixedFormat(0, str(pdf), 0, False, False)
            subprocess.run([pdftoppm, "-r", str(DPI), "-png", "-singlefile", str(pdf), str(tmp / name)], check=True)
            save_png(trim(Image.open(tmp / f"{name}.png").convert("RGB")), OUT / f"{name}.png")
            print(f"  {name}.png")
            if "hide" in opts:
                ws.Range(opts["hide"]).EntireColumn.Hidden = False
            if "hide_rows" in opts:
                ws.Range(opts["hide_rows"]).EntireRow.Hidden = False
        wb.Close(SaveChanges=False)
    finally:
        xl.UseSystemSeparators, xl.DecimalSeparator, xl.ThousandsSeparator = original
        xl.Quit()


def annotate() -> None:
    from PIL import Image, ImageDraw, ImageFont
    image = Image.open(OUT / "cuadro-de-mando.png").convert("RGB")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("C:/Windows/Fonts/segoeuib.ttf", 17)
    for n, x, y in MARKS:
        draw.ellipse((x - 14, y - 14, x + 14, y + 14), fill=(192, 0, 0), outline=(255, 255, 255), width=2)
        width = draw.textlength(str(n), font=font)
        draw.text((x - width / 2, y - 12), str(n), fill="white", font=font)
    save_png(image, OUT / "cuadro-de-mando-anotado.png")
    print("  cuadro-de-mando-anotado.png")


def main() -> None:
    pdftoppm = find_pdftoppm()
    OUT.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp_dir:
        work = Path(tmp_dir)
        print("Generando los libros del ejemplo (24 meses)...")
        single, multi, several = build_books(work)
        print("Exportando capturas a", OUT)
        export(single, SINGLE, pdftoppm, work)
        export(multi, MULTI, pdftoppm, work)
        export(several, SEVERAL, pdftoppm, work)
        annotate()


if __name__ == "__main__":
    main()
