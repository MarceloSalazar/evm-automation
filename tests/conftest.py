import contextlib
import importlib.util
import io
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def _load_generator():
    spec = importlib.util.spec_from_file_location("generar_ejemplo", ROOT / "ejemplos" / "generar_ejemplo.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules["generar_ejemplo"] = mod
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def generator():
    return _load_generator()


@pytest.fixture(scope="session")
def example_files(tmp_path_factory, generator):
    """XML de ejemplo: línea base y avances a junio y julio de 2027."""
    d = tmp_path_factory.mktemp("ejemplo")
    files = {"base": d / "base.xml", "jun": d / "jun.xml", "jul": d / "jul.xml"}
    generator.write(generator.build_plans(), files["base"], None)
    generator.write(generator.build_plans(), files["jun"], date(2027, 6, 30))
    generator.write(generator.build_plans(), files["jul"], date(2027, 7, 31))
    cfg = d / "config.toml"
    cfg.write_text('[campos]\nmetodo = "Método EVM"\npeso_hito = "Peso hito"\n\n'
                   '[exportacion]\nplanificacion = "INTERNO"\n', encoding="utf-8")
    files["config"] = cfg
    return files


@pytest.fixture(scope="session")
def example_template(tmp_path_factory, example_files):
    """Libro del proyecto de ejemplo, sin BAC, generado una sola vez: las pruebas usan una copia."""
    out = tmp_path_factory.mktemp("plantilla") / "EVM.xlsx"
    quiet_main(["generar", "--origen", str(example_files["base"]), "--config", str(example_files["config"]),
                "--salida", str(out)])
    return out


@pytest.fixture(scope="session")
def example_bac_template(tmp_path_factory, example_template, generator):
    """El mismo libro con el BAC de cada fila del ejemplo (se ingresa a mano)."""
    out = tmp_path_factory.mktemp("plantilla_bac") / "EVM.xlsx"
    shutil.copy(example_template, out)
    generator.fill_book(out)
    return out


@pytest.fixture(scope="session")
def example_jun_template(tmp_path_factory, example_bac_template, example_files, generator):
    """El libro del ejemplo después del cierre de junio de 2027: AC de junio cargado a mano y 'actualizar'."""
    out = tmp_path_factory.mktemp("plantilla_jun") / "EVM.xlsx"
    shutil.copy(example_bac_template, out)
    generator.fill_book(out, date(2027, 6, 30))
    quiet_main(["actualizar", "--origen", str(example_files["jun"]), "--config", str(example_files["config"]),
                "--libro", str(out), "--sin-respaldo"])
    return out


@pytest.fixture()
def book(tmp_path, example_bac_template):
    """Copia, en la carpeta de la prueba, del libro del ejemplo con el BAC de cada fila."""
    out = tmp_path / "EVM.xlsx"
    shutil.copy(example_bac_template, out)
    return out


@pytest.fixture()
def book_jun(tmp_path, example_jun_template):
    """Copia, en la carpeta de la prueba, del libro del ejemplo actualizado a junio de 2027."""
    out = tmp_path / "EVM.xlsx"
    shutil.copy(example_jun_template, out)
    return out


def quiet_main(argv: list[str]) -> None:
    """Ejecuta un comando para armar una plantilla, sin mostrar su salida: la plantilla se arma durante la primera
    prueba que la usa, y su salida no debe aparecer en la de esa prueba (capsys)."""
    from evm.cli import main
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
        code = main(argv)
    assert code == 0, out.getvalue()


def find_soffice():
    for c in (shutil.which("soffice"), r"C:\Program Files\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if c and Path(c).exists():
            return c
    return None


@pytest.fixture(scope="session")
def recalc(tmp_path_factory):
    """Devuelve una función que recalcula un .xlsx con LibreOffice (o salta la prueba).

    Se usa un perfil de LibreOffice propio de la sesión de pruebas: con el perfil del
    usuario, si LibreOffice está abierto, la conversión termina sin error pero no se hace.
    """
    soffice = find_soffice()
    if not soffice:
        pytest.skip("LibreOffice no está instalado: no se verifican las fórmulas")
    profile = tmp_path_factory.mktemp("perfil_libreoffice").as_uri()

    def run(path: Path) -> Path:
        out = path.parent / "recalc"
        out.mkdir(exist_ok=True)
        result = out / path.name
        result.unlink(missing_ok=True)
        proc = subprocess.run([soffice, f"-env:UserInstallation={profile}", "--headless", "--calc",
                               "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                              capture_output=True, timeout=240)
        if not result.exists():
            raise RuntimeError(f"LibreOffice no recalculó {path.name}: {proc.stderr.decode(errors='replace')}")
        return result

    return run


def pytest_collection_modifyitems(items):
    """Marca como 'libreoffice' las pruebas que recalculan libros (se omiten con -m "not libreoffice")."""
    for item in items:
        if "recalc" in getattr(item, "fixturenames", ()):
            item.add_marker(pytest.mark.libreoffice)
