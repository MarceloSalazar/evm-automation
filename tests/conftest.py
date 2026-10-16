import importlib.util
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
    generator.write(generator.build_program(), files["base"], None)
    generator.write(generator.build_program(), files["jun"], date(2027, 6, 30))
    generator.write(generator.build_program(), files["jul"], date(2027, 7, 31))
    cfg = d / "config.toml"
    cfg.write_text('[campos]\nmetodo = "Método EVM"\npeso_hito = "Peso hito"\n', encoding="utf-8")
    files["config"] = cfg
    return files


def find_soffice():
    for c in (shutil.which("soffice"), r"C:\Program Files\LibreOffice\program\soffice.exe",
              "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if c and Path(c).exists():
            return c
    return None


@pytest.fixture(scope="session")
def recalc():
    """Devuelve una función que recalcula un .xlsx con LibreOffice (o salta la prueba)."""
    soffice = find_soffice()
    if not soffice:
        pytest.skip("LibreOffice no está instalado: no se verifican las fórmulas")

    def run(path: Path) -> Path:
        out = path.parent / "recalc"
        out.mkdir(exist_ok=True)
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                       check=True, capture_output=True, timeout=240)
        return out / path.name

    return run
