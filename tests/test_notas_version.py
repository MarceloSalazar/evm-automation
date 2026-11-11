"""Notas de las publicaciones de GitHub generadas desde el CHANGELOG (docs/herramientas/notas_version.py)."""
import importlib.util
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("notas_version", ROOT / "docs" / "herramientas" / "notas_version.py")
notas = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(notas)

CHANGELOG = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
VERSIONS = re.findall(r"^## \[(\d+\.\d+\.\d+)\]", CHANGELOG, re.M)
REPO = notas.REPO

EJEMPLO = """# Registro de cambios

## [Sin publicar]

### Nuevas funciones

- Ninguna.

<a id="v1.1.0"></a>

## [1.1.0] - 2026-01-02 · Segunda versión

### Resumen

**Nuevas funciones**

- **Función nueva:** se explica en una línea, con su documento
  ([Comandos](docs/comandos.md#generar)).

**Otros**

- Instalador adjunto, ver la [versión anterior](https://example.com/v1).

### Nuevas funciones

- **Función nueva:** detalle largo que sigue
  en otra línea
  - con una viñeta anidada
    que también sigue.

### Correcciones

- Ninguna.

### Otros

- Instalador adjunto.

<a id="v1.0.0"></a>

## [1.0.0] - 2026-01-01 · Primera versión

### Resumen

**Nuevas funciones**

- Primera versión ([El libro](README.md#libro)).

### Nuevas funciones

- Todo.

### Correcciones

- Ninguna.

### Otros

- Ninguna.

[Sin publicar]: https://github.com/MarceloSalazar/evm-automation/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/MarceloSalazar/evm-automation/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/MarceloSalazar/evm-automation/releases/tag/v1.0.0
"""


@pytest.mark.parametrize("version", VERSIONS)
def test_cada_version_del_changelog_tiene_resumen_para_la_publicacion(version):
    notes = notas.release_notes(version, CHANGELOG)
    assert re.match(r"### (Nuevas funciones|Correcciones|Otros)\n", notes)
    assert "- " in notes
    # Ningún enlace relativo: en la publicación de GitHub no llevarían al documento.
    assert not re.search(r"\]\((?!https?://)", notes)
    assert f"]({REPO}/blob/main/CHANGELOG.md#v{version})" in notes
    # Una línea por viñeta: GitHub muestra cada salto de línea de las notas.
    assert not re.search(r"^ +\S", notes, re.M)
    assert notas.release_title(version, CHANGELOG).startswith(f"v{version} · ")
    detail = notas.detailed_notes(version, CHANGELOG)
    assert "### Resumen" not in detail and "### Nuevas funciones" in detail


def test_resumen_con_enlaces_completos_y_pie():
    notes = notas.release_notes("1.1.0", EJEMPLO)
    assert notes == (
        "### Nuevas funciones\n\n"
        f"- **Función nueva:** se explica en una línea, con su documento ([Comandos]({REPO}/blob/v1.1.0/"
        "docs/comandos.md#generar)).\n\n"
        "### Otros\n\n"
        "- Instalador adjunto, ver la [versión anterior](https://example.com/v1).\n\n"
        "---\n\n"
        f"**Más información:** [notas detalladas de la versión]({REPO}/blob/main/CHANGELOG.md#v1.1.0) · "
        f"[cambios en el código desde la versión 1.0.0]({REPO}/compare/v1.0.0...v1.1.0)\n")


def test_primera_version_enlaza_con_su_codigo():
    notes = notas.release_notes("1.0.0", EJEMPLO)
    assert f"[El libro]({REPO}/blob/v1.0.0/README.md#libro)" in notes
    assert f"[código de la versión]({REPO}/tree/v1.0.0)" in notes
    # La última entrada termina antes de los enlaces de comparación del final.
    assert "[Sin publicar]:" not in notas.detailed_notes("1.0.0", EJEMPLO)


def test_detalle_sin_resumen_y_con_vinetas_anidadas():
    detail = notas.detailed_notes("1.1.0", EJEMPLO)
    assert detail.startswith("### Nuevas funciones\n\n- **Función nueva:** detalle largo que sigue en otra línea\n"
                             "  - con una viñeta anidada que también sigue.\n")
    assert "### Resumen" not in detail and "### Otros\n\n- Instalador adjunto." in detail


@pytest.mark.parametrize("cambio, mensaje", [
    (('<a id="v1.1.0"></a>\n\n', ""), "falta el ancla"),
    (("### Resumen\n\n**Nuevas funciones**\n\n- **Función nueva:** se explica", "- **Función nueva:** se explica"),
     "falta la sección '### Resumen'"),
    (("**Otros**\n\n- Instalador adjunto, ver", "**Varios**\n\n- Instalador adjunto, ver"), "no reconocida"),
    (("### Correcciones\n\n- Ninguna.\n\n### Otros\n\n- Instalador", "### Otros\n\n- Instalador"),
     "falta la sección '### Correcciones'"),
])
def test_errores_de_la_entrada(cambio, mensaje):
    texto = EJEMPLO.replace(*cambio, 1)
    assert texto != EJEMPLO
    with pytest.raises(SystemExit, match=re.escape(mensaje)):
        notas.release_notes("1.1.0", texto)


def test_version_que_no_esta():
    with pytest.raises(SystemExit, match="no está"):
        notas.release_notes("9.9.9", EJEMPLO)
