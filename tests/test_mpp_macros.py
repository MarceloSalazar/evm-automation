"""Macros de un .mpp leídas sin MS Project (evm.mpp_macros): el archivo compuesto, el flujo PROJECT y los casos en que
no se puede leer."""
from __future__ import annotations

from evm.mpp_macros import macro_modules, modules_in_project_stream

from compound import compound_file, mpp_with_macros, project_stream


def test_modules_of_an_mpp_with_macros(tmp_path):
    path = tmp_path / "plan.mpp"
    path.write_bytes(mpp_with_macros(["Módulo1", "Informes"], classes=["Clase1"]))
    assert macro_modules(path) == ["Módulo1", "Informes", "Clase1"]          # ThisProject no cuenta


def test_mpp_with_only_this_project(tmp_path):
    path = tmp_path / "plan.mpp"
    path.write_bytes(mpp_with_macros([]))
    assert macro_modules(path) == []


def test_mpp_without_vba_project(tmp_path):
    path = tmp_path / "plan.mpp"
    path.write_bytes(compound_file({"   114": {"Props": b"\0" * 100}, "PROJECT": b"Module=NoEsVBA"}))
    assert macro_modules(path) == []                # un flujo PROJECT sin almacenamiento VBA al lado no cuenta


def test_files_that_cannot_be_read(tmp_path):
    text = tmp_path / "plan.mpp"
    text.write_bytes(b"no es un archivo compuesto" * 30)
    assert macro_modules(text) is None
    assert macro_modules(tmp_path / "no_existe.mpp") is None
    broken = tmp_path / "roto.mpp"
    broken.write_bytes(mpp_with_macros(["Módulo1"])[:1024])          # cortado: sectores fuera del archivo
    assert macro_modules(broken) is None


def test_project_stream_keys():
    stream = project_stream(["Módulo1"], ["Clase1"]) + b"\r\nBaseClass=Formulario1\r\nPackage={AC9F2F90}"
    assert modules_in_project_stream(stream) == ["Módulo1", "Clase1", "Formulario1"]
