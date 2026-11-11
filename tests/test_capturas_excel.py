"""Herramienta de capturas del libro (docs/herramientas/capturas_excel.py): reintentos ante errores transitorios de
Excel y carpeta de trabajo (ticket #18). Las capturas con Excel no se prueban aquí: requieren Excel y Poppler."""
import datetime as dt
import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("capturas_excel", ROOT / "docs" / "herramientas" / "capturas_excel.py")
capturas = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(capturas)

OBJECT_REQUIRED = -2146827864           # 0x800A01A8, el error del ticket #18
CALL_REJECTED = -2147418111             # RPC_E_CALL_REJECTED
EXCEL_ERROR = -2146827284               # 0x800A03EC: error de Excel por un dato no válido, no transitorio


class FakeComError(Exception):
    """Como pywintypes.com_error: (hresult, texto, excepinfo, argumento)."""


def failing(errors, result="ok"):
    """Acción que lanza los errores indicados, uno por llamada, y después devuelve result."""
    calls = []

    def action():
        calls.append(1)
        if len(calls) <= len(errors):
            raise errors[len(calls) - 1]
        return result
    return action, calls


def test_reintenta_el_error_del_ticket_y_termina_bien(capsys):
    waits = []
    action, calls = failing([FakeComError(OBJECT_REQUIRED, "OLE error 0x800a01a8", None, None)] * 2)
    assert capturas.with_retries(action, "cdm-indicadores.png", sleep=waits.append) == "ok"
    assert len(calls) == 3 and waits == [1, 2]
    assert "cdm-indicadores.png (0x800A01A8): nuevo intento en 1 s" in capsys.readouterr().out


def test_reintenta_la_excepcion_de_excel_con_un_codigo_transitorio():
    waits = []
    excepinfo = (0, "Microsoft Excel", "Llamada rechazada", None, 0, CALL_REJECTED)
    action, calls = failing([FakeComError(capturas.DISP_E_EXCEPTION, "Excepción", excepinfo, None)])
    assert capturas.with_retries(action, "pv.png", sleep=waits.append) == "ok"
    assert len(calls) == 2 and waits == [1]


@pytest.mark.parametrize("error", [FakeComError(EXCEL_ERROR, "OLE error 0x800a03ec", None, None),
                                   ValueError("otro error")])
def test_los_demas_errores_no_se_reintentan(error):
    waits = []
    action, calls = failing([error])
    with pytest.raises(type(error)):
        capturas.with_retries(action, "pv.png", sleep=waits.append)
    assert len(calls) == 1 and waits == []


def test_sin_respuesta_despues_de_todos_los_intentos():
    waits = []
    action, calls = failing([FakeComError(OBJECT_REQUIRED, "OLE error 0x800a01a8", None, None)] * 10)
    with pytest.raises(SystemExit, match="cdm-indicadores.png: error 0x800A01A8 en los 5 intentos.*--carpeta-trabajo"):
        capturas.with_retries(action, "cdm-indicadores.png", sleep=waits.append)
    assert len(calls) == 5 and waits == list(capturas.RETRY_WAITS)


def test_codigo_de_error_com():
    assert capturas.com_error_code(FakeComError(OBJECT_REQUIRED, "texto", None, None)) == OBJECT_REQUIRED
    assert capturas.com_error_code(ValueError("texto")) is None
    excepinfo = (0, "Microsoft Excel", "texto", None, 0, EXCEL_ERROR)
    assert capturas.com_error_code(FakeComError(capturas.DISP_E_EXCEPTION, "texto", excepinfo, None)) == EXCEL_ERROR


def test_subcarpeta_nueva_en_la_carpeta_de_trabajo(tmp_path):
    folder = capturas.work_folder(tmp_path / "capturas", dt.datetime(2026, 11, 11, 9, 30, 5))
    assert folder == (tmp_path / "capturas" / "capturas-20261111-093005").resolve() and folder.is_dir()
    other = capturas.work_folder(tmp_path / "capturas", dt.datetime(2026, 11, 11, 9, 31, 0))
    assert other != folder and other.is_dir()
