"""Pruebas de `ExcelManager` que no requieren una instancia real de Excel."""

from __future__ import annotations

from pathlib import Path

import pytest

from rpa_excel_ui_automation import config
from rpa_excel_ui_automation import excel_manager as module
from rpa_excel_ui_automation.excel_manager import (
    DialogNotRaisedError,
    ExcelManager,
    ExcelNotAvailableError,
)

from .conftest import make


@pytest.fixture
def manager(tmp_path: Path) -> ExcelManager:
    executable = tmp_path / "EXCEL.EXE"
    executable.write_bytes(b"")
    return ExcelManager(executable=executable, app_timeout=0.1, dialog_timeout=0.1)


def test_window_exige_haber_iniciado_la_aplicacion(manager):
    with pytest.raises(ExcelNotAvailableError):
        _ = manager.window


def test_open_file_devuelve_el_dialogo_desplegado(manager, monkeypatch):
    window = make("Window", name="Excel")
    dialog = window.add(make("Window", name="Abrir", class_name=config.DIALOG_CLASS))
    manager._window = window

    enviados: list[str] = []
    monkeypatch.setattr(module.auto, "SendKeys", lambda keys, waitTime=0: enviados.append(keys))

    assert manager.open_file() is dialog
    assert enviados == [config.SHORTCUT_OPEN]


def test_save_as_usa_el_atajo_global_f12(manager, monkeypatch):
    window = make("Window", name="origen.xlsx - Excel")
    window.add(make("Window", name="Guardar como", class_name=config.DIALOG_CLASS))
    manager._window = window

    enviados: list[str] = []
    monkeypatch.setattr(module.auto, "SendKeys", lambda keys, waitTime=0: enviados.append(keys))

    manager.save_as()

    assert enviados == [config.SHORTCUT_SAVE_AS]


def test_el_atajo_se_reintenta_y_luego_falla_de_forma_explicita(manager, monkeypatch):
    manager._window = make("Window", name="Excel")  # sin dialogo hijo

    enviados: list[str] = []
    monkeypatch.setattr(module.auto, "SendKeys", lambda keys, waitTime=0: enviados.append(keys))

    with pytest.raises(DialogNotRaisedError):
        manager.save_as()

    assert len(enviados) == config.SHORTCUT_ATTEMPTS


def test_localizar_el_ejecutable_falla_cuando_excel_no_esta_instalado(monkeypatch):
    monkeypatch.setattr(config, "EXCEL_FALLBACK_PATHS", ())
    monkeypatch.setattr(
        module.winreg, "OpenKey", lambda *args, **kwargs: (_ for _ in ()).throw(OSError())
    )

    with pytest.raises(ExcelNotAvailableError):
        ExcelManager._locate_executable()
