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

from .conftest import FakeSpec


@pytest.fixture
def manager(tmp_path: Path) -> ExcelManager:
    executable = tmp_path / "EXCEL.EXE"
    executable.write_bytes(b"")
    return ExcelManager(executable=executable, app_timeout=0.1, dialog_timeout=0.1)


@pytest.fixture
def keys(monkeypatch) -> list[str]:
    """Captura los atajos enviados en lugar de teclearlos en el escritorio."""
    enviados: list[str] = []
    monkeypatch.setattr(module, "send_keys", enviados.append)
    return enviados


def excel_window(title: str, dialog_title: str | None = None) -> FakeSpec:
    window = FakeSpec(title=title, class_name=config.EXCEL_WINDOW_CLASS, control_type="Window")
    if dialog_title is not None:
        window.add(
            FakeSpec(
                title=dialog_title,
                class_name=config.DIALOG_CLASS,
                control_type="Window",
            )
        )
    return window


def test_window_exige_haber_iniciado_la_aplicacion(manager):
    with pytest.raises(ExcelNotAvailableError):
        _ = manager.window


def test_open_file_devuelve_el_dialogo_desplegado(manager, keys):
    manager._window = excel_window("Excel", dialog_title="Abrir")

    dialogo = manager.open_file()

    assert dialogo.window_text() == "Abrir"
    assert keys == [config.SHORTCUT_OPEN]


def test_save_as_usa_el_atajo_global_f12(manager, keys):
    manager._window = excel_window("origen.xlsx - Excel", dialog_title="Guardar como")

    manager.save_as()

    assert keys == [config.SHORTCUT_SAVE_AS]


def test_la_especificacion_devuelta_incluye_el_titulo_observado(manager, keys):
    """Sin el titulo, la especificacion deja de ser univoca cuando Excel

    despliega la advertencia de sobreescritura: pasan a existir dos ventanas
    `#32770` bajo la misma ventana padre.
    """
    window = excel_window("origen.xlsx - Excel", dialog_title="Guardar como")
    advertencia = window.add(
        FakeSpec(
            title="Confirmar Guardar como",
            class_name=config.DIALOG_CLASS,
            control_type="Window",
        )
    )

    manager._window = window
    dialogo = manager.save_as()

    assert dialogo.window_text() == "Guardar como"
    assert dialogo is not advertencia


def test_el_atajo_se_reintenta_y_luego_falla_de_forma_explicita(manager, keys):
    manager._window = excel_window("Excel")  # sin dialogo hijo

    with pytest.raises(DialogNotRaisedError):
        manager.save_as()

    assert len(keys) == config.SHORTCUT_ATTEMPTS


def test_localizar_el_ejecutable_falla_cuando_excel_no_esta_instalado(monkeypatch):
    monkeypatch.setattr(config, "EXCEL_FALLBACK_PATHS", ())
    monkeypatch.setattr(
        module.winreg, "OpenKey", lambda *args, **kwargs: (_ for _ in ()).throw(OSError())
    )

    with pytest.raises(ExcelNotAvailableError):
        ExcelManager._locate_executable()
