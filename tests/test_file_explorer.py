"""Pruebas de la logica de `FileExplorer` sobre dobles de pywinauto."""

from __future__ import annotations

import pytest

from rpa_excel_ui_automation import config
from rpa_excel_ui_automation.file_explorer import (
    DialogControlNotFoundError,
    DialogStillOpenError,
    FileExplorer,
)

from .conftest import FakeSpec


@pytest.fixture
def explorer() -> FileExplorer:
    return FileExplorer(dialog_timeout=0.1, confirm_timeout=0.1)


def edit_of(dialog: FakeSpec) -> FakeSpec:
    return next(child for child in dialog.children if child.control_type == "Edit")


def accept_of(dialog: FakeSpec) -> FakeSpec:
    return next(child for child in dialog.children if child.auto_id == config.ACCEPT_BUTTON_ID)


# --- Caso 01 ---------------------------------------------------------------


def test_open_document_inyecta_ruta_absoluta_y_confirma(explorer, open_dialog, tmp_path):
    source = tmp_path / "origen.xlsx"
    source.write_bytes(b"contenido")

    result = explorer.open_document(open_dialog, source)

    assert result == source.resolve()
    assert edit_of(open_dialog).invocations == [
        "set_focus",
        f"set_edit_text:{source.resolve()}",
    ]
    assert accept_of(open_dialog).invocations == ["invoke"]


def test_open_document_rechaza_un_origen_inexistente(explorer, open_dialog, tmp_path):
    with pytest.raises(FileNotFoundError):
        explorer.open_document(open_dialog, tmp_path / "no-existe.xlsx")

    assert edit_of(open_dialog).invocations == []


def test_open_document_falla_si_el_dialogo_no_se_cierra(explorer, open_dialog, tmp_path):
    source = tmp_path / "origen.xlsx"
    source.write_bytes(b"contenido")
    open_dialog.disappeared = False

    with pytest.raises(DialogStillOpenError):
        explorer.open_document(open_dialog, source)


def test_open_document_falla_sin_campo_de_texto(explorer, tmp_path):
    source = tmp_path / "origen.xlsx"
    source.write_bytes(b"contenido")
    dialog = FakeSpec(title="Abrir", class_name=config.DIALOG_CLASS, control_type="Window")
    dialog.add(FakeSpec(title="Abrir", auto_id="1", class_name="Button", control_type="SplitButton"))

    with pytest.raises(DialogControlNotFoundError):
        explorer.open_document(dialog, source)


def test_el_boton_de_confirmacion_no_se_confunde_con_un_archivo_de_la_carpeta(
    explorer, tmp_path
):
    """El explorador numera los elementos listados con auto_id "0", "1", "2"...

    Sin filtrar por `class_name`, `auto_id="1"` tambien identifica al segundo
    archivo de la carpeta mostrada y la busqueda resulta ambigua.
    """
    source = tmp_path / "origen.xlsx"
    source.write_bytes(b"contenido")

    dialog = FakeSpec(title="Abrir", class_name=config.DIALOG_CLASS, control_type="Window")
    dialog.add(FakeSpec(auto_id="1148", control_type="Edit", patterns=()))
    homonimo = dialog.add(
        FakeSpec(title="bovedas", auto_id="1", class_name="UIItem", control_type="ListItem")
    )
    boton = dialog.add(
        FakeSpec(title="Abrir", auto_id="1", class_name="Button", control_type="SplitButton")
    )
    dialog.disappeared = True

    explorer.open_document(dialog, source)

    assert boton.invocations == ["invoke"]
    assert homonimo.invocations == []


# --- Caso 02 ---------------------------------------------------------------


def test_save_document_crea_el_directorio_destino(explorer, save_dialog, tmp_path):
    target = tmp_path / "output" / "destino.xlsx"

    result = explorer.save_document(save_dialog, target)

    assert result.parent.is_dir()
    assert edit_of(save_dialog).invocations == [
        "set_focus",
        f"set_edit_text:{target.resolve()}",
    ]
    assert accept_of(save_dialog).invocations == ["invoke"]


def test_save_document_confirma_el_reemplazo_cuando_el_destino_ya_existe(
    explorer, save_dialog, tmp_path
):
    warning = FakeSpec(
        title="Confirmar Guardar como",
        class_name=config.DIALOG_CLASS,
        control_type="Window",
    )
    yes_button = warning.add(
        FakeSpec(title="Sí", auto_id=config.OVERWRITE_YES_ID, control_type="Button")
    )
    warning.disappeared = True
    save_dialog.add(warning)

    target = tmp_path / "destino.xlsx"
    target.write_bytes(b"anterior")

    explorer.save_document(save_dialog, target)

    assert yes_button.invocations == ["invoke"]


def test_save_document_no_busca_boton_afirmativo_si_no_hay_advertencia(explorer, save_dialog):
    assert explorer._resolve_overwrite_warning(save_dialog) is False


def test_boton_afirmativo_se_localiza_por_nombre_si_falta_el_automation_id(explorer):
    warning = FakeSpec(
        title="Confirmar Guardar como",
        class_name=config.DIALOG_CLASS,
        control_type="Window",
    )
    yes_button = warning.add(
        FakeSpec(
            title="Sí",
            class_name=config.OVERWRITE_BUTTON_CLASS,
            control_type="Button",
        )
    )

    assert explorer._find_confirmation_button(warning) is yes_button


# --- Accionamiento sin coordenadas -----------------------------------------


def test_invoke_usa_el_patron_heredado_cuando_no_hay_invoke():
    control = FakeSpec(title="Sí", patterns=("legacy",))

    FileExplorer._invoke(control)

    assert control.invocations == ["DoDefaultAction"]


def test_invoke_falla_si_el_control_no_es_accionable():
    control = FakeSpec(title="inerte", patterns=())

    with pytest.raises(DialogControlNotFoundError):
        FileExplorer._invoke(control)
