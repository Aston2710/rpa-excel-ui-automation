"""Dobles de prueba que emulan las especificaciones de ventana de pywinauto.

Permiten ejercitar la logica de `FileExplorer` sin abrir Microsoft Excel: solo
se reproduce el contrato que la clase realmente consume (`exists`, `wait_not`,
`child_window`, `wrapper_object` y los metodos de accion del envoltorio).
"""

from __future__ import annotations

import pytest
from pywinauto.timings import TimeoutError as WaitTimeoutError
from pywinauto.uia_defines import NoPatternInterfaceError

CRITERIA = ("auto_id", "control_type", "class_name", "title")


class FakeElementInfo:
    def __init__(self, automation_id: str) -> None:
        self.automation_id = automation_id


class FakeLegacyInterface:
    def __init__(self, control: "FakeSpec") -> None:
        self._control = control

    def DoDefaultAction(self) -> None:
        self._control.invocations.append("DoDefaultAction")


class FakeSpec:
    """Doble que actua a la vez como `WindowSpecification` y como envoltorio.

    `wrapper_object()` devuelve el propio objeto, igual que ocurre en la
    practica cuando la especificacion ya resolvio un unico control.
    """

    def __init__(
        self,
        title: str = "",
        auto_id: str = "",
        class_name: str = "",
        control_type: str = "",
        exists: bool = True,
        patterns: tuple[str, ...] = ("invoke",),
    ) -> None:
        self.title = title
        self.auto_id = auto_id
        self.class_name = class_name
        self.control_type = control_type
        self.element_info = FakeElementInfo(auto_id)
        self.invocations: list[str] = []
        self.disappeared = False
        self.value = ""
        self._exists = exists
        self._patterns = patterns
        self.children: list[FakeSpec] = []

    # -- Contrato de WindowSpecification ------------------------------------

    def exists(self, timeout: float = 5, retry_interval: float = 0.5) -> bool:
        return self._exists

    def wait_not(self, condition: str, timeout: float = 5, retry_interval: float = 0.5) -> None:
        if not self.disappeared:
            raise WaitTimeoutError(f"{self.title!r} sigue cumpliendo {condition!r}")

    def window_text(self) -> str:
        return self.title

    def wrapper_object(self) -> "FakeSpec":
        return self

    def child_window(self, **criteria: object) -> "FakeSpec":
        wanted = {key: value for key, value in criteria.items() if key in CRITERIA}
        for child in self.children:
            if all(getattr(child, key) == value for key, value in wanted.items()):
                return child
        return FakeSpec(exists=False)

    # -- Contrato del envoltorio --------------------------------------------

    def set_focus(self) -> "FakeSpec":
        self.invocations.append("set_focus")
        return self

    def set_edit_text(self, text: str) -> None:
        self.value = text
        self.invocations.append(f"set_edit_text:{text}")

    def get_value(self) -> str:
        return self.value

    def invoke(self) -> None:
        self._require("invoke")
        self.invocations.append("invoke")

    def select(self) -> None:
        self._require("select")
        self.invocations.append("select")

    @property
    def iface_legacy_iaccessible(self) -> FakeLegacyInterface:
        self._require("legacy")
        return FakeLegacyInterface(self)

    def _require(self, pattern: str) -> None:
        if pattern not in self._patterns:
            raise NoPatternInterfaceError(f"{pattern} no soportado por {self.title!r}")

    # -- Utilidades del doble ------------------------------------------------

    def add(self, child: "FakeSpec") -> "FakeSpec":
        self.children.append(child)
        return child


@pytest.fixture
def open_dialog() -> FakeSpec:
    """Dialogo "Abrir" con su campo de texto y su boton de confirmacion."""
    dialog = FakeSpec(title="Abrir", class_name="#32770", control_type="Window")
    dialog.add(FakeSpec(auto_id="1148", control_type="Edit", patterns=()))
    dialog.add(FakeSpec(title="Abrir", auto_id="1", class_name="Button", control_type="SplitButton"))
    dialog.disappeared = True
    return dialog


@pytest.fixture
def save_dialog() -> FakeSpec:
    """Dialogo "Guardar como" sin advertencia de sobreescritura."""
    dialog = FakeSpec(title="Guardar como", class_name="#32770", control_type="Window")
    dialog.add(FakeSpec(auto_id="1001", control_type="Edit", patterns=()))
    dialog.add(FakeSpec(title="Guardar", auto_id="1", class_name="Button", control_type="Button"))
    dialog.disappeared = True
    return dialog
