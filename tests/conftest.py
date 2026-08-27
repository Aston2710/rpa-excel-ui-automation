"""Dobles de prueba que emulan los controles de UI Automation.

Permiten ejercitar la logica de `FileExplorer` sin abrir Microsoft Excel:
solo se reproduce el contrato que la clase realmente consume (`Exists`,
`Disappears`, `GetPattern` y los constructores de controles hijos).
"""

from __future__ import annotations

import pytest
import uiautomation as auto


class FakePattern:
    """Registra las acciones invocadas sobre un control."""

    def __init__(self, control: "FakeControl") -> None:
        self._control = control
        self.Value = ""

    def Invoke(self) -> None:
        self._control.invocations.append("Invoke")

    def Select(self) -> None:
        self._control.invocations.append("Select")

    def DoDefaultAction(self) -> None:
        self._control.invocations.append("DoDefaultAction")

    def SetValue(self, value: str) -> None:
        self.Value = value
        self._control.invocations.append(f"SetValue:{value}")


class FakeControl:
    """Control minimo compatible con el uso que hace `FileExplorer`."""

    def __init__(
        self,
        name: str = "",
        automation_id: str = "",
        class_name: str = "",
        exists: bool = True,
        patterns: tuple[int, ...] = (auto.PatternId.InvokePattern,),
    ) -> None:
        self.Name = name
        self.AutomationId = automation_id
        self.ClassName = class_name
        self.invocations: list[str] = []
        self.disappeared = False
        self._exists = exists
        self._patterns = {pattern_id: FakePattern(self) for pattern_id in patterns}
        self.children: list[FakeControl] = []

    # -- Contrato consumido por FileExplorer --------------------------------

    def Exists(self, maxSearchSeconds: float = 5, searchIntervalSeconds: float = 0.5) -> bool:
        return self._exists

    def Disappears(self, maxSearchSeconds: float = 5, searchIntervalSeconds: float = 0.5) -> bool:
        return self.disappeared

    def GetPattern(self, pattern_id: int) -> FakePattern | None:
        return self._patterns.get(pattern_id)

    def SetActive(self, waitTime: float = 0) -> bool:
        self.invocations.append("SetActive")
        return True

    def SetFocus(self) -> bool:
        self.invocations.append("SetFocus")
        return True

    # -- Utilidades del doble ------------------------------------------------

    def add(self, control: "FakeControl") -> "FakeControl":
        self.children.append(control)
        return control

    def _match(self, kind: str, kwargs: dict[str, object]) -> "FakeControl":
        for child in self.children:
            if child.kind != kind:
                continue
            if all(getattr(child, key, None) == value for key, value in _translate(kwargs).items()):
                return child
        return FakeControl(exists=False)

    kind = "Control"

    def Control(self, **kwargs: object) -> "FakeControl":
        return self._match("Control", kwargs)

    def EditControl(self, **kwargs: object) -> "FakeControl":
        return self._match("Edit", kwargs)

    def ButtonControl(self, **kwargs: object) -> "FakeControl":
        return self._match("Button", kwargs)

    def WindowControl(self, **kwargs: object) -> "FakeControl":
        return self._match("Window", kwargs)


def _translate(kwargs: dict[str, object]) -> dict[str, object]:
    """Traduce los argumentos de busqueda de uiautomation a atributos del doble."""
    mapping = {"AutomationId": "AutomationId", "Name": "Name", "ClassName": "ClassName"}
    return {mapping[key]: value for key, value in kwargs.items() if key in mapping}


def make(kind: str, **kwargs: object) -> FakeControl:
    """Crea un `FakeControl` del tipo indicado (Edit, Button, Window, Control)."""
    control = FakeControl(**kwargs)  # type: ignore[arg-type]
    control.kind = kind
    return control


@pytest.fixture
def open_dialog() -> FakeControl:
    """Dialogo "Abrir" con su campo de texto y su boton de confirmacion."""
    dialog = make("Window", name="Abrir")
    dialog.add(
        make("Edit", automation_id="1148", patterns=(auto.PatternId.ValuePattern,))
    )
    dialog.add(make("Control", name="Abrir", automation_id="1", class_name="Button"))
    dialog.disappeared = True
    return dialog


@pytest.fixture
def save_dialog() -> FakeControl:
    """Dialogo "Guardar como" sin advertencia de sobreescritura."""
    dialog = make("Window", name="Guardar como")
    dialog.add(
        make("Edit", automation_id="1001", patterns=(auto.PatternId.ValuePattern,))
    )
    dialog.add(make("Control", name="Guardar", automation_id="1", class_name="Button"))
    dialog.disappeared = True
    return dialog
