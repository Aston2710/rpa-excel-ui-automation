"""Interaccion con los cuadros de dialogo nativos del explorador de Windows."""

from __future__ import annotations

import logging
from pathlib import Path

from pywinauto.application import WindowSpecification
from pywinauto.timings import TimeoutError as WaitTimeoutError
from pywinauto.uia_defines import NoPatternInterfaceError

from . import config

logger = logging.getLogger(__name__)


class DialogControlNotFoundError(RuntimeError):
    """Un control esperado del cuadro de dialogo no esta presente."""


class DialogStillOpenError(RuntimeError):
    """El cuadro de dialogo no se cerro tras confirmar la accion."""


class FileExplorer:
    """Opera exclusivamente los dialogos nativos "Abrir" y "Guardar como".

    Responsabilidad unica: localizar los controles del explorador de archivos
    por su identificador de accesibilidad, inyectar rutas absolutas y resolver
    las ventanas modales de advertencia. No conoce ni manipula la aplicacion
    Excel: recibe la ventana de dialogo ya desplegada por `ExcelManager`.
    """

    def __init__(
        self,
        dialog_timeout: float = config.DIALOG_TIMEOUT,
        confirm_timeout: float = config.CONFIRM_TIMEOUT,
    ) -> None:
        self._dialog_timeout = dialog_timeout
        self._confirm_timeout = confirm_timeout

    # -- Acciones publicas --------------------------------------------------

    def open_document(self, dialog: WindowSpecification, source: Path) -> Path:
        """Inyecta `source` en el dialogo "Abrir" y confirma la seleccion.

        Devuelve la ruta absoluta efectivamente enviada al explorador.
        """
        absolute = source.resolve()
        logger.info("Abriendo documento: %s", absolute)

        if not absolute.is_file():
            raise FileNotFoundError(f"El archivo de origen no existe: {absolute}")

        self._inject_path(dialog, absolute)
        self._accept(dialog)
        self._wait_until_closed(dialog, "Abrir")

        logger.info("Documento abierto correctamente: %s", absolute.name)
        return absolute

    def save_document(self, dialog: WindowSpecification, target: Path) -> Path:
        """Inyecta `target` en el dialogo "Guardar como", confirma y reemplaza.

        Crea el directorio destino si hace falta y resuelve automaticamente la
        advertencia de sobreescritura. Devuelve la ruta absoluta escrita.
        """
        absolute = target.resolve()
        logger.info("Exportando documento a: %s", absolute)

        absolute.parent.mkdir(parents=True, exist_ok=True)
        existed_before = absolute.exists()
        logger.info("El destino %s antes de guardar.", "ya existia" if existed_before else "no existia")

        self._inject_path(dialog, absolute)
        self._accept(dialog)
        self._resolve_overwrite_warning(dialog)
        self._wait_until_closed(dialog, "Guardar como")

        logger.info("Documento exportado correctamente: %s", absolute)
        return absolute

    # -- Internos -----------------------------------------------------------

    def _inject_path(self, dialog: WindowSpecification, path: Path) -> None:
        """Escribe la ruta absoluta en el campo "Nombre de archivo" del dialogo.

        Se direcciona el control por su `auto_id` y se escribe con el patron de
        valor de UI Automation, sin desplazamientos con `Tab` ni tecleo ciego.
        """
        edit = self._find_file_name_edit(dialog).wrapper_object()
        edit.set_focus()
        edit.set_edit_text(str(path))
        logger.info(
            "Ruta inyectada en el control Edit %r: %s",
            edit.element_info.automation_id,
            edit.get_value(),
        )

    def _find_file_name_edit(self, dialog: WindowSpecification) -> WindowSpecification:
        """Localiza el campo "Nombre de archivo" probando sus `auto_id` conocidos.

        La primera pasada es breve porque el dialogo ya fue confirmado como
        existente: solo sirve para descartar el identificador que no aplica.
        """
        for timeout in (config.CONTROL_PROBE_TIMEOUT, self._dialog_timeout):
            for automation_id in config.FILE_NAME_EDIT_IDS:
                edit = dialog.child_window(auto_id=automation_id, control_type="Edit")
                if edit.exists(timeout=timeout, retry_interval=config.POLL_INTERVAL):
                    return edit

        raise DialogControlNotFoundError(
            "No se encontro el campo 'Nombre de archivo' con los auto_id "
            f"{config.FILE_NAME_EDIT_IDS} en el dialogo {dialog.window_text()!r}."
        )

    def _accept(self, dialog: WindowSpecification) -> None:
        """Acciona el boton de confirmacion (Abrir / Guardar) del dialogo.

        Se filtra por `class_name` y no por `control_type`: el tipo cambia entre
        dialogos (`SplitButton` en "Abrir", `Button` en "Guardar como") y el
        `auto_id` por si solo colisiona con los elementos de la carpeta listada.
        """
        button = dialog.child_window(
            auto_id=config.ACCEPT_BUTTON_ID, class_name=config.ACCEPT_BUTTON_CLASS
        )
        if not button.exists(timeout=self._dialog_timeout, retry_interval=config.POLL_INTERVAL):
            raise DialogControlNotFoundError(
                f"El boton de confirmacion (auto_id={config.ACCEPT_BUTTON_ID}) "
                f"no existe en el dialogo {dialog.window_text()!r}."
            )

        logger.info(
            "Accionando el boton %r del dialogo %r", button.window_text(), dialog.window_text()
        )
        self._invoke(button)

    def _resolve_overwrite_warning(self, dialog: WindowSpecification) -> bool:
        """Detecta y confirma la advertencia "El archivo ya existe".

        La deteccion es dinamica mediante `.exists()`; si la ventana no aparece
        el flujo continua sin interrupcion. Devuelve True si hubo reemplazo.
        """
        warning = dialog.child_window(
            class_name=config.DIALOG_CLASS, control_type="Window"
        )
        if not warning.exists(timeout=self._confirm_timeout, retry_interval=config.POLL_INTERVAL):
            logger.info("Sin advertencia de sobreescritura; el destino era nuevo.")
            return False

        logger.warning(
            "Sobreescritura detectada: %r. Confirmando reemplazo.", warning.window_text()
        )
        self._invoke(self._find_confirmation_button(warning))
        try:
            warning.wait_not(
                "exists", timeout=self._dialog_timeout, retry_interval=config.POLL_INTERVAL
            )
        except WaitTimeoutError:
            logger.warning("La advertencia de sobreescritura sigue visible.")
        logger.info("Reemplazo confirmado sin intervencion humana.")
        return True

    def _find_confirmation_button(self, warning: WindowSpecification) -> WindowSpecification:
        """Localiza el boton afirmativo del TaskDialog de sobreescritura."""
        button = warning.child_window(
            auto_id=config.OVERWRITE_YES_ID, control_type="Button"
        )
        if button.exists(
            timeout=config.CONTROL_PROBE_TIMEOUT, retry_interval=config.POLL_INTERVAL
        ):
            return button

        for name in config.OVERWRITE_YES_NAMES:
            button = warning.child_window(
                title=name, class_name=config.OVERWRITE_BUTTON_CLASS, control_type="Button"
            )
            if button.exists(
                timeout=config.CONTROL_PROBE_TIMEOUT, retry_interval=config.POLL_INTERVAL
            ):
                return button

        raise DialogControlNotFoundError(
            f"No se encontro el boton de confirmacion en {warning.window_text()!r}."
        )

    def _wait_until_closed(self, dialog: WindowSpecification, description: str) -> None:
        """Espera a que el dialogo desaparezca usando el mecanismo nativo."""
        try:
            dialog.wait_not(
                "exists", timeout=self._dialog_timeout, retry_interval=config.POLL_INTERVAL
            )
        except WaitTimeoutError as error:
            raise DialogStillOpenError(
                f"El dialogo '{description}' sigue abierto tras {self._dialog_timeout}s."
            ) from error
        logger.info("El dialogo '%s' se cerro.", description)

    @staticmethod
    def _invoke(control: WindowSpecification) -> None:
        """Acciona un control por patron de UI Automation, nunca por coordenadas.

        Se descartan de forma deliberada `click_input()` y cualquier variante
        basada en el puntero: solo se usan los patrones de accesibilidad.
        """
        wrapper = control.wrapper_object()

        for action in ("invoke", "select"):
            method = getattr(wrapper, action, None)
            if method is None:
                continue
            try:
                method()
                return
            except NoPatternInterfaceError:
                continue

        try:
            wrapper.iface_legacy_iaccessible.DoDefaultAction()
            return
        except NoPatternInterfaceError as error:
            raise DialogControlNotFoundError(
                f"El control {wrapper.window_text()!r} no expone ningun patron accionable."
            ) from error
