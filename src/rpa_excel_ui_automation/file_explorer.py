"""Interaccion con los cuadros de dialogo nativos del explorador de Windows."""

from __future__ import annotations

import logging
from pathlib import Path

import uiautomation as auto

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

    def open_document(self, dialog: auto.WindowControl, source: Path) -> Path:
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

    def save_document(self, dialog: auto.WindowControl, target: Path) -> Path:
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

    def _inject_path(self, dialog: auto.WindowControl, path: Path) -> None:
        """Escribe la ruta absoluta en el campo "Nombre de archivo" del dialogo.

        Se direcciona el control por su `AutomationId` y se usa el patron de
        valor de UI Automation, sin desplazamientos con `Tab` ni tecleo ciego.
        """
        edit = self._find_file_name_edit(dialog)
        value = edit.GetPattern(auto.PatternId.ValuePattern)
        if value is None:
            raise DialogControlNotFoundError(
                "El campo 'Nombre de archivo' no expone ValuePattern."
            )

        if not edit.SetFocus():
            logger.warning("El control Edit %r no acepto el foco.", edit.AutomationId)
        value.SetValue(str(path))
        logger.info("Ruta inyectada en el control Edit %r: %s", edit.AutomationId, value.Value)

    def _find_file_name_edit(self, dialog: auto.WindowControl) -> auto.EditControl:
        """Localiza el campo "Nombre de archivo" probando sus AutomationId conocidos."""
        for timeout in (config.CONTROL_PROBE_TIMEOUT, self._dialog_timeout):
            for automation_id in config.FILE_NAME_EDIT_IDS:
                edit = dialog.EditControl(
                    searchDepth=config.DIALOG_SEARCH_DEPTH, AutomationId=automation_id
                )
                if edit.Exists(timeout, config.POLL_INTERVAL):
                    return edit

        raise DialogControlNotFoundError(
            "No se encontro el campo 'Nombre de archivo' con los AutomationId "
            f"{config.FILE_NAME_EDIT_IDS} en el dialogo {dialog.Name!r}."
        )

    def _accept(self, dialog: auto.WindowControl) -> None:
        """Acciona el boton de confirmacion (Abrir / Guardar) del dialogo.

        Se filtra tambien por `ClassName` para que el selector siga siendo
        univoco al margen del alcance de la busqueda: el `AutomationId` por si
        solo colisiona con los elementos de la carpeta listada.
        """
        button = dialog.Control(
            searchDepth=1,
            AutomationId=config.ACCEPT_BUTTON_ID,
            ClassName=config.ACCEPT_BUTTON_CLASS,
        )
        if not button.Exists(self._dialog_timeout, config.POLL_INTERVAL):
            raise DialogControlNotFoundError(
                f"El boton de confirmacion (AutomationId={config.ACCEPT_BUTTON_ID}, "
                f"ClassName={config.ACCEPT_BUTTON_CLASS}) no existe en el dialogo "
                f"{dialog.Name!r}."
            )

        logger.info("Accionando el boton %r del dialogo %r", button.Name, dialog.Name)
        self._invoke(button)

    def _resolve_overwrite_warning(self, dialog: auto.WindowControl) -> bool:
        """Detecta y confirma la advertencia "El archivo ya existe".

        La deteccion es dinamica mediante `.Exists()`; si la ventana no aparece
        el flujo continua sin interrupcion. Devuelve True si hubo reemplazo.
        """
        warning = dialog.WindowControl(searchDepth=1, ClassName=config.DIALOG_CLASS)
        if not warning.Exists(self._confirm_timeout, config.POLL_INTERVAL):
            logger.info("Sin advertencia de sobreescritura; el destino era nuevo.")
            return False

        logger.warning("Sobreescritura detectada: %r. Confirmando reemplazo.", warning.Name)
        self._invoke(self._find_confirmation_button(warning))
        warning.Disappears(self._dialog_timeout, config.POLL_INTERVAL)
        logger.info("Reemplazo confirmado sin intervencion humana.")
        return True

    def _find_confirmation_button(self, warning: auto.WindowControl) -> auto.Control:
        """Localiza el boton afirmativo del TaskDialog de sobreescritura."""
        button = warning.ButtonControl(
            searchDepth=config.DIALOG_SEARCH_DEPTH, AutomationId=config.OVERWRITE_YES_ID
        )
        if button.Exists(config.CONTROL_PROBE_TIMEOUT, config.POLL_INTERVAL):
            return button

        for name in config.OVERWRITE_YES_NAMES:
            button = warning.ButtonControl(
                searchDepth=config.DIALOG_SEARCH_DEPTH,
                ClassName=config.OVERWRITE_BUTTON_CLASS,
                Name=name,
            )
            if button.Exists(config.CONTROL_PROBE_TIMEOUT, config.POLL_INTERVAL):
                return button

        raise DialogControlNotFoundError(
            f"No se encontro el boton de confirmacion en {warning.Name!r}."
        )

    def _wait_until_closed(self, dialog: auto.WindowControl, description: str) -> None:
        """Espera a que el dialogo desaparezca usando el mecanismo nativo."""
        if not dialog.Disappears(self._dialog_timeout, config.POLL_INTERVAL):
            raise DialogStillOpenError(
                f"El dialogo '{description}' sigue abierto tras {self._dialog_timeout}s."
            )
        logger.info("El dialogo '%s' se cerro.", description)

    @staticmethod
    def _invoke(control: auto.Control) -> None:
        """Acciona un control por patron de UI Automation, nunca por coordenadas."""
        invoke = control.GetPattern(auto.PatternId.InvokePattern)
        if invoke is not None:
            invoke.Invoke()
            return

        selection = control.GetPattern(auto.PatternId.SelectionItemPattern)
        if selection is not None:
            selection.Select()
            return

        legacy = control.GetPattern(auto.PatternId.LegacyIAccessiblePattern)
        if legacy is not None:
            legacy.DoDefaultAction()
            return

        raise DialogControlNotFoundError(
            f"El control {control.Name!r} no expone ningun patron accionable."
        )
