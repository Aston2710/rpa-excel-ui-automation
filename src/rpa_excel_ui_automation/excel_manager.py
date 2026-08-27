"""Gestion de la instancia de Microsoft Excel y de sus comandos de ventana."""

from __future__ import annotations

import logging
import subprocess
import winreg
from pathlib import Path

from pywinauto import Desktop
from pywinauto.application import WindowSpecification
from pywinauto.keyboard import send_keys
from pywinauto.timings import TimeoutError as WaitTimeoutError

from . import config

logger = logging.getLogger(__name__)


class ExcelNotAvailableError(RuntimeError):
    """Excel no esta instalado o su ventana principal nunca aparecio."""


class DialogNotRaisedError(RuntimeError):
    """Excel no desplego el cuadro de dialogo esperado tras el atajo global."""


class ExcelManager:
    """Levanta Excel y dispara los eventos principales de su ventana.

    Responsabilidad unica: el ciclo de vida de la aplicacion y el envio de los
    atajos globales que despliegan los cuadros de dialogo. La interaccion con
    esos cuadros de dialogo se delega por completo a `FileExplorer`.
    """

    def __init__(
        self,
        executable: Path | None = None,
        app_timeout: float = config.APP_TIMEOUT,
        dialog_timeout: float = config.DIALOG_TIMEOUT,
    ) -> None:
        self._executable = executable or self._locate_executable()
        self._app_timeout = app_timeout
        self._dialog_timeout = dialog_timeout
        self._window: WindowSpecification | None = None

    # -- Propiedades --------------------------------------------------------

    @property
    def executable(self) -> Path:
        """Ruta al ejecutable de Excel resuelta al construir el gestor."""
        return self._executable

    @property
    def window(self) -> WindowSpecification:
        """Ventana principal de Excel. Requiere haber llamado a `start()`."""
        if self._window is None:
            raise ExcelNotAvailableError(
                "La aplicacion aun no fue iniciada; invoque start() u open_file()."
            )
        return self._window

    @property
    def title(self) -> str:
        """Titulo actual de la ventana principal (refleja el libro activo)."""
        return self.window.window_text()

    # -- Acciones publicas --------------------------------------------------

    def start(self) -> WindowSpecification:
        """Inicia Excel y espera a que su ventana principal exista.

        Se lanza el proceso y se busca la ventana desde el escritorio en lugar
        de usar `Application.start()`: Excel delega en la instancia ya abierta
        cuando existe, y en ese caso el PID lanzado no es el dueno de la ventana.
        """
        logger.info("Iniciando Excel desde %s", self._executable)
        subprocess.Popen([str(self._executable)])

        window = Desktop(backend="uia").window(class_name=config.EXCEL_WINDOW_CLASS)
        if not window.exists(timeout=self._app_timeout, retry_interval=config.POLL_INTERVAL):
            raise ExcelNotAvailableError(
                f"La ventana '{config.EXCEL_WINDOW_CLASS}' no aparecio en "
                f"{self._app_timeout}s."
            )

        self._window = window
        logger.info("Excel listo. Ventana activa: %r", window.window_text())
        return window

    def open_file(self) -> WindowSpecification:
        """Inicializa Excel y despliega el cuadro de dialogo nativo "Abrir".

        Devuelve la ventana modal recien desplegada para que `FileExplorer`
        tome el control de la interaccion.
        """
        if self._window is None:
            self.start()

        logger.info("Solicitando el cuadro de dialogo 'Abrir' con %s", config.SHORTCUT_OPEN)
        return self._raise_modal_dialog(config.SHORTCUT_OPEN, "Abrir")

    def save_as(self) -> WindowSpecification:
        """Despliega el cuadro de dialogo nativo "Guardar como" (F12).

        Devuelve la ventana modal recien desplegada para que `FileExplorer`
        tome el control de la interaccion.
        """
        logger.info(
            "Solicitando el cuadro de dialogo 'Guardar como' con %s sobre %r",
            config.SHORTCUT_SAVE_AS,
            self.title,
        )
        return self._raise_modal_dialog(config.SHORTCUT_SAVE_AS, "Guardar como")

    def close(self) -> None:
        """Cierra la ventana principal de Excel si sigue disponible."""
        if self._window is None:
            return

        logger.info("Cerrando Excel (%r)", self.title)
        self._window.close()
        try:
            self._window.wait_not(
                "exists", timeout=self._dialog_timeout, retry_interval=config.POLL_INTERVAL
            )
        except WaitTimeoutError:
            logger.warning("La ventana de Excel sigue abierta tras solicitar el cierre.")
        self._window = None

    # -- Internos -----------------------------------------------------------

    def _raise_modal_dialog(self, shortcut: str, description: str) -> WindowSpecification:
        """Envia un atajo global hasta que Excel despliegue una ventana modal.

        El reintento cubre el caso en que Excel todavia estaba inicializando y
        descarto la combinacion de teclas; la condicion de salida es la
        existencia real del dialogo, nunca una pausa arbitraria.
        """
        dialog = self.window.child_window(
            class_name=config.DIALOG_CLASS, control_type="Window"
        )

        for attempt in range(1, config.SHORTCUT_ATTEMPTS + 1):
            self.window.set_focus()
            send_keys(shortcut)

            if dialog.exists(timeout=self._dialog_timeout, retry_interval=config.POLL_INTERVAL):
                title = dialog.window_text()
                logger.info("Cuadro de dialogo desplegado: %r", title)
                return self._narrow_by_title(dialog, title)

            logger.warning(
                "Intento %d/%d: Excel no desplego '%s'; reenviando %s.",
                attempt,
                config.SHORTCUT_ATTEMPTS,
                description,
                shortcut,
            )

        raise DialogNotRaisedError(
            f"Excel no desplego el cuadro de dialogo '{description}' tras "
            f"{config.SHORTCUT_ATTEMPTS} envios de {shortcut}."
        )

    def _narrow_by_title(
        self, dialog: WindowSpecification, title: str
    ) -> WindowSpecification:
        """Reconstruye la especificacion agregando el titulo leido en ejecucion.

        Una especificacion basada solo en `class_name="#32770"` deja de ser
        univoca en cuanto Excel despliega la advertencia de sobreescritura:
        pasan a existir dos ventanas de esa clase bajo la misma ventana padre y
        pywinauto aborta con `ElementAmbiguousError`.

        El titulo no se codifica en el programa, se observa cuando el dialogo
        aparece, asi que el criterio sigue siendo independiente del idioma de
        Office. Se descarta anclar al `handle` porque, al destruirse la ventana,
        `ElementFromHandle` lanza `COMError` en lugar de reportar su ausencia y
        romperia la espera de cierre.
        """
        return self.window.child_window(
            class_name=config.DIALOG_CLASS, control_type="Window", title=title
        )

    @staticmethod
    def _locate_executable() -> Path:
        """Resuelve la ruta de EXCEL.EXE via registro, con respaldo por rutas conocidas."""
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, config.EXCEL_EXECUTABLE_KEY) as key:
                registered = Path(winreg.QueryValue(key, ""))
            if registered.exists():
                logger.debug("EXCEL.EXE resuelto por registro: %s", registered)
                return registered
        except OSError:
            logger.debug("EXCEL.EXE no esta registrado en App Paths; usando respaldos.")

        for candidate in config.EXCEL_FALLBACK_PATHS:
            if candidate.exists():
                return candidate

        raise ExcelNotAvailableError(
            "No se encontro EXCEL.EXE ni en el registro ni en las rutas conocidas."
        )
