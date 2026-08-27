"""Rutas, tiempos de espera y selectores de UI Automation del proyecto.

Todos los selectores fueron obtenidos inspeccionando el arbol real de UI
Automation de Excel (Office 16) y del explorador de archivos de Windows 11.
Se privilegian los `AutomationId`, que son invariables al idioma de Office;
los nombres localizados se usan solo como respaldo.
"""

from pathlib import Path

# --- Rutas -----------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_ROOT / ".data"
DEFAULT_SOURCE_FILE = DATA_DIR / "input" / "origen.xlsx"
DEFAULT_TARGET_FILE = DATA_DIR / "output" / "destino.xlsx"

# --- Tiempos de espera (sincronizacion por eventos, nunca time.sleep) -------

APP_TIMEOUT = 60.0
"""Segundos maximos para que la ventana principal de Excel exista."""

DIALOG_TIMEOUT = 20.0
"""Segundos maximos para que un cuadro de dialogo nativo aparezca o desaparezca."""

CONFIRM_TIMEOUT = 5.0
"""Segundos maximos para detectar la advertencia de sobreescritura."""

POLL_INTERVAL = 0.25
"""Intervalo de sondeo que usan los metodos nativos `Exists`/`Disappears`."""

CONTROL_PROBE_TIMEOUT = 1.0
"""Sondeo breve para discriminar entre selectores alternativos de un mismo control.

El dialogo ya fue confirmado como existente, asi que sus controles estan
construidos: una primera pasada corta descarta rapido el selector que no aplica
y solo si ninguno responde se repite la busqueda con el tiempo completo.
"""

SHORTCUT_ATTEMPTS = 3
"""Reintentos del atajo global si Excel aun no estaba listo para recibirlo."""

# --- Selectores de la aplicacion Excel -------------------------------------

EXCEL_WINDOW_CLASS = "XLMAIN"
EXCEL_EXECUTABLE_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
EXCEL_FALLBACK_PATHS = (
    Path(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"),
)

SHORTCUT_OPEN = "{Ctrl}{F12}"
"""Atajo global que despliega el cuadro de dialogo nativo "Abrir"."""

SHORTCUT_SAVE_AS = "{F12}"
"""Atajo global que despliega el cuadro de dialogo nativo "Guardar como"."""

# --- Selectores de los dialogos nativos de Windows -------------------------

DIALOG_CLASS = "#32770"
"""Clase de ventana de los dialogos comunes de Windows (Abrir / Guardar como)."""

FILE_NAME_EDIT_IDS = ("1148", "1001")
"""`AutomationId` del campo "Nombre de archivo": 1148 en Abrir, 1001 en Guardar como."""

ACCEPT_BUTTON_ID = "1"
"""`AutomationId` del boton de confirmacion (Abrir / Guardar)."""

DIALOG_SEARCH_DEPTH = 8
"""Profundidad maxima al buscar controles dentro de un dialogo."""

OVERWRITE_BUTTON_CLASS = "CCPushButton"
"""Clase de los botones del TaskDialog de confirmacion de sobreescritura."""

OVERWRITE_YES_NAMES = ("Sí", "Si", "Yes")
"""Nombres localizados del boton afirmativo, usados como respaldo del AutomationId."""

OVERWRITE_YES_ID = "CommandButton_6"
"""`AutomationId` del boton "Si" en el TaskDialog "Confirmar Guardar como"."""
