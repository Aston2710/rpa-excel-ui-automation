"""Automatizacion RPA de la interfaz de usuario de Microsoft Excel."""

from .excel_manager import DialogNotRaisedError, ExcelManager, ExcelNotAvailableError
from .file_explorer import DialogControlNotFoundError, DialogStillOpenError, FileExplorer

__all__ = [
    "DialogControlNotFoundError",
    "DialogNotRaisedError",
    "DialogStillOpenError",
    "ExcelManager",
    "ExcelNotAvailableError",
    "FileExplorer",
]
