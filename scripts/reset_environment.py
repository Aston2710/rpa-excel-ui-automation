"""Restablece el proyecto al estado inicial de los casos de prueba.

Elimina unicamente lo que generan las ejecuciones del bot y las herramientas de
desarrollo. El insumo `.data/input/origen.xlsx` esta versionado y nunca se toca;
el script verifica que siga presente porque sin el no se puede ejecutar el
Caso 01.

Importa las rutas desde `rpa_excel_ui_automation.config` para que no puedan
divergir de las que usa el bot.

Uso:
    pdm run reset               # muestra y ejecuta la limpieza
    pdm run reset --dry-run     # solo muestra que borraria
    pdm run reset --kill-excel  # cierra Excel a la fuerza antes de borrar
"""

from __future__ import annotations

import argparse
import csv
import ctypes
import logging
import shutil
import subprocess
import sys
from ctypes import wintypes
from pathlib import Path

from rpa_excel_ui_automation import config

logger = logging.getLogger("reset")

EXCEL_IMAGE = "EXCEL.EXE"

SYNCHRONIZE = 0x00100000
WAIT_TIMEOUT = 0x00000102
EXIT_WAIT_SECONDS = 15.0

GENERATED_FILES = ("@AutomationLog.txt",)
"""Archivos sueltos que dejan las herramientas en la raiz del proyecto."""

GENERATED_DIRS = (".pytest_cache",)
"""Directorios de cache en la raiz del proyecto."""

GENERATED_PATTERNS = ("src/**/__pycache__", "tests/**/__pycache__", "src/*.egg-info")
"""Patrones acotados a `src/` y `tests/` para no recorrer nunca `.venv`.

Nunca se eliminan `.venv/`, `.pdm-python` ni `pdm.lock`: reconstruirlos exige
volver a resolver el entorno.
"""


class ResetError(RuntimeError):
    """La limpieza no pudo completarse."""


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Enumera lo que se borraria sin modificar nada.",
    )
    parser.add_argument(
        "--kill-excel",
        action="store_true",
        help=(
            "Cierra Excel a la fuerza antes de borrar. Sin esta bandera el "
            "script solo advierte, porque matar Excel descarta cambios sin "
            "guardar de cualquier libro abierto."
        ),
    )
    parser.add_argument("--verbose", action="store_true", help="Eleva el log a DEBUG.")
    return parser.parse_args(argv)


def inside_project(path: Path) -> bool:
    """Salvaguarda: nada se borra fuera de la raiz del proyecto."""
    return config.PROJECT_ROOT in path.resolve().parents


def excel_pids() -> list[int]:
    """Identificadores de los procesos de Excel en ejecucion."""
    result = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {EXCEL_IMAGE}", "/NH", "/FO", "CSV"],
        capture_output=True,
        text=True,
        check=False,
    )
    pids = []
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) >= 2 and row[0].lower() == EXCEL_IMAGE.lower():
            pids.append(int(row[1]))
    return pids


def _wait_for_exit(pids: list[int], timeout_seconds: float) -> bool:
    """Espera a que los procesos indicados terminen realmente.

    `taskkill` regresa antes de que el proceso muera, y mientras muere Windows
    conserva el bloqueo sobre los archivos abiertos: borrar de inmediato falla
    con `WinError 32`. Se espera sobre el handle de cada proceso, que es la
    senal real de terminacion, en lugar de intercalar pausas arbitrarias.
    """
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)

    timeout_ms = int(timeout_seconds * 1000)
    for pid in pids:
        handle = kernel32.OpenProcess(SYNCHRONIZE, False, pid)
        if not handle:
            continue  # El proceso ya termino.
        try:
            if kernel32.WaitForSingleObject(handle, timeout_ms) == WAIT_TIMEOUT:
                return False
        finally:
            kernel32.CloseHandle(handle)
    return True


def kill_excel() -> None:
    pids = excel_pids()
    if not pids:
        return

    logger.warning(
        "Cerrando %d proceso(s) %s a la fuerza; se descartan cambios sin guardar.",
        len(pids),
        EXCEL_IMAGE,
    )
    subprocess.run(["taskkill", "/F", "/IM", EXCEL_IMAGE], capture_output=True, check=False)

    if not _wait_for_exit(pids, EXIT_WAIT_SECONDS):
        raise ResetError(
            f"{EXCEL_IMAGE} no termino en {EXIT_WAIT_SECONDS}s; el destino seguiria bloqueado."
        )
    logger.info("%s termino; los archivos quedaron liberados.", EXCEL_IMAGE)


def remove(path: Path, dry_run: bool) -> bool:
    """Elimina un archivo o directorio. Devuelve True si habia algo que borrar."""
    if not path.exists():
        return False

    if not inside_project(path):
        raise ResetError(f"Ruta fuera del proyecto, no se elimina: {path}")

    kind = "directorio" if path.is_dir() else "archivo"
    if dry_run:
        logger.info("[dry-run] se borraria el %s %s", kind, path.relative_to(config.PROJECT_ROOT))
        return True

    try:
        if path.is_dir():
            shutil.rmtree(path)
        else:
            path.unlink()
    except OSError as error:
        raise ResetError(
            f"No se pudo borrar {path.relative_to(config.PROJECT_ROOT)}: {error}. "
            "Si es un .xlsx, Excel probablemente lo tiene bloqueado: cierre Excel "
            "o use --kill-excel."
        ) from error

    logger.info("Borrado el %s %s", kind, path.relative_to(config.PROJECT_ROOT))
    return True


def clean_output_dir(dry_run: bool) -> int:
    """Vacia `.data/output/` conservando el directorio."""
    output_dir = config.DEFAULT_TARGET_FILE.parent
    if not output_dir.is_dir():
        logger.info(
            "El directorio %s no existe; el bot lo crea al guardar.",
            output_dir.relative_to(config.PROJECT_ROOT),
        )
        return 0

    return sum(remove(entry, dry_run) for entry in sorted(output_dir.iterdir()))


def clean_run_artifacts(dry_run: bool) -> int:
    """Elimina los rastros que deja una ejecucion del bot fuera de `.data/`."""
    return sum(remove(config.PROJECT_ROOT / name, dry_run) for name in GENERATED_FILES)


def clean_caches(dry_run: bool) -> int:
    """Elimina cache de herramientas.

    Se cuenta por separado porque el propio script regenera `__pycache__` al
    importar el paquete: incluirlo en el total haria que nunca se pudiera
    informar que el proyecto ya estaba en su estado inicial.
    """
    removed = 0
    for name in GENERATED_DIRS:
        removed += remove(config.PROJECT_ROOT / name, dry_run)
    for pattern in GENERATED_PATTERNS:
        for path in sorted(config.PROJECT_ROOT.glob(pattern)):
            removed += remove(path, dry_run)
    return removed


def verify_fixture() -> None:
    """El insumo del Caso 01 debe seguir presente tras la limpieza."""
    source = config.DEFAULT_SOURCE_FILE
    if not source.is_file():
        raise ResetError(
            f"Falta el insumo versionado {source.relative_to(config.PROJECT_ROOT)}. "
            "Recupere con: git checkout -- .data/input/origen.xlsx"
        )
    logger.info(
        "Insumo intacto: %s (%d bytes)",
        source.relative_to(config.PROJECT_ROOT),
        source.stat().st_size,
    )


def run(dry_run: bool, kill: bool) -> int:
    if excel_pids():
        if kill:
            kill_excel()
        else:
            logger.warning(
                "%s esta en ejecucion. Si mantiene abierto el destino, el borrado "
                "fallara. Cierre Excel o repita con --kill-excel.",
                EXCEL_IMAGE,
            )

    artifacts = clean_output_dir(dry_run) + clean_run_artifacts(dry_run)
    caches = clean_caches(dry_run)
    verify_fixture()

    verb = "se borrarian" if dry_run else "eliminados"
    logger.info("Salidas de ejecucion: %d %s. Cache: %d %s.", artifacts, verb, caches, verb)

    if dry_run:
        logger.info("[dry-run] No se modifico nada.")
    elif artifacts == 0:
        logger.info("El proyecto ya estaba en su estado inicial.")
    else:
        logger.info(
            "Estado inicial restablecido. La proxima ejecucion de `pdm run bot` "
            "probara la creacion del destino; la siguiente, el reemplazo."
        )
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s | %(message)s",
    )
    try:
        return run(args.dry_run, args.kill_excel)
    except ResetError as error:
        logger.error("%s", error)
        return 1


if __name__ == "__main__":
    sys.exit(main())
