"""Ejecutor de los casos de prueba 01 y 02 del plan de pruebas."""

from __future__ import annotations

import argparse
import hashlib
import logging
import sys
from pathlib import Path

from . import config
from .excel_manager import ExcelManager
from .file_explorer import FileExplorer

logger = logging.getLogger("rpa_excel_ui_automation.runner")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--source",
        type=Path,
        default=config.DEFAULT_SOURCE_FILE,
        help="Archivo de origen a abrir (por defecto .data/input/origen.xlsx).",
    )
    parser.add_argument(
        "--target",
        type=Path,
        default=config.DEFAULT_TARGET_FILE,
        help="Archivo destino a exportar (por defecto .data/output/destino.xlsx).",
    )
    parser.add_argument(
        "--close",
        action="store_true",
        help="Cierra Excel al finalizar en lugar de dejar la ventana activa.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Eleva el nivel de log a DEBUG.",
    )
    return parser.parse_args(argv)


def digest(path: Path) -> str:
    """Huella SHA-256 del archivo, usada para probar que el origen no se altera."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(source: Path, target: Path, close: bool) -> int:
    manager = ExcelManager()
    explorer = FileExplorer()

    source = source.resolve()
    target = target.resolve()
    source_digest_before = digest(source)

    logger.info("=== Caso 01: inicializacion e importacion dinamica ===")
    open_dialog = manager.open_file()
    explorer.open_document(open_dialog, source)

    if source.name.lower() not in manager.title.lower():
        logger.error("El titulo %r no corresponde a %s", manager.title, source.name)
        return 1
    logger.info("Caso 01 superado. Ventana activa: %r", manager.title)

    logger.info("=== Caso 02: exportacion segura con reemplazo ===")
    save_dialog = manager.save_as()
    explorer.save_document(save_dialog, target)

    if not target.is_file():
        logger.error("El archivo destino no se creo: %s", target)
        return 1
    if digest(source) != source_digest_before:
        logger.error("El archivo de origen fue alterado: %s", source)
        return 1
    logger.info("Caso 02 superado. Destino: %s (%d bytes)", target, target.stat().st_size)
    logger.info("Origen intacto (SHA-256 sin cambios): %s", source)

    if close:
        manager.close()
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    )
    return run(args.source, args.target, args.close)


if __name__ == "__main__":
    sys.exit(main())
