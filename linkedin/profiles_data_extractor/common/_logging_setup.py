"""
Shared logging bootstrap for the LinkedIn profiles data extractor package.

``excel_format_manager`` and ``history_manager`` both import ``setup_logging``
from here so the file+console handler wiring lives in exactly one place.
Kept dependency-free (no imports of sibling modules) so it can be imported by
``excel_format_manager`` without creating an import cycle with ``_lib``
(which itself imports from ``excel_format_manager``).
"""

import logging
import os
from pathlib import Path

# Tests point this at a temp file so importing the package never writes into the tree.
LOG_FILE_ENV = "LINKEDIN_EXTRACTOR_LOG_FILE"


def setup_logging(name: str) -> logging.Logger:
    """Configure a logger that writes INFO+ to the package's shared logging.log
    and WARNING+ to the console.
    """
    override = os.environ.get(LOG_FILE_ENV)
    log_file = Path(override) if override else Path(__file__).parent.parent / "logging.log"
    # runtime-write-exempt: the shared package log; tests redirect it to a temp file via LOG_FILE_ENV
    log_file.parent.mkdir(parents=True, exist_ok=True)

    # runtime-write-exempt: the shared package log; tests redirect it to a temp file via LOG_FILE_ENV
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    console_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger
