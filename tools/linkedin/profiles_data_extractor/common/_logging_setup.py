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
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[4]))  # repo root, so the shared helper imports when run as a script
from tools.system._lib import runtime_files  # noqa: E402

# Tests point this at a temp file so importing the package never writes into the tree.
LOG_FILE_ENV = "LINKEDIN_EXTRACTOR_LOG_FILE"


def setup_logging(name: str) -> logging.Logger:
    """Configure a logger that writes INFO+ to the package's shared log in the
    runtime folder (``runtime_files.runtime_data_dir()``) and WARNING+ to the console.
    """
    override = os.environ.get(LOG_FILE_ENV)
    log_file = Path(override) if override else runtime_files.runtime_data_dir() / "linkedin_profiles_data_extractor.log"
    file_handler = runtime_files.file_handler(log_file, encoding='utf-8')
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
