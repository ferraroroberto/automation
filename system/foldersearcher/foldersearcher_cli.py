#!/usr/bin/env python3
"""
Folder Searcher headless CLI - rebuild the index without the tray app.

Usage (from the repo root):
    & .\\.venv\\Scripts\\python.exe system\\foldersearcher\\foldersearcher_cli.py scan

Built for unattended runs (the nightly app-launcher chain hop,
``run-scan-nightly.bat``). It deliberately differs from the tray's
**Scan All Roots** in two ways:

- A missing root aborts the run and leaves the existing index untouched.
  The core only warns and skips a missing root, which on an unattended run
  would silently shrink the index.
- The index is written atomically (temp file in the same directory, then
  ``os.replace``), so the tray app never reads a half-written file.

It imports no ``tkinter``/``pystray``, takes no single-instance mutex, and
never writes ``foldersearcher.json`` back, so it runs fine while the tray app
is open.

Exit codes:
    0  every configured root existed and the index was replaced
    1  the index could not be written, or an unexpected crash
    2  usage error (argparse)
    3  no roots configured (index left untouched)
    4  one or more configured roots are missing (index left untouched)
"""

import argparse
import logging
import os
import sys
import tempfile
import time
from typing import List, Optional

from foldersearcher_core import FolderIndex, load_config

logger = logging.getLogger(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG_FILE = os.path.join(SCRIPT_DIR, "foldersearcher.json")
DEFAULT_STRUCTURE_FILE = os.path.join(SCRIPT_DIR, "folder_structure.txt")

EXIT_OK = 0
EXIT_WRITE_FAILED = 1
EXIT_NO_ROOTS = 3
EXIT_MISSING_ROOT = 4


def write_index_atomically(index: FolderIndex, structure_file: str) -> None:
    """Save ``index`` to a sibling temp file, then swap it onto ``structure_file``.

    The temp file lives in the target's directory so ``os.replace`` is a
    same-volume rename. On any failure the temp file is removed and the
    existing index is left as it was.
    """
    directory = os.path.dirname(os.path.abspath(structure_file))
    fd, temp_path = tempfile.mkstemp(dir=directory, prefix=".folder_structure.", suffix=".tmp")
    os.close(fd)
    try:
        index.save(temp_path)
        os.replace(temp_path, structure_file)
    except BaseException:
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


def run_scan(config_file: str, structure_file: str) -> int:
    """Scan every configured root and replace the index. Returns an exit code."""
    config = load_config(config_file, structure_file)
    roots = config.root_paths

    if not roots:
        logger.error("❌ No roots configured in %s; index left untouched", config_file)
        return EXIT_NO_ROOTS

    logger.info("ℹ️ Roots (%d): %s", len(roots), ", ".join(roots))
    missing = [root for root in roots if not os.path.isdir(root)]
    if missing:
        for root in missing:
            logger.error("❌ Configured root is missing: %s", root)
        logger.error("❌ %d of %d root(s) missing; index left untouched: %s",
                     len(missing), len(roots), config.structure_file)
        return EXIT_MISSING_ROOT

    started = time.monotonic()
    index = FolderIndex()
    count = index.scan(roots)
    try:
        write_index_atomically(index, config.structure_file)
    except OSError as exc:
        logger.error("❌ Could not write index %s: %s", config.structure_file, exc)
        return EXIT_WRITE_FAILED
    duration = time.monotonic() - started

    logger.info("✅ Index replaced: %d folders across %d root(s) in %.1fs -> %s",
                count, len(roots), duration, config.structure_file)
    return EXIT_OK


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Folder Searcher headless index tools.")
    subcommands = parser.add_subparsers(dest="command", required=True)
    subcommands.add_parser("scan", help="Rebuild folder_structure.txt from foldersearcher.json's roots.")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    """Entry point. Returns the process exit code."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            reconfigure(encoding="utf-8")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    # "scan" is the only subcommand; argparse exits 2 on anything else.
    build_parser().parse_args(argv)
    return run_scan(DEFAULT_CONFIG_FILE, DEFAULT_STRUCTURE_FILE)


if __name__ == "__main__":
    sys.exit(main())
