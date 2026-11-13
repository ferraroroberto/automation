#!/usr/bin/env python3
"""
illustrations_check.py - Check source folder for matching .afdesign and .png pairs

Scans a folder for *.afdesign and *.png files. Every Affinity file should have
a matching PNG (same base name, different extension). Reports:
- Counts of .afdesign and .png files
- Orphan .afdesign (no matching .png)
- Orphan .png (no matching .afdesign)
- Matched pairs summary

Source folder is read from Illustrations_check.json; CLI argument overrides it.
"""

import argparse
import json
import logging
import sys
from pathlib import Path

log = logging.getLogger(__name__)

CONFIG_FILE = "Illustrations_check.json"


def load_config() -> dict:
    """Load config from Illustrations_check.json (same dir as this script)."""
    config_path = Path(__file__).resolve().parent / CONFIG_FILE
    try:
        if config_path.exists():
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
    except (json.JSONDecodeError, OSError):
        pass
    return {}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Check folder for matching .afdesign and .png file pairs."
    )
    parser.add_argument(
        "source_folder",
        nargs="?",
        help="Folder to scan (default: source_folder from Illustrations_check.json)",
    )
    args = parser.parse_args()

    if args.source_folder:
        source = Path(args.source_folder)
    else:
        config = load_config()
        source_str = (config.get("source_folder") or "").strip()
        if not source_str:
            source_str = input("Enter source folder path: ").strip()
        if not source_str:
            log.error("No folder given. Exiting.")
            return 1
        source = Path(source_str)

    if not source.exists():
        log.error("Error: Folder does not exist: %s", source)
        return 1
    if not source.is_dir():
        log.error("Error: Not a directory: %s", source)
        return 1

    afdesign_files = {f.stem: f for f in source.glob("*.afdesign")}
    png_files = {f.stem: f for f in source.glob("*.png")}

    afdesign_names = set(afdesign_files)
    png_names = set(png_files)

    matched = afdesign_names & png_names
    only_afdesign = afdesign_names - png_names
    only_png = png_names - afdesign_names

    # Spare = any other files in folder that we're not tracking (optional report)
    all_expected = {f.name for f in source.iterdir() if f.is_file()}
    tracked = {f.name for f in afdesign_files.values()} | {f.name for f in png_files.values()}
    spare_files = sorted(all_expected - tracked)

    # --- Report ---
    log.info("=" * 60)
    log.info("ILLUSTRATIONS CHECK")
    log.info("=" * 60)
    log.info("Source folder: %s", source.resolve())
    log.info("  .afdesign files: %d", len(afdesign_files))
    log.info("  .png files:      %d", len(png_files))
    log.info("  Matched pairs:   %d", len(matched))

    if only_afdesign:
        log.info("--- Orphan .afdesign (no matching .png) ---")
        for name in sorted(only_afdesign):
            log.info("  %s", afdesign_files[name].name)

    if only_png:
        log.info("--- Orphan .png (no matching .afdesign) ---")
        for name in sorted(only_png):
            log.info("  %s", png_files[name].name)

    if spare_files:
        log.info("--- Spare files (other than .afdesign / .png) ---")
        for name in spare_files:
            log.info("  %s", name)

    # Summary
    log.info("=" * 60)
    if not only_afdesign and not only_png:
        log.info("OK: Every .afdesign has a matching .png and vice versa.")
    else:
        log.info("Differences found:")
        if only_afdesign:
            log.info("  - %d .afdesign without .png", len(only_afdesign))
        if only_png:
            log.info("  - %d .png without .afdesign", len(only_png))
    if spare_files:
        log.info("  - %d other file(s) in folder", len(spare_files))
    log.info("=" * 60)

    return 0


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    sys.exit(main())
