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
import sys
from pathlib import Path

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
            print("No folder given. Exiting.")
            return 1
        source = Path(source_str)

    if not source.exists():
        print(f"Error: Folder does not exist: {source}")
        return 1
    if not source.is_dir():
        print(f"Error: Not a directory: {source}")
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
    print()
    print("=" * 60)
    print("ILLUSTRATIONS CHECK")
    print("=" * 60)
    print(f"Source folder: {source.resolve()}")
    print()
    print(f"  .afdesign files: {len(afdesign_files)}")
    print(f"  .png files:      {len(png_files)}")
    print(f"  Matched pairs:   {len(matched)}")
    print()

    if only_afdesign:
        print("--- Orphan .afdesign (no matching .png) ---")
        for name in sorted(only_afdesign):
            print(f"  {afdesign_files[name].name}")
        print()

    if only_png:
        print("--- Orphan .png (no matching .afdesign) ---")
        for name in sorted(only_png):
            print(f"  {png_files[name].name}")
        print()

    if spare_files:
        print("--- Spare files (other than .afdesign / .png) ---")
        for name in spare_files:
            print(f"  {name}")
        print()

    # Summary
    print("=" * 60)
    if not only_afdesign and not only_png:
        print("OK: Every .afdesign has a matching .png and vice versa.")
    else:
        print("Differences found:")
        if only_afdesign:
            print(f"  - {len(only_afdesign)} .afdesign without .png")
        if only_png:
            print(f"  - {len(only_png)} .png without .afdesign")
    if spare_files:
        print(f"  - {len(spare_files)} other file(s) in folder")
    print("=" * 60)
    print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
