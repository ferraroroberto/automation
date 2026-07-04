#!/usr/bin/env python
"""
Shared folder-batch image -> JPG conversion scaffolding.

Consolidates the "folder-pick -> collect files by extension -> per-file
validate -> PIL load -> RGB convert -> save JPG -> count/skip -> completion
messagebox" shape duplicated between heic_to_jpg_converter.py and
png_to_jpg_converter.py (dedup: audit issue #64).

The two callers differ meaningfully in how they make an image JPEG-safe
(HEIC: naive RGB convert; PNG: alpha-composite onto a white background),
whether they protect an existing same-named .jpg from being overwritten
(HEIC: yes; PNG: no - preserved as-is, not fixed here), and their JPEG
save quality kwargs - so only the genuinely identical scaffolding lives
here; those differences stay as caller-supplied parameters.
"""

import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Callable, Dict, Iterable

from PIL import Image, UnidentifiedImageError


def select_folder(dialog_title: str) -> str:
    """Open a folder-selection dialog and return the chosen path (or '' if cancelled)."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory(title=dialog_title)
    root.destroy()
    return folder_path


def show_completion_messagebox(title: str, message: str) -> None:
    """Show a simple info messagebox (used for the "conversion complete" popup)."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def convert_folder_to_jpg(
    folder_path: str,
    source_extensions: Iterable[str],
    validate_fn: Callable[[Path], bool],
    convert_to_rgb_fn: Callable[["Image.Image"], "Image.Image"],
    save_kwargs: Dict,
    *,
    skip_existing_jpg: bool,
    logger: logging.Logger,
    kind_label: str = "image",
) -> None:
    """
    Convert every file under `folder_path` matching `source_extensions` to JPG.

    Args:
        source_extensions: file extensions to scan for (e.g. ['.heic', '.heif']).
        validate_fn:       fn(path) -> bool, run before conversion (integrity check).
        convert_to_rgb_fn: fn(img) -> img, applied after img.load() to make the
                            image JPEG-savable (handles transparency/mode per caller).
        save_kwargs:       passed straight to Image.save(path, 'JPEG', **save_kwargs).
        skip_existing_jpg: if True, an existing same-named .jpg is left untouched
                            (counted as an error/skip); if False, it's overwritten.
        kind_label:        short label used in log messages (e.g. "HEIC/HEIF", "PNG").
    """
    if not folder_path:
        logger.info("No folder selected.")
        return

    folder = Path(folder_path)
    if not folder.exists():
        logger.error("❌ Folder does not exist: %s", folder_path)
        return

    exts = {e.lower() for e in source_extensions}
    files = [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in exts]

    if not files:
        logger.info("❌ No %s files found in the selected folder.", kind_label)
        return

    converted_count = 0
    error_count = 0

    logger.info("📂 Found %d %s files in %s", len(files), kind_label, folder_path)
    logger.info("🔄 Starting conversion...")

    for src_file in files:
        if not validate_fn(src_file):
            logger.warning("⚠️  Skipping %s - invalid or corrupted %s file", src_file.name, kind_label)
            error_count += 1
            continue

        jpg_file = src_file.with_suffix('.jpg')

        if skip_existing_jpg and jpg_file.exists():
            logger.warning("⚠️  JPG already exists for %s, skipping", src_file.name)
            error_count += 1
            continue

        try:
            with Image.open(src_file) as img:
                img.load()
                img = convert_to_rgb_fn(img)
                img.save(jpg_file, 'JPEG', **save_kwargs)
                logger.info("✅ Converted %s -> %s", src_file.name, jpg_file.name)
                converted_count += 1
        except UnidentifiedImageError as e:
            logger.error("❌ Cannot identify image file '%s': %s", src_file, e)
            error_count += 1
        except OSError as e:
            logger.error("❌ OS error processing '%s': %s", src_file, e)
            error_count += 1
        except Exception as e:
            logger.error("❌ Error converting %s: %s", src_file.name, e)
            error_count += 1

    logger.info("🎉 Conversion complete!")
    logger.info("✅ Converted: %d files", converted_count)
    logger.info("❌ Errors: %d files", error_count)

    if error_count > 0:
        logger.info("📝 Note: %d files had errors and were skipped.", error_count)
        logger.info("This could be due to corrupted files, unsupported formats, or permission issues.")
