#!/usr/bin/env python
"""
HEIC to JPG Converter

Converts HEIC/HEIF images to JPG format while preserving size and quality.
Saves converted files with the same name in the same folder.

Features:
- High-quality conversion (95% JPG quality)
- Preserves original image dimensions
- Handles transparency and various color modes
- Comprehensive error handling and user feedback
- Supports both .heic and .heif extensions
- Prevents overwriting existing JPG files

Requirements:
- PIL/Pillow (for image processing)
- pillow-heif (for HEIC format support)
- tkinter (for GUI folder selection)

Usage:
    python heic_to_jpg_converter.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, UnidentifiedImageError
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

# HEIC Support Registration
# HEIC format requires pillow-heif library to extend PIL's capabilities
# This registers the HEIF opener with PIL to handle .heic/.heif files
try:
    import pillow_heif
    pillow_heif.register_heif_opener()  # Register HEIC/HEIF format with PIL
    logger.info("✅ HEIC support registered successfully")
    pillow_heif_available = True
except ImportError as e:
    # pillow-heif not installed - HEIC files cannot be opened
    logger.warning("⚠️  pillow-heif not available: %s. HEIC conversion may not work.", e)
    pillow_heif_available = False

def select_folder() -> str:
    """Open folder dialog and return selected folder path"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    folder_path = filedialog.askdirectory(
        title="Select folder containing HEIC images"
    )

    root.destroy()
    return folder_path

def validate_heic_file(file_path: Path) -> bool:
    """
    Validate if a file is a valid HEIC/HEIF image that can be opened.

    Performs integrity check by attempting to open and verify the image.
    This catches corrupted files early before attempting conversion.
    """
    try:
        with Image.open(file_path) as img:
            # img.verify() checks file integrity without fully loading into memory
            # This is crucial for HEIC files which may have complex compression
            img.verify()
            return True
    except UnidentifiedImageError as e:
        # Most common error: PIL cannot identify the image format
        if not pillow_heif_available:
            # Specific error for missing HEIC support library
            logger.error("❌ Cannot open HEIC file '%s': pillow-heif not available. Install with: pip install pillow-heif", file_path)
        else:
            # pillow-heif is available but file is still unreadable (corrupted, wrong format, etc.)
            logger.error("❌ Cannot identify HEIC file '%s': %s", file_path, e)
        return False
    except OSError as e:
        # File system errors: permissions, disk space, file locked, etc.
        logger.error("❌ OS error reading HEIC file '%s': %s", file_path, e)
        return False
    except Exception as e:
        # Catch any other unexpected errors during validation
        logger.error("❌ Unexpected error with HEIC file '%s': %s", file_path, e)
        return False

def convert_heic_to_jpg(folder_path: str) -> None:
    """Convert all HEIC/HEIF images in the folder to JPG format"""
    if not folder_path:
        logger.info("No folder selected.")
        return

    folder = Path(folder_path)
    if not folder.exists():
        logger.error("❌ Folder does not exist: %s", folder_path)
        return

    # Find all HEIC/HEIF files in the folder (case insensitive)
    # HEIC (High Efficiency Image Container) and HEIF (High Efficiency Image Format)
    # are Apple's efficient image formats used by iPhones and other devices
    heic_files = []
    heic_extensions = ['.heic', '.heif']  # Both extensions are supported

    for file in folder.iterdir():
        # Check file extension case-insensitively to catch .HEIC, .Heic, etc.
        if file.is_file() and file.suffix.lower() in heic_extensions:
            heic_files.append(file)

    if not heic_files:
        logger.info("❌ No HEIC/HEIF files found in the selected folder.")
        return

    converted_count = 0
    error_count = 0

    logger.info("📂 Found %d HEIC/HEIF files in %s", len(heic_files), folder_path)
    logger.info("🔄 Starting conversion...")

    for heic_file in heic_files:
        # Validate the HEIC file first
        if not validate_heic_file(heic_file):
            logger.warning("⚠️  Skipping %s - invalid or corrupted HEIC file", heic_file.name)
            error_count += 1
            continue

        # Create JPG filename with same name but .jpg extension
        jpg_file = heic_file.with_suffix('.jpg')

        # Check if JPG already exists
        if jpg_file.exists():
            logger.warning("⚠️  JPG already exists for %s, skipping", heic_file.name)
            error_count += 1
            continue

        try:
            # Open HEIC image using PIL with pillow-heif extension
            with Image.open(heic_file) as img:
                # img.load() forces loading of image data into memory
                # This is necessary for HEIC files which may use lazy loading
                img.load()

                # Convert image mode to RGB for JPG compatibility
                # HEIC files can be in various modes (RGBA, CMYK, etc.)
                # JPG format requires RGB or grayscale (L) mode
                if img.mode not in ('RGB', 'L'):
                    # Convert to RGB to ensure compatibility with JPG format
                    # This handles transparency (RGBA) and other color modes
                    img = img.convert('RGB')

                # Save as JPG with optimized settings for quality preservation
                # quality=95: High quality (95%) - excellent visual quality with reasonable file size
                # optimize=True: Enable optimization to reduce file size without quality loss
                # progressive=True: Progressive JPG for better web loading (loads blurry to sharp)
                img.save(jpg_file, 'JPEG', quality=95, optimize=True, progressive=True)

                logger.info("✅ Converted %s -> %s", heic_file.name, jpg_file.name)
                converted_count += 1

        except UnidentifiedImageError as e:
            if not pillow_heif_available:
                logger.error("❌ Cannot convert '%s': pillow-heif not available. Install with: pip install pillow-heif", heic_file.name)
            else:
                logger.error("❌ Cannot identify image file '%s': %s", heic_file, e)
            error_count += 1
        except OSError as e:
            logger.error("❌ OS error processing '%s': %s", heic_file, e)
            error_count += 1
        except Exception as e:
            logger.error("❌ Error converting %s: %s", heic_file.name, e)
            error_count += 1

    logger.info("🎉 Conversion complete!")
    logger.info("✅ Converted: %d files", converted_count)
    logger.info("❌ Errors: %d files", error_count)

    if error_count > 0:
        logger.info("📝 Note: %d files had errors and were skipped.", error_count)
        logger.info("This could be due to corrupted files, unsupported formats, or permission issues.")

def main() -> None:
    """Main function to run the HEIC to JPG converter"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger.info("📸 HEIC to JPG Converter")
    logger.info("=" * 50)
    logger.info("Converts HEIC/HEIF images to JPG while preserving size and quality")

    # Check HEIC support before proceeding
    # This prevents user confusion by warning about missing dependencies upfront
    if not pillow_heif_available:
        logger.warning("⚠️  Warning: pillow-heif library is not available.")
        logger.warning("   HEIC conversion may fail. Please install pillow-heif:")
        logger.warning("   pip install pillow-heif")
        logger.warning("   Note: Continue anyway to see specific error messages per file")

    # Select folder
    folder_path = select_folder()

    if folder_path:
        # Convert HEIC files to JPG
        convert_heic_to_jpg(folder_path)

        # Show completion message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Conversion Complete",
            f"HEIC to JPG conversion completed!\n\nCheck the selected folder for converted files."
        )
        root.destroy()
    else:
        logger.info("❌ No folder selected. Exiting.")

if __name__ == "__main__":
    main()
