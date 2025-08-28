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

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
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
    logger.warning(f"⚠️  pillow-heif not available: {e}. HEIC conversion may not work.")
    pillow_heif_available = False
else:
    pillow_heif_available = True

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
            logger.error(f"❌ Cannot open HEIC file '{file_path}': pillow-heif not available. Install with: pip install pillow-heif")
        else:
            # pillow-heif is available but file is still unreadable (corrupted, wrong format, etc.)
            logger.error(f"❌ Cannot identify HEIC file '{file_path}': {str(e)}")
        return False
    except OSError as e:
        # File system errors: permissions, disk space, file locked, etc.
        logger.error(f"❌ OS error reading HEIC file '{file_path}': {str(e)}")
        return False
    except Exception as e:
        # Catch any other unexpected errors during validation
        logger.error(f"❌ Unexpected error with HEIC file '{file_path}': {str(e)}")
        return False

def convert_heic_to_jpg(folder_path: str) -> None:
    """Convert all HEIC/HEIF images in the folder to JPG format"""
    if not folder_path:
        print("No folder selected.")
        return

    folder = Path(folder_path)
    if not folder.exists():
        print(f"❌ Folder does not exist: {folder_path}")
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
        print("❌ No HEIC/HEIF files found in the selected folder.")
        return

    converted_count = 0
    error_count = 0

    print(f"📂 Found {len(heic_files)} HEIC/HEIF files in {folder_path}")
    print("🔄 Starting conversion...")

    for heic_file in heic_files:
        # Validate the HEIC file first
        if not validate_heic_file(heic_file):
            print(f"⚠️  Skipping {heic_file.name} - invalid or corrupted HEIC file")
            error_count += 1
            continue

        # Create JPG filename with same name but .jpg extension
        jpg_file = heic_file.with_suffix('.jpg')

        # Check if JPG already exists
        if jpg_file.exists():
            print(f"⚠️  JPG already exists for {heic_file.name}, skipping")
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

                print(f"✅ Converted {heic_file.name} -> {jpg_file.name}")
                converted_count += 1

        except UnidentifiedImageError as e:
            if not pillow_heif_available:
                print(f"❌ Cannot convert '{heic_file.name}': pillow-heif not available. Install with: pip install pillow-heif")
            else:
                print(f"❌ Cannot identify image file '{heic_file}': {str(e)}")
            error_count += 1
        except OSError as e:
            print(f"❌ OS error processing '{heic_file}': {str(e)}")
            error_count += 1
        except Exception as e:
            print(f"❌ Error converting {heic_file.name}: {str(e)}")
            error_count += 1

    print(f"\n🎉 Conversion complete!")
    print(f"✅ Converted: {converted_count} files")
    print(f"❌ Errors: {error_count} files")

    if error_count > 0:
        print(f"\n📝 Note: {error_count} files had errors and were skipped.")
        print("This could be due to corrupted files, unsupported formats, or permission issues.")

def main() -> None:
    """Main function to run the HEIC to JPG converter"""
    print("📸 HEIC to JPG Converter")
    print("=" * 50)
    print("Converts HEIC/HEIF images to JPG while preserving size and quality")
    print()

    # Check HEIC support before proceeding
    # This prevents user confusion by warning about missing dependencies upfront
    if not pillow_heif_available:
        print("⚠️  Warning: pillow-heif library is not available.")
        print("   HEIC conversion may fail. Please install pillow-heif:")
        print("   pip install pillow-heif")
        print("   Note: Continue anyway to see specific error messages per file")
        print()

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
        print("❌ No folder selected. Exiting.")

if __name__ == "__main__":
    main()
