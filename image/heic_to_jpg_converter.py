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

from PIL import Image, UnidentifiedImageError
from pathlib import Path
import logging

from _image_to_jpg_converter import select_folder, convert_folder_to_jpg, show_completion_messagebox

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

def _heic_to_rgb(img: "Image.Image") -> "Image.Image":
    """
    Convert image mode to RGB for JPG compatibility.
    HEIC files can be in various modes (RGBA, CMYK, etc.); JPG format
    requires RGB or grayscale (L) mode.
    """
    if img.mode not in ('RGB', 'L'):
        img = img.convert('RGB')
    return img


def convert_heic_to_jpg(folder_path: str) -> None:
    """Convert all HEIC/HEIF images in the folder to JPG format"""
    convert_folder_to_jpg(
        folder_path,
        source_extensions=['.heic', '.heif'],  # HEIC and HEIF: Apple's efficient image formats
        validate_fn=validate_heic_file,
        convert_to_rgb_fn=_heic_to_rgb,
        # quality=95: excellent visual quality with reasonable file size
        # optimize=True: reduce file size without quality loss
        # progressive=True: progressive JPG for better web loading (loads blurry to sharp)
        save_kwargs={'quality': 95, 'optimize': True, 'progressive': True},
        skip_existing_jpg=True,
        logger=logger,
        kind_label='HEIC/HEIF',
    )

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
    folder_path = select_folder("Select folder containing HEIC images")

    if folder_path:
        # Convert HEIC files to JPG
        convert_heic_to_jpg(folder_path)

        # Show completion message
        show_completion_messagebox(
            "Conversion Complete",
            "HEIC to JPG conversion completed!\n\nCheck the selected folder for converted files."
        )
    else:
        logger.info("❌ No folder selected. Exiting.")

if __name__ == "__main__":
    main()
