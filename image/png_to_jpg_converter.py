import logging
from pathlib import Path

from _image_to_jpg_converter import (
    convert_folder_to_jpg,
    flatten_to_rgb,
    select_folder,
    show_completion_messagebox,
)
from PIL import Image, UnidentifiedImageError

logger = logging.getLogger(__name__)

def validate_image_file(file_path: Path) -> bool:
    """Validate if a file is a valid image that can be opened"""
    try:
        with Image.open(file_path) as img:
            # Try to access basic properties to ensure it's a valid image
            img.verify()
            return True
    except (UnidentifiedImageError, OSError, Exception) as e:
        logger.error("Cannot identify image file '%s': %s", file_path, e)
        return False

def convert_png_to_jpg(folder_path: str) -> None:
    """Convert all PNG images in the folder to JPG format with improved error handling"""
    convert_folder_to_jpg(
        folder_path,
        source_extensions=['.png'],
        validate_fn=validate_image_file,
        convert_to_rgb_fn=flatten_to_rgb,
        save_kwargs={'quality': 95, 'optimize': True},
        # Note: unlike heic_to_jpg_converter.py, this does NOT protect an
        # existing same-named .jpg from being overwritten - preserved as-is
        # (pre-existing behavior, not something this dedup changes).
        skip_existing_jpg=False,
        logger=logger,
        kind_label='PNG',
    )

def main() -> None:
    """Main function to run the PNG to JPG converter"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger.info("PNG to JPG Converter (Improved)")
    logger.info("=" * 40)

    # Select folder
    folder_path = select_folder("Select folder containing PNG images")

    if folder_path:
        # Convert PNG files to JPG
        convert_png_to_jpg(folder_path)

        # Show completion message
        show_completion_messagebox(
            "Conversion Complete",
            "PNG to JPG conversion completed!\n\nCheck the selected folder for converted files."
        )
    else:
        logger.info("No folder selected. Exiting.")

if __name__ == "__main__":
    main()
