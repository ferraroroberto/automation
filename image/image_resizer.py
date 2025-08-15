import os
import sys
import logging
import argparse
from pathlib import Path
from PIL import Image
import tkinter as tk
from tkinter import filedialog
from typing import Optional

# Configure logging to display logs on the console only
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# Create logger at module level
logger = logging.getLogger(__name__)

# Try to import pillow_heif for HEIC support
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORT = True
except ImportError:
    HEIC_SUPPORT = False
    logger.warning("pillow-heif not installed. HEIC support disabled. Install with: pip install pillow-heif")

SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}
if HEIC_SUPPORT:
    SUPPORTED_FORMATS.update({'.heic', '.heif'})

def resize_image_quality_only(image_path: Path, max_size_kb: float, quality_start: int = 95) -> bool:
    """
    Resize image by reducing only quality (not dimensions) to meet maximum file size requirement.
    Returns True if successful, False otherwise.
    """
    try:
        # Create a temporary file for the resized image
        temp_path = image_path.parent / f"{image_path.stem}_temp{image_path.suffix}"
        
        # Open image
        img = Image.open(image_path)
        
        # Convert RGBA to RGB if necessary for JPEG
        if img.mode == 'RGBA' and image_path.suffix.lower() in ['.jpg', '.jpeg']:
            rgb_img = Image.new('RGB', img.size, (255, 255, 255))
            rgb_img.paste(img, mask=img.split()[3])
            img = rgb_img
        
        # Convert HEIC to RGB for better compatibility when saving
        if image_path.suffix.lower() in ['.heic', '.heif'] and img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Save the image with decreasing quality until the file size is under the limit
        for quality in range(quality_start, 0, -5):
            # Save with current quality without resizing
            if image_path.suffix.lower() in ['.jpg', '.jpeg']:
                img.save(temp_path, optimize=True, quality=quality)
            elif image_path.suffix.lower() in ['.heic', '.heif']:
                # Save HEIC as JPEG with quality control
                temp_path = temp_path.with_suffix('.jpg')
                img.save(temp_path, 'JPEG', optimize=True, quality=quality)
            else:
                img.save(temp_path, optimize=True)
            
            # Check the file size
            new_size_kb = os.path.getsize(temp_path) / 1024
            if new_size_kb <= max_size_kb:
                # Size is good, save with proper name
                if image_path.suffix.lower() in ['.heic', '.heif']:
                    # Save HEIC as JPEG
                    output_path = image_path.parent / f"{image_path.stem}_resize.jpg"
                else:
                    output_path = image_path.parent / f"{image_path.stem}_resize{image_path.suffix}"
                temp_path.rename(output_path)
                logger.info(f"Success: Resized to {new_size_kb:.2f} KB (quality: {quality}, dimensions unchanged)")
                return True
        
        # If we reach here, we couldn't reduce the size enough with quality alone
        logger.warning(f"Failed to resize {image_path} to under {max_size_kb} KB with quality reduction only")
        return False
    except Exception as e:
        logger.error(f"Error processing {image_path}: {str(e)}")
        return False

def resize_image_to_max_size(image_path: Path, max_size_kb: float, quality_start: int = 95) -> bool:
    """
    Resize image to meet maximum file size requirement.
    Returns True if successful, False otherwise.
    """
    try:
        # Create a temporary file for the resized image
        temp_path = image_path.parent / f"{image_path.stem}_temp{image_path.suffix}"
        
        # Open image
        img = Image.open(image_path)
        
        # Convert RGBA to RGB if necessary for JPEG
        if img.mode == 'RGBA' and image_path.suffix.lower() in ['.jpg', '.jpeg']:
            rgb_img = Image.new('RGB', img.size, (255, 255, 255))
            rgb_img.paste(img, mask=img.split()[3])
            img = rgb_img
        
        # Convert HEIC to RGB for better compatibility when saving
        if image_path.suffix.lower() in ['.heic', '.heif'] and img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Save the image with decreasing quality until the file size is under the limit
        for quality in range(quality_start, 0, -5):
            # Resize image to maintain aspect ratio - using LANCZOS instead of deprecated ANTIALIAS
            img.thumbnail((img.width, img.height), Image.LANCZOS)
            
            # Save with current quality
            if image_path.suffix.lower() in ['.jpg', '.jpeg']:
                img.save(temp_path, optimize=True, quality=quality)
            elif image_path.suffix.lower() in ['.heic', '.heif']:
                # Save HEIC as JPEG with quality control
                temp_path = temp_path.with_suffix('.jpg')
                img.save(temp_path, 'JPEG', optimize=True, quality=quality)
            else:
                img.save(temp_path, optimize=True)
            
            # Check the file size
            new_size_kb = os.path.getsize(temp_path) / 1024
            if new_size_kb <= max_size_kb:
                # Size is good, save with proper name
                if image_path.suffix.lower() in ['.heic', '.heif']:
                    # Save HEIC as JPEG
                    output_path = image_path.parent / f"{image_path.stem}_resize.jpg"
                else:
                    output_path = image_path.parent / f"{image_path.stem}_resize{image_path.suffix}"
                temp_path.rename(output_path)
                logger.info(f"Success: Resized to {new_size_kb:.2f} KB (quality: {quality})")
                return True
            
            # Reset image for the next iteration
            img = Image.open(image_path)
        
        logger.warning(f"Failed to resize {image_path} to under {max_size_kb} KB")
        return False
    except Exception as e:
        logger.error(f"Error processing {image_path}: {str(e)}")
        return False

def process_folder(folder_path: Path, max_size_kb: float, preserve_dimensions: bool = False):
    """
    Process all images in the folder, resizing them to the maximum file size.
    """
    for file_path in folder_path.glob('*'):
        if file_path.suffix.lower() in SUPPORTED_FORMATS:
            logger.info(f"Processing {file_path}")
            if preserve_dimensions:
                success = resize_image_quality_only(file_path, max_size_kb)
                # Fall back to normal resize if quality-only didn't work
                if not success:
                    logger.info(f"Trying standard resize with dimension reduction for {file_path}")
                    resize_image_to_max_size(file_path, max_size_kb)
            else:
                resize_image_to_max_size(file_path, max_size_kb)
        else:
            logger.info(f"Skipping unsupported file format: {file_path}")

def select_folder_with_tkinter() -> Optional[Path]:
    """
    Open a folder selection dialog using Tkinter.
    Returns the selected folder path or None if canceled.
    """
    try:
        # Create and hide the main Tkinter window
        root = tk.Tk()
        root.withdraw()
        
        # Open the folder selection dialog
        folder_path = filedialog.askdirectory(title="Select folder with images to resize")
        
        # Close Tkinter
        root.destroy()
        
        if folder_path:
            return Path(folder_path)
        return None
    except Exception as e:
        logger.error(f"Error opening folder dialog: {str(e)}")
        return None

def main():
    parser = argparse.ArgumentParser(description="Resize images in a folder to a maximum file size")
    parser.add_argument('--folder', '-f', type=str, help='Path to the folder containing images')
    parser.add_argument('--max-size', '-m', type=float, default=500, 
                        help='Maximum file size in KB (default: 500)')
    parser.add_argument('--preserve-dimensions', '-p', action='store_true',
                        help='Preserve image dimensions and only reduce quality')
    
    args = parser.parse_args()
    
    # Log HEIC support status
    if HEIC_SUPPORT:
        logger.info("HEIC support enabled")
    else:
        logger.info("HEIC support disabled (install pillow-heif to enable)")
    
    # Process the specified folder or ask user to select one
    folder_path = None
    if args.folder:
        folder_path = Path(args.folder)
    else:
        logger.info("No folder specified via command line. Opening folder selection dialog...")
        folder_path = select_folder_with_tkinter()
    
    # Process the folder if valid
    if folder_path and folder_path.is_dir():
        logger.info(f"Processing folder: {folder_path}")
        process_folder(folder_path, args.max_size, args.preserve_dimensions)
    elif folder_path:
        logger.error(f"Invalid folder path: {folder_path}")
    else:
        logger.error("No folder selected. Exiting.")

if __name__ == "__main__":
    main()