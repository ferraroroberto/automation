import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, UnidentifiedImageError
import os
from pathlib import Path
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def select_folder():
    """Open folder dialog and return selected folder path"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    folder_path = filedialog.askdirectory(
        title="Select folder containing PNG images"
    )
    
    root.destroy()
    return folder_path

def validate_image_file(file_path):
    """Validate if a file is a valid image that can be opened"""
    try:
        with Image.open(file_path) as img:
            # Try to access basic properties to ensure it's a valid image
            img.verify()
            return True
    except (UnidentifiedImageError, OSError, Exception) as e:
        logger.error(f"Cannot identify image file '{file_path}': {str(e)}")
        return False

def convert_png_to_jpg(folder_path):
    """Convert all PNG images in the folder to JPG format with improved error handling"""
    if not folder_path:
        print("No folder selected.")
        return
    
    folder = Path(folder_path)
    if not folder.exists():
        print(f"Folder does not exist: {folder_path}")
        return
    
    # Find all PNG files in the folder (case insensitive)
    png_files = []
    for file in folder.iterdir():
        if file.is_file() and file.suffix.lower() == '.png':
            png_files.append(file)
    
    if not png_files:
        print("No PNG files found in the selected folder.")
        return
    
    converted_count = 0
    error_count = 0
    
    print(f"Found {len(png_files)} PNG files in {folder_path}")
    print("Starting conversion...")
    
    for png_file in png_files:
        # Validate the image file first
        if not validate_image_file(png_file):
            print(f"Skipping {png_file.name} - invalid or corrupted image file")
            error_count += 1
            continue
        
        # Create JPG filename with same name but .jpg extension
        jpg_file = png_file.with_suffix('.jpg')
        
        try:
            # Open PNG image with explicit format
            with Image.open(png_file) as img:
                # Force load the image data
                img.load()
                
                # Convert to RGB if necessary (JPG doesn't support transparency)
                if img.mode in ('RGBA', 'LA', 'P'):
                    # Create white background for transparent images
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    if img.mode == 'P':
                        img = img.convert('RGBA')
                    background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                    img = background
                elif img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # Save as JPG with explicit format
                img.save(jpg_file, 'JPEG', quality=95, optimize=True)
                print(f"Converted {png_file.name} -> {jpg_file.name}")
                converted_count += 1
                
        except UnidentifiedImageError as e:
            print(f"Error: Cannot identify image file '{png_file}': {str(e)}")
            error_count += 1
        except OSError as e:
            print(f"Error: OS error processing '{png_file}': {str(e)}")
            error_count += 1
        except Exception as e:
            print(f"Error converting {png_file.name}: {str(e)}")
            error_count += 1
    
    print(f"\nConversion complete!")
    print(f"Converted: {converted_count} files")
    print(f"Errors: {error_count} files")
    
    if error_count > 0:
        print(f"\nNote: {error_count} files had errors and were skipped.")
        print("This could be due to corrupted files, unsupported formats, or permission issues.")

def main():
    """Main function to run the PNG to JPG converter"""
    print("PNG to JPG Converter (Improved)")
    print("=" * 40)
    
    # Select folder
    folder_path = select_folder()
    
    if folder_path:
        # Convert PNG files to JPG
        convert_png_to_jpg(folder_path)
        
        # Show completion message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Conversion Complete", 
            f"PNG to JPG conversion completed!\n\nCheck the selected folder for converted files."
        )
        root.destroy()
    else:
        print("No folder selected. Exiting.")

if __name__ == "__main__":
    main() 