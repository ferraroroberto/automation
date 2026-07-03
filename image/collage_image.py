import logging
import os
import random
from typing import List, Tuple
from PIL import Image

log = logging.getLogger(__name__)

# Folder containing images — set the COLLAGE_SOURCE_FOLDER environment variable on each machine
source_folder = os.getenv("COLLAGE_SOURCE_FOLDER", "")

# Base dimensions for an A4 sheet at 300 DPI
BASE_WIDTH = 2480  # A4 width in pixels at 300 DPI
BASE_HEIGHT = 3508  # A4 height in pixels at 300 DPI
DPI = 300

# Desired A4 aspect ratio
A4_RATIO = 0.707091103

# Output file name
output_file = os.path.join(source_folder, "collage_a4.png")

# Function to check and get all PNG images in the folder and its subfolders
def get_all_images(folder: str) -> List[str]:
    image_list = []
    if not os.path.exists(folder):
        log.error("Error: The folder '%s' does not exist.", folder)
        return image_list

    # Walk through all directories and collect PNG images
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(".png"):
                image_list.append(os.path.join(root, file))
    return image_list

# Function to calculate the optimal grid size for A4 ratio
def find_best_grid(num_images: int, max_value: int, a4_ratio: float) -> Tuple[int, int]:
    best_ratio = None
    best_diff = float('inf')
    best_grid = (0, 0)

    # Iterate over possible grid configurations
    for rows in range(1, max_value):
        for cols in range(1, max_value):
            ratio = cols / rows
            diff = abs(ratio - a4_ratio)

            # Check if this ratio is closer to the A4 ratio
            if diff < best_diff and cols * rows <= num_images:
                best_ratio = ratio
                best_diff = diff
                best_grid = (cols, rows)

    return best_grid

# Function to create the collage
def create_collage(image_paths: List[str], output_file: str, base_width: int, base_height: int, dpi: int, a4_ratio: float) -> None:
    if len(image_paths) == 0:
        log.error("Error: No PNG images found in the specified folder.")
        return

    # Determine optimal grid size based on available images and desired A4 ratio
    max_grid_size = 20  # Maximum number of columns and rows to consider
    num_images = len(image_paths)
    cols, rows = find_best_grid(num_images, max_grid_size, a4_ratio)
    log.info("Optimal grid size: %d columns x %d rows", cols, rows)

    # Calculate the dimensions of the final canvas based on the grid size
    canvas_width = base_width * cols // 10  # Adjusted based on grid size
    canvas_height = base_height * rows // 14  # Adjusted based on grid size

    # Calculate image size to fit the grid exactly
    image_size = canvas_width // cols

    # Create a blank canvas with white background
    canvas = Image.new('RGB', (canvas_width, canvas_height), 'white')

    # Use only the required number of images for the grid
    image_paths = image_paths[:cols * rows]

    # Shuffle the images randomly
    random.shuffle(image_paths)

    # Position variables
    x_offset = 0
    y_offset = 0

    for index, image_path in enumerate(image_paths):
        # Open the image
        img = Image.open(image_path)

        # Resize the image to fit in the grid cell while maintaining the square aspect ratio
        img = img.resize((image_size, image_size), Image.Resampling.LANCZOS)

        # Paste the image into the canvas
        canvas.paste(img, (x_offset, y_offset))

        # Update position for next image
        x_offset += image_size

        # Move to the next row if x exceeds the canvas width
        if x_offset + image_size > canvas_width:
            x_offset = 0
            y_offset += image_size

    # Save the collage image
    canvas.save(output_file, dpi=(dpi, dpi))
    log.info("Collage created successfully at %s", output_file)

# Main execution
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    images = get_all_images(source_folder)
    create_collage(images, output_file, BASE_WIDTH, BASE_HEIGHT, DPI, A4_RATIO)
