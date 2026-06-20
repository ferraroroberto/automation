import logging
import os
import sys
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

log = logging.getLogger(__name__)

# Source folder — set the CARROUSEL_SOURCE_FOLDER environment variable on each machine
HARDCODED_SOURCE_FOLDER = os.getenv("CARROUSEL_SOURCE_FOLDER", "")

# Poppler path if not added to system PATH — set the POPPLER_PATH environment variable on each machine
POPPLER_PATH = os.getenv("POPPLER_PATH", "")


def find_pdf_files(root_folder):
    """Recursively find all PDF files in the root folder and its subfolders."""
    pdf_files = []
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith(".pdf"):
                pdf_files.append(os.path.join(dirpath, filename))
    return pdf_files


def clean_up_existing_images(folder_path, pdf_name):
    """Delete existing JPG images that match the PDF file name and the naming convention in the specified folder."""
    deleted_files = []
    failed_deletions = []

    for filename in os.listdir(folder_path):
        if filename.startswith(pdf_name) and filename.endswith(".jpg"):
            file_path = os.path.join(folder_path, filename)
            try:
                os.remove(file_path)
                deleted_files.append(filename)
            except Exception as e:
                failed_deletions.append((filename, str(e)))
    return deleted_files, failed_deletions


def pdf_to_images(pdf_files):
    """Convert each PDF file to JPG images, handling errors and logging the process."""
    total_images_created = 0
    total_files_processed = 0
    failed_conversions = []

    for pdf_path in pdf_files:
        log.info("Processing: %s", pdf_path)
        try:
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            folder_path = os.path.dirname(pdf_path)
            output_folder = os.path.join(folder_path, "pdf to image")

            # Create the output folder if it doesn't exist
            if not os.path.exists(output_folder):
                try:
                    os.makedirs(output_folder)
                    log.info("Created output folder: %s", output_folder)
                except Exception as e:
                    log.error("Error creating output folder %s: %s", output_folder, e)
                    failed_conversions.append((pdf_path, str(e)))
                    continue

            # Clean up existing images in the output folder
            deleted_files, failed_deletions = clean_up_existing_images(output_folder, pdf_name)
            if deleted_files:
                log.info("Deleted existing images: %s", ', '.join(deleted_files))
            if failed_deletions:
                for filename, error in failed_deletions:
                    log.warning("Failed to delete %s: %s", filename, error)

            # Convert PDF to images
            images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)

            for i, image in enumerate(images):
                image_filename = f"{pdf_name}_P{i + 1:02d}.jpg"
                image_path = os.path.join(output_folder, image_filename)

                try:
                    image.save(image_path, "JPEG")
                    total_images_created += 1
                    log.info("Created image: %s", image_filename)
                except Exception as e:
                    log.error("Error saving %s: %s", image_filename, e)
                    failed_conversions.append((image_filename, str(e)))

            total_files_processed += 1

        except Exception as e:
            log.error("Error processing %s: %s", pdf_path, e)
            failed_conversions.append((pdf_path, str(e)))

    # Summary log
    log.info("=== Conversion Summary ===")
    log.info("Total files processed: %d", total_files_processed)
    log.info("Total images created: %d", total_images_created)
    if failed_conversions:
        log.warning("Failed to convert the following files/images:")
        for item, error in failed_conversions:
            log.warning("- %s: %s", item, error)


def select_folder():
    """Open a popup window to select the root folder."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    root_folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
    return root_folder


def main():
    """Main function to execute the PDF to images conversion."""
    root_folder = HARDCODED_SOURCE_FOLDER

    # Check if the hardcoded source folder is valid
    if not os.path.isdir(root_folder):
        log.warning("The hardcoded source folder '%s' is not valid.", root_folder)
        user_choice = input("Do you want to select a folder manually? (Y/N): ").strip().lower()
        if user_choice == 'y':
            root_folder = select_folder()
            if not root_folder:
                log.info("No folder selected. Exiting.")
                sys.exit()
        else:
            log.info("Process aborted.")
            sys.exit()

    # Find all PDF files
    pdf_files = find_pdf_files(root_folder)

    # Print summary and ask for confirmation
    log.info("Found %d PDF file(s) in '%s' and its subfolders.", len(pdf_files), root_folder)
    if not pdf_files:
        log.info("No PDF files found. Exiting.")
        sys.exit()

    proceed = input("Do you want to continue with the conversion? (Y/N): ").strip().lower()
    if proceed != "y":
        log.info("Process aborted.")
        sys.exit()

    # Convert PDFs to images
    pdf_to_images(pdf_files)

    log.info("Process completed.")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    main()
