# chatGPT source and first iteration > https://chatgpt.com/c/5001a6f9-4247-4a02-b6e9-2a84acd468f5
# log execution > https://onedrive.live.com/edit.aspx?resid=5492cf8639ca0b0b!196221


import os
import sys
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Hardcoded source folder
HARDCODED_SOURCE_FOLDER = r"C:\Users\rober\iCloudDrive\6LVTQB9699~com~seriflabs~affinitydesigner\Roberto\thread\books"

# Specify Poppler path if not added to system PATH
POPPLER_PATH = r"E:\onedrive\Documentos\Roberto\projects\automation\notion-automation-files\poppler\Library\bin"  # Update this path as per your installation


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
        print(f"\nProcessing: {pdf_path}")
        try:
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            folder_path = os.path.dirname(pdf_path)
            output_folder = os.path.join(folder_path, "pdf to image")

            # Create the output folder if it doesn't exist
            if not os.path.exists(output_folder):
                try:
                    os.makedirs(output_folder)
                    print(f"Created output folder: {output_folder}")
                except Exception as e:
                    print(f"Error creating output folder {output_folder}: {e}")
                    failed_conversions.append((pdf_path, str(e)))
                    continue

            # Clean up existing images in the output folder
            deleted_files, failed_deletions = clean_up_existing_images(output_folder, pdf_name)
            if deleted_files:
                print(f"Deleted existing images: {', '.join(deleted_files)}")
            if failed_deletions:
                for filename, error in failed_deletions:
                    print(f"Failed to delete {filename}: {error}")

            # Convert PDF to images
            images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)

            for i, image in enumerate(images):
                image_filename = f"{pdf_name}_P{i + 1:02d}.jpg"
                image_path = os.path.join(output_folder, image_filename)

                try:
                    image.save(image_path, "JPEG")
                    total_images_created += 1
                    print(f"Created image: {image_filename}")
                except Exception as e:
                    print(f"Error saving {image_filename}: {e}")
                    failed_conversions.append((image_filename, str(e)))

            total_files_processed += 1

        except Exception as e:
            print(f"Error processing {pdf_path}: {e}")
            failed_conversions.append((pdf_path, str(e)))

    # Summary log
    print("\n=== Conversion Summary ===")
    print(f"Total files processed: {total_files_processed}")
    print(f"Total images created: {total_images_created}")
    if failed_conversions:
        print(f"\nFailed to convert the following files/images:")
        for item, error in failed_conversions:
            print(f"- {item}: {error}")


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
        print(f"The hardcoded source folder '{root_folder}' is not valid.")
        user_choice = input("Do you want to select a folder manually? (Y/N): ").strip().lower()
        if user_choice == 'y':
            root_folder = select_folder()
            if not root_folder:
                print("No folder selected. Exiting.")
                sys.exit()
        else:
            print("Process aborted.")
            sys.exit()

    # Find all PDF files
    pdf_files = find_pdf_files(root_folder)

    # Print summary and ask for confirmation
    print(f"\nFound {len(pdf_files)} PDF file(s) in '{root_folder}' and its subfolders.")
    if not pdf_files:
        print("No PDF files found. Exiting.")
        sys.exit()

    proceed = input("Do you want to continue with the conversion? (Y/N): ").strip().lower()
    if proceed != "y":
        print("Process aborted.")
        sys.exit()

    # Convert PDFs to images
    pdf_to_images(pdf_files)

    print("\nProcess completed.")


if __name__ == "__main__":
    main()