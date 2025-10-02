"""
File Lister with GUI
Lists all files in a selected folder with filename+extension and modification date,
sorted by date descending, and saves results to JSON.
"""

import json
import logging
import os
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Dict, List, Optional

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def select_folder() -> Optional[str]:
    """Open folder selection dialog using tkinter.

    Returns:
        Selected folder path or None if cancelled.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring dialog to front

    folder_path = filedialog.askdirectory(title="Select folder to list files")

    if folder_path:
        logger.info(f"📂 Selected folder: {folder_path}")
        return folder_path

    logger.info("❌ No folder selected")
    return None


def get_file_info(file_path: Path) -> Dict[str, str]:
    """Get file information including name and modification date.

    Args:
        file_path: Path object of the file

    Returns:
        Dictionary with filename and modification date
    """
    try:
        stat = file_path.stat()
        modified_time = datetime.fromtimestamp(stat.st_mtime)

        return {
            "filename": file_path.name,
            "modified_date": modified_time.isoformat()
        }
    except OSError as e:
        logger.error(f"❌ Error getting info for {file_path}: {e}")
        return {
            "filename": file_path.name,
            "modified_date": "Error retrieving date"
        }


def list_files_in_folder(folder_path: str) -> List[Dict[str, str]]:
    """List all files in folder with their modification dates.

    Args:
        folder_path: Path to the folder to scan

    Returns:
        List of file information dictionaries, sorted by date descending
    """
    folder = Path(folder_path)

    if not folder.exists():
        raise ValueError(f"Folder does not exist: {folder_path}")

    if not folder.is_dir():
        raise ValueError(f"Path is not a directory: {folder_path}")

    logger.info(f"🔍 Scanning folder: {folder_path}")

    files_info = []

    # Walk through all files in directory and subdirectories
    for file_path in folder.rglob('*'):
        if file_path.is_file():
            file_info = get_file_info(file_path)
            files_info.append(file_info)

    # Sort by modification date descending (newest first)
    files_info.sort(key=lambda x: x['modified_date'], reverse=True)

    logger.info(f"✅ Found {len(files_info)} files")
    return files_info


def save_to_json(data: List[Dict[str, str]], output_path: str) -> None:
    """Save file list data to JSON file.

    Args:
        data: List of file information dictionaries
        output_path: Path where to save the JSON file
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        logger.info(f"💾 JSON saved to: {output_path}")
        logger.info(f"📊 Total files processed: {len(data)}")

    except IOError as e:
        logger.error(f"❌ Error saving JSON to {output_path}: {e}")
        raise


def main() -> None:
    """Main function to run the file lister application."""
    try:
        logger.info("🚀 Starting File Lister application")

        # Select folder using tkinter
        folder_path = select_folder()
        if not folder_path:
            logger.info("ℹ️ Application cancelled by user")
            return

        # List files
        files_data = list_files_in_folder(folder_path)

        if not files_data:
            messagebox.showinfo("No Files Found", f"No files found in {folder_path}")
            logger.warning(f"⚠️ No files found in {folder_path}")
            return

        # Generate output path
        folder_name = Path(folder_path).name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"files_list_{folder_name}_{timestamp}.json"
        output_path = str(Path(folder_path) / output_filename)

        # Save to JSON
        save_to_json(files_data, output_path)

        # Show success message
        messagebox.showinfo(
            "Success",
            f"File list saved to:\n{output_filename}\n\nTotal files: {len(files_data)}"
        )

        logger.info("✅ File listing completed successfully")

    except Exception as e:
        error_msg = f"❌ Application error: {str(e)}"
        logger.error(error_msg)
        messagebox.showerror("Error", error_msg)


if __name__ == "__main__":
    main()
