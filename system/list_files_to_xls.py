#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File List to Excel Converter

A utility tool that creates an Excel spreadsheet listing all files in a selected directory.
The tool provides a GUI interface for folder selection and generates a structured Excel file
with file names and extensions.

Features:
- GUI-based folder selection using tkinter
- Excel output with file metadata
- Configurable output formats and columns
- Error handling and user feedback
- Cross-platform compatibility

Author: Automation Tools Collection
License: MIT
"""

import sys
import logging
from pathlib import Path
from typing import List, Tuple, Optional
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
except ImportError as e:
    print("Error: openpyxl is required but not installed.")
    print("Install it with: pip install openpyxl")
    sys.exit(1)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class FileListExporter:
    """
    A class to handle file listing and Excel export operations.
    
    This class provides methods to scan directories, collect file information,
    and export the data to Excel format with proper formatting and error handling.
    """
    
    def __init__(self, output_filename: str = "lista.xlsx"):
        """
        Initialize the FileListExporter.
        
        Args:
            output_filename (str): Name of the output Excel file
        """
        self.output_filename = output_filename
        self.workbook = None
        self.worksheet = None
        
    def select_folder(self) -> Optional[Path]:
        """
        Open a folder selection dialog.
        
        Returns:
            Optional[Path]: Selected folder path or None if cancelled
        """
        try:
            # Hide the root Tkinter window
            root = tk.Tk()
            root.withdraw()
            
            folder_path = filedialog.askdirectory(
                title="Select a folder to list files from"
            )
            
            if not folder_path:
                logger.info("Folder selection cancelled by user")
                return None
                
            return Path(folder_path)
            
        except Exception as e:
            logger.error(f"Error selecting folder: {e}")
            messagebox.showerror("Error", f"Failed to select folder: {e}")
            return None
            
    def scan_directory(self, folder_path: Path) -> List[Tuple[str, str, str, str]]:
        """
        Scan directory and collect file information.
        
        Args:
            folder_path (Path): Path to the directory to scan
            
        Returns:
            List[Tuple[str, str, str, str]]: List of tuples containing 
                (filename, extension, size, modified_date)
        """
        files_info = []
        
        try:
            for item in sorted(folder_path.iterdir(), key=lambda p: p.name.lower()):
                if item.is_file():
                    # Get file information
                    filename = item.name
                    extension = item.suffix[1:] if item.suffix.startswith(".") else ""
                    
                    # Get file size in bytes
                    try:
                        size = item.stat().st_size
                        size_str = self._format_file_size(size)
                    except OSError:
                        size_str = "N/A"
                    
                    # Get modification date
                    try:
                        modified_date = item.stat().st_mtime
                        date_str = self._format_timestamp(modified_date)
                    except OSError:
                        date_str = "N/A"
                    
                    files_info.append((filename, extension, size_str, date_str))
                    
            logger.info(f"Found {len(files_info)} files in directory")
            return files_info
            
        except Exception as e:
            logger.error(f"Error scanning directory: {e}")
            raise
            
    def _format_file_size(self, size_bytes: int) -> str:
        """
        Format file size in human-readable format.
        
        Args:
            size_bytes (int): File size in bytes
            
        Returns:
            str: Formatted file size string
        """
        if size_bytes == 0:
            return "0 B"
            
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
            
        return f"{size_bytes:.1f} {size_names[i]}"
        
    def _format_timestamp(self, timestamp: float) -> str:
        """
        Format timestamp to readable date string.
        
        Args:
            timestamp (float): Unix timestamp
            
        Returns:
            str: Formatted date string
        """
        from datetime import datetime
        try:
            dt = datetime.fromtimestamp(timestamp)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except (OSError, ValueError):
            return "N/A"
            
    def create_excel_workbook(self) -> None:
        """
        Create and configure the Excel workbook and worksheet.
        """
        try:
            self.workbook = Workbook()
            self.worksheet = self.workbook.active
            self.worksheet.title = "File List"
            
            # Set up headers with formatting
            headers = ["Filename", "Extension", "Size", "Modified Date"]
            for col, header in enumerate(headers, 1):
                cell = self.worksheet.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
                
            logger.info("Excel workbook created successfully")
            
        except Exception as e:
            logger.error(f"Error creating Excel workbook: {e}")
            raise
            
    def populate_worksheet(self, files_info: List[Tuple[str, str, str, str]]) -> None:
        """
        Populate the worksheet with file information.
        
        Args:
            files_info (List[Tuple[str, str, str, str]]): List of file information tuples
        """
        try:
            for row, (filename, extension, size, date) in enumerate(files_info, 2):
                self.worksheet.cell(row=row, column=1, value=filename)
                self.worksheet.cell(row=row, column=2, value=extension)
                self.worksheet.cell(row=row, column=3, value=size)
                self.worksheet.cell(row=row, column=4, value=date)
                
            # Auto-adjust column widths
            for column in self.worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                self.worksheet.column_dimensions[column_letter].width = adjusted_width
                
            logger.info(f"Worksheet populated with {len(files_info)} files")
            
        except Exception as e:
            logger.error(f"Error populating worksheet: {e}")
            raise
            
    def save_workbook(self, output_path: Path) -> bool:
        """
        Save the workbook to the specified path.
        
        Args:
            output_path (Path): Full path where to save the Excel file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            self.workbook.save(str(output_path))
            logger.info(f"Excel file saved successfully to: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving Excel file: {e}")
            messagebox.showerror("Error", f"Failed to save Excel file: {e}")
            return False


def main() -> None:
    """
    Main function to run the file list to Excel converter.
    
    This function orchestrates the entire process:
    1. Folder selection
    2. Directory scanning
    3. Excel creation and population
    4. File saving
    5. User feedback
    """
    try:
        logger.info("Starting File List to Excel converter")
        
        # Initialize exporter
        exporter = FileListExporter()
        
        # Select folder
        folder_path = exporter.select_folder()
        if not folder_path:
            logger.info("No folder selected, exiting")
            sys.exit(0)
            
        logger.info(f"Selected folder: {folder_path}")
        
        # Scan directory
        files_info = exporter.scan_directory(folder_path)
        
        if not files_info:
            messagebox.showinfo("Information", "No files found in the selected directory.")
            return
            
        # Create Excel workbook
        exporter.create_excel_workbook()
        
        # Populate with file information
        exporter.populate_worksheet(files_info)
        
        # Save to selected folder
        output_path = folder_path / exporter.output_filename
        if exporter.save_workbook(output_path):
            messagebox.showinfo(
                "Success",
                f"Excel file created successfully!\n\n"
                f"Saved to: {output_path}\n"
                f"Files listed: {len(files_info)}"
            )
            logger.info("Process completed successfully")
        else:
            logger.error("Failed to save Excel file")
            
    except Exception as e:
        logger.error(f"Unexpected error in main function: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
