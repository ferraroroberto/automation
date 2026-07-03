#!/usr/bin/env python3
"""
Simple password-protected archive file extractor with GUI path selection.
Supports ZIP, 7Z, and RAR files.
"""

import os
import sys
import zipfile
import argparse
import logging
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess

def setup_logging():
    """Setup logging configuration."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

def select_archive_file():
    """Open tkinter dialog to select archive file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select archive file to extract",
        filetypes=[
            ("Archive files", "*.zip;*.7z;*.rar"),
            ("ZIP files", "*.zip"),
            ("7Z files", "*.7z"),
            ("RAR files", "*.rar"),
            ("All files", "*.*")
        ]
    )
    
    root.destroy()
    return file_path

def get_password():
    """Get password from command line input."""
    return input("Enter password for the archive file (press Enter if no password): ").strip()

def try_extract_with_python(archive_path, password=None):
    """Try to extract using Python libraries (zipfile, py7zr, rarfile)."""
    logger = logging.getLogger(__name__)
    
    # Try ZIP extraction
    if archive_path.lower().endswith('.zip') or not archive_path.lower().endswith(('.7z', '.rar')):
        try:
            zip_ref = zipfile.ZipFile(archive_path, 'r')
            if password:
                zip_ref.setpassword(password.encode('utf-8'))
            
            extracted_files = zip_ref.namelist()
            return extracted_files, zip_ref, 'zipfile'
        except Exception as e:
            logger.debug(f"zipfile extraction failed: {e}")
    
    # Try 7Z extraction
    if archive_path.lower().endswith('.7z') or not archive_path.lower().endswith(('.zip', '.rar')):
        try:
            import py7zr  # type: ignore
            archive = py7zr.SevenZipFile(archive_path, mode='r', password=password)
            extracted_files = archive.getnames()
            return extracted_files, archive, 'py7zr'
        except ImportError:
            logger.debug("py7zr library not installed")
        except Exception as e:
            logger.debug(f"py7zr extraction failed: {e}")
    
    # Try RAR extraction
    if archive_path.lower().endswith('.rar') or not archive_path.lower().endswith(('.zip', '.7z')):
        try:
            import rarfile  # type: ignore
            archive = rarfile.RarFile(archive_path, 'r')
            if password:
                archive.setpassword(password)
            
            extracted_files = archive.namelist()
            return extracted_files, archive, 'rarfile'
        except ImportError:
            logger.debug("rarfile library not installed")
        except Exception as e:
            logger.debug(f"rarfile extraction failed: {e}")
    
    return None, None, None

SEVEN_ZIP_DEFAULT_PATHS = [
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
]

def find_7zip() -> str:
    """Return the 7z executable path, checking PATH and common install locations."""
    import shutil
    if shutil.which("7z"):
        return "7z"
    for path in SEVEN_ZIP_DEFAULT_PATHS:
        if os.path.isfile(path):
            return path
    return "7z"  # fall back; will fail with a clear error

def extract_with_7zip(archive_path, output_dir, password=None, seven_zip_path=None):
    """Extract archive using 7-Zip command-line tool."""
    logger = logging.getLogger(__name__)

    if seven_zip_path is None:
        seven_zip_path = find_7zip()
    
    cmd = [seven_zip_path, 'x', archive_path, f'-o{output_dir}', '-y']
    if password:
        cmd.append(f'-p{password}')
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.stdout.strip():
            logger.info("%s", result.stdout.strip())

        if 'Wrong password' in result.stdout or 'Wrong password' in result.stderr:
            logger.error("Extraction failed: Wrong password!")
            return False

        if result.returncode == 0:
            logger.info("Extraction successful!")
            return True
        else:
            logger.error("Extraction failed: %s", result.stderr)
            return False
    except Exception as e:
        logger.error("Error running 7-Zip: %s", e)
        return False

def extract_archive(archive_path, password=None):
    """Extract archive file to a folder with the same name."""
    logger = logging.getLogger(__name__)
    
    if not os.path.exists(archive_path):
        logger.error(f"Archive file not found: {archive_path}")
        return False
    
    # Create output directory (same name as archive file without extension)
    archive_name = Path(archive_path).stem
    archive_dir = Path(archive_path).parent / archive_name
    
    try:
        # Create output directory
        archive_dir.mkdir(exist_ok=True)
        logger.info(f"Created output directory: {archive_dir}")
        
        # Try Python-based extraction first
        extracted_files, archive_ref, method = try_extract_with_python(archive_path, password)
        
        if extracted_files and archive_ref:
            try:
                # Extract all files
                if hasattr(archive_ref, 'extractall'):
                    archive_ref.extractall(archive_dir)
                else:
                    # For py7zr, we need to extract differently
                    archive_ref.extractall(archive_dir)
                
                # Log extracted files
                logger.info(f"Successfully extracted {len(extracted_files)} files using {method}:")
                for file in extracted_files:
                    logger.info(f"  - {file}")
                
                logger.info(f"Extraction completed to: {archive_dir}")
                return True
            finally:
                # Ensure the archive is properly closed
                try:
                    archive_ref.close()
                except Exception as e:
                    logger.debug(f"Error closing archive: {e}")
        
        # If Python-based methods fail, try 7-Zip CLI
        logger.warning("Python-based extraction methods failed. Trying 7-Zip command-line extraction...")
        success = extract_with_7zip(str(archive_path), str(archive_dir), password)
        if success:
            logger.info(f"Extraction completed to: {archive_dir} using 7-Zip CLI.")
            return True
        else:
            logger.error("7-Zip extraction also failed. The file might be:")
            logger.error("1. Corrupted")
            logger.error("2. Using an unsupported compression method")
            logger.error("3. Requiring a different password")
            logger.error("4. Not actually a supported archive format")
            logger.error("\nSuggestion: Try using WinRAR or 7-Zip directly.")
            return False
        
    except Exception as e:
        logger.error(f"Unexpected error during extraction: {e}")
        logger.warning("Trying 7-Zip command-line extraction as fallback...")
        success = extract_with_7zip(str(archive_path), str(archive_dir), password)
        if success:
            logger.info(f"Extraction completed to: {archive_dir} using 7-Zip CLI.")
            return True
        else:
            logger.error("All extraction methods failed.")
            return False

def check_dependencies():
    """Check if optional dependencies are available."""
    logger = logging.getLogger(__name__)
    
    missing_deps = []
    
    try:
        import py7zr  # type: ignore
    except ImportError:
        missing_deps.append("py7zr")
    
    try:
        import rarfile  # type: ignore
    except ImportError:
        missing_deps.append("rarfile")
    
    if missing_deps:
        logger.warning(f"Optional dependencies not installed: {', '.join(missing_deps)}")
        logger.warning("Install them with: pip install py7zr rarfile")
        logger.warning("Note: rarfile requires WinRAR or UnRAR to be installed on your system")
    
    return len(missing_deps) == 0

def main():
    """Main function."""
    parser = argparse.ArgumentParser(description="Extract password-protected archive files (ZIP, 7Z, RAR)")
    parser.add_argument("--archive-path", help="Path to the archive file (optional, will show GUI if not provided)")
    parser.add_argument("--password", help="Password for the archive file (optional, will prompt if not provided)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")
    parser.add_argument("--check-deps", action="store_true", help="Check optional dependencies")
    
    args = parser.parse_args()
    
    # Setup logging
    logger = setup_logging()
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    # Check dependencies if requested
    if args.check_deps:
        check_dependencies()
        return
    
    logger.info("Starting archive extraction process")
    
    # Get archive file path
    archive_path = args.archive_path
    if not archive_path:
        logger.info("Opening file selection dialog...")
        archive_path = select_archive_file()
        
        if not archive_path:
            logger.warning("No file selected. Exiting.")
            return
    
    logger.info(f"Selected archive file: {archive_path}")
    
    # Get password
    password = args.password
    if not password:
        password = get_password()
        if not password:
            logger.info("No password provided, attempting extraction without password")
            password = None
    
    # Extract the archive file
    success = extract_archive(archive_path, password)
    
    if success:
        logger.info("Extraction completed successfully!")
    else:
        logger.error("Extraction failed!")
        logger.info("\nIf you continue having issues, try:")
        logger.info("1. Using WinRAR directly")
        logger.info("2. Installing optional dependencies: pip install py7zr rarfile")
        logger.info("3. Checking if the file is corrupted")
        sys.exit(1)

if __name__ == "__main__":
    main() 