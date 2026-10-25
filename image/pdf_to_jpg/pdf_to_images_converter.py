#!/usr/bin/env python3
"""
PDF to Images Converter - Standalone Application

A self-contained application that converts PDF files to JPG images.
Can be built into an exe file using PyInstaller.

Features:
- GUI folder selection
- Option to delete source PDF files after conversion
- Console logging (no file logging)
- Parallel processing support
- Progress tracking

Related but intentionally separate from ``image/carrousel_pdf_to_jpg.py``,
which is a pdf2image/poppler-based batch converter with recursive folder
scanning and stale-image cleanup, meant to be run from this repo's own venv
against a folder tree rather than distributed. Pick this module when you need
a distributable single-file exe (see ``README_PDF_Converter.md``); pick the
other for ad-hoc batch runs.
"""

import os
import sys
import time
import logging
from pathlib import Path
from typing import List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# Try to import PyMuPDF
try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ PyMuPDF is required. Install it with: pip install PyMuPDF")
    input("Press Enter to exit...")
    sys.exit(1)

# Try to import tkinter for folder selection
try:
    import tkinter as tk
    from tkinter import filedialog
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    logging.warning("⚠️  tkinter not available. Will use command line input for folder selection.")

# Default settings
DEFAULT_DPI = 150
DEFAULT_WORKERS = 4

# Thread-safe logging lock
_log_lock = Lock()

def setup_logging():
    """Setup console-only logging."""
    # Clear any existing handlers
    logging.getLogger().handlers.clear()
    
    # Create console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                datefmt='%H:%M:%S')
    console_handler.setFormatter(formatter)
    
    # Setup root logger
    logging.getLogger().setLevel(logging.INFO)
    logging.getLogger().addHandler(console_handler)

def select_folder_gui() -> Optional[Path]:
    """Select folder using GUI dialog."""
    if not TKINTER_AVAILABLE:
        return None
    
    # Create and hide the main window
    root = tk.Tk()
    root.withdraw()
    
    # Show folder selection dialog
    folder_path = filedialog.askdirectory(
        title="Select folder containing PDF files",
        initialdir=os.getcwd()
    )
    
    if folder_path:
        return Path(folder_path)
    return None

def select_folder_cli() -> Optional[Path]:
    """Select folder using command line input."""
    log = logging.getLogger(__name__)
    log.info("📁 Folder Selection")
    log.info("=" * 50)

    while True:
        folder_input = input("Enter the path to the folder containing PDF files: ").strip()

        if not folder_input:
            log.error("❌ Please enter a valid folder path.")
            continue

        folder_path = Path(folder_input)

        if not folder_path.exists():
            log.error("❌ Folder does not exist: %s", folder_path)
            continue

        if not folder_path.is_dir():
            log.error("❌ Path is not a directory: %s", folder_path)
            continue

        return folder_path

def get_optimal_workers(requested_workers: int, pdf_count: int) -> Tuple[int, str]:
    """
    Determine optimal number of workers based on system resources.
    
    Args:
        requested_workers: Number of workers requested by user
        pdf_count: Number of PDFs to process
        
    Returns:
        Tuple of (optimal_workers, reason_message)
    """
    # Get system information
    cpu_count = os.cpu_count() or 1
    
    # Start with requested workers
    optimal_workers = requested_workers
    reasons = []
    
    # Rule 1: Don't exceed CPU cores for CPU-bound tasks
    if optimal_workers > cpu_count:
        optimal_workers = cpu_count
        reasons.append(f"limited to {cpu_count} CPU cores")
    
    # Rule 2: Don't use more workers than PDFs
    if optimal_workers > pdf_count:
        optimal_workers = pdf_count
        reasons.append(f"limited to {pdf_count} PDFs")
    
    # Rule 3: Ensure at least 1 worker
    optimal_workers = max(1, optimal_workers)
    
    # Build reason message
    if optimal_workers < requested_workers:
        reason = f"Optimized from {requested_workers} to {optimal_workers} workers"
        if reasons:
            reason += f" ({', '.join(reasons)})"
    else:
        reason = f"Using {optimal_workers} workers as requested"
    
    return optimal_workers, reason

def convert_pdf_to_images(pdf_path: Path, dpi: int = DEFAULT_DPI, delete_after: bool = False, 
                         pdf_index: int = None, total_pdfs: int = None) -> List[Path]:
    """
    Convert a PDF file to JPG images.
    
    Args:
        pdf_path: Path to the PDF file
        dpi: Resolution for the output images (default: 150)
        delete_after: Whether to delete source PDF files after conversion
        pdf_index: Current PDF index for progress tracking
        total_pdfs: Total number of PDFs for progress tracking
        
    Returns:
        List of paths to the generated images
    """
    generated_images = []
    
    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        
        with _log_lock:
            progress_info = f" ({pdf_index}/{total_pdfs})" if pdf_index is not None and total_pdfs is not None else ""
            logging.info(f"📄 Converting PDF: {pdf_path.name} ({total_pages} pages){progress_info}")
        
        for i in range(total_pages):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=dpi)
            
            # Use three digits for page numbers (p001, p002, etc.)
            out_name = f"{pdf_path.stem}_p{i+1:03d}dp{total_pages:03d}.jpg"
            out_path = pdf_path.with_name(out_name)
            
            pix.save(str(out_path))
            generated_images.append(out_path)
            
            with _log_lock:
                logging.info(f"✓ Page {i+1}/{total_pages} converted: {out_name}")
        
        doc.close()
        
        # Log completion with progress information if available
        progress_info = f" ({pdf_index}/{total_pdfs})" if pdf_index is not None and total_pdfs is not None else ""
        with _log_lock:
            logging.info(f"✅ PDF conversion {pdf_path.name} done: {len(generated_images)} images generated{progress_info}")
        
        # Delete source PDF file if requested
        if delete_after and generated_images:
            try:
                os.remove(pdf_path)
                with _log_lock:
                    logging.info(f"🗑️ Deleted source PDF: {pdf_path.name}")
            except Exception as e:
                with _log_lock:
                    logging.warning(f"⚠️ Failed to delete source PDF {pdf_path.name}: {e}")
        
        return generated_images
    
    except Exception as e:
        with _log_lock:
            logging.error(f"❌ Failed to convert PDF {pdf_path.name}: {e}")
        return []

def _process_single_pdf_wrapper(args: Tuple[Path, int, bool, Optional[int], Optional[int]]) -> Tuple[int, List[Path]]:
    """
    Wrapper for convert_pdf_to_images to handle parallel processing.
    
    Args:
        args: Tuple of (pdf_path, dpi, delete_after, pdf_index, total_pdfs)
        
    Returns:
        Tuple of (1 if processed, list of generated image paths)
    """
    pdf_path, dpi, delete_after, pdf_index, total_pdfs = args
    
    try:
        images = convert_pdf_to_images(pdf_path, dpi, delete_after, pdf_index, total_pdfs)
        return (1 if images else 0, images)
    except Exception as e:
        with _log_lock:
            logging.error(f"❌ Error processing PDF {pdf_path.name}: {e}")
        return (0, [])

def process_pdf_files(folder_path: Path, dpi: int = DEFAULT_DPI, delete_after: bool = False, 
                     num_workers: int = DEFAULT_WORKERS) -> Tuple[int, int]:
    """
    Scan a folder for PDF files and convert them to images.
    
    Args:
        folder_path: Path to the folder containing PDF files
        dpi: Resolution for the output images (default: 150)
        delete_after: Whether to delete source PDF files after conversion
        num_workers: Number of parallel workers (default: 4)
        
    Returns:
        Tuple of (number of PDFs processed, number of images generated)
    """
    pdf_count = 0
    image_count = 0
    
    if not folder_path.exists():
        logging.error(f"❌ Folder does not exist: {folder_path}")
        return pdf_count, image_count
    
    logging.info(f"🔍 Scanning for PDF files in {folder_path}")
    
    # Find all PDF files in the folder (non-recursive)
    pdf_files = [f for f in folder_path.glob("*.pdf")]
    
    if not pdf_files:
        logging.info("ℹ️  No PDF files found in the folder")
        return pdf_count, image_count
    
    total_pdfs = len(pdf_files)
    logging.info(f"📚 Found {total_pdfs} PDF files to process")
    
    # Ensure at least 1 worker
    num_workers = max(1, num_workers)
    
    # Record start time for performance metrics
    total_start_time = time.time()
    
    # Check if we should use parallel processing
    if num_workers > 1:
        # Parallel processing mode
        # Determine optimal worker count based on system resources
        effective_workers, optimization_reason = get_optimal_workers(num_workers, total_pdfs)
        logging.info(f"⚡ {optimization_reason}")
        logging.info(f"🚀 Using parallel PDF processing with {effective_workers} workers")
        
        # Prepare arguments for each PDF
        processing_args = [
            (pdf_file, dpi, delete_after, idx, total_pdfs)
            for idx, pdf_file in enumerate(pdf_files, 1)
        ]
        
        # Process PDFs in parallel
        with ThreadPoolExecutor(max_workers=effective_workers) as executor:
            # Submit all tasks
            future_to_args = {
                executor.submit(_process_single_pdf_wrapper, args): args
                for args in processing_args
            }
            
            # Process completed tasks
            for future in as_completed(future_to_args):
                args = future_to_args[future]
                pdf_path = args[0]
                
                try:
                    processed, images = future.result()
                    pdf_count += processed
                    image_count += len(images)
                except Exception as e:
                    with _log_lock:
                        logging.error(f"❌ Failed to process {pdf_path.name}: {e}")
    else:
        # Sequential processing mode
        logging.info("🚀 Using sequential PDF processing")
        
        # Process each PDF file sequentially
        for idx, pdf_file in enumerate(pdf_files, 1):
            pdf_count += 1
            images = convert_pdf_to_images(pdf_file, dpi, delete_after, idx, total_pdfs)
            image_count += len(images)
    
    # Calculate total processing time
    total_elapsed_time = time.time() - total_start_time
    
    logging.info(f"📊 PDF conversion summary: {pdf_count} PDFs processed, {image_count} images generated")
    logging.info(f"⏱️  Total processing time: {total_elapsed_time:.2f} seconds")
    
    return pdf_count, image_count

def _prompt_positive_int(prompt: str, default: int, positive_msg: str, log: logging.Logger) -> int:
    """Ask until the answer is empty (the default) or a positive integer."""
    while True:
        raw = input(prompt).strip()
        if not raw:
            return default
        try:
            value = int(raw)
            if value <= 0:
                log.error(positive_msg)
                continue
            return value
        except ValueError:
            log.error("❌ Please enter a valid number.")


def _prompt_delete_after(log: logging.Logger) -> bool:
    """Ask whether to delete the source PDFs; an empty answer means no."""
    while True:
        answer = input("Delete source PDF files after conversion? (y/n, default: n): ").strip().lower()
        if not answer or answer in ['n', 'no']:
            return False
        if answer in ['y', 'yes']:
            return True
        log.error("❌ Please enter 'y' or 'n'.")


def _confirm_conversion(log: logging.Logger) -> bool:
    """Ask for a final y/n; a refusal is logged as a cancelled conversion."""
    while True:
        answer = input("\nProceed with conversion? (y/n): ").strip().lower()
        if answer in ['y', 'yes']:
            return True
        if answer in ['n', 'no']:
            log.info("❌ Conversion cancelled.")
            return False
        log.error("❌ Please enter 'y' or 'n'.")


def main():
    """Main application function."""
    # Setup logging first so all subsequent calls can use it
    setup_logging()
    log = logging.getLogger(__name__)

    log.info("🚀 PDF to Images Converter")
    log.info("=" * 50)

    # Select folder
    folder_path = None
    if TKINTER_AVAILABLE:
        log.info("📁 Using GUI folder selection...")
        folder_path = select_folder_gui()
        if not folder_path:
            log.error("❌ No folder selected. Exiting...")
            input("Press Enter to exit...")
            return
    else:
        folder_path = select_folder_cli()
        if not folder_path:
            log.error("❌ No folder selected. Exiting...")
            return

    log.info("✅ Selected folder: %s", folder_path)

    # Ask about DPI
    log.info("⚙️  Configuration")
    log.info("=" * 50)
    dpi = _prompt_positive_int(
        f"Enter DPI for image conversion (default: {DEFAULT_DPI}): ",
        DEFAULT_DPI,
        "❌ DPI must be a positive number.",
        log,
    )

    # Ask about deleting PDFs
    delete_after = _prompt_delete_after(log)

    # Ask about parallel processing
    num_workers = _prompt_positive_int(
        f"Number of parallel workers (default: {DEFAULT_WORKERS}): ",
        DEFAULT_WORKERS,
        "❌ Number of workers must be a positive number.",
        log,
    )

    # Confirm settings
    log.info("📋 Settings Summary")
    log.info("=" * 50)
    log.info("📁 Folder: %s", folder_path)
    log.info("🎯 DPI: %d", dpi)
    log.info("🗑️  Delete PDFs after conversion: %s", 'Yes' if delete_after else 'No')
    log.info("⚡ Parallel workers: %d", num_workers)

    # Ask for confirmation
    if not _confirm_conversion(log):
        input("Press Enter to exit...")
        return

    # Start conversion
    log.info("🚀 Starting PDF conversion...")
    log.info("=" * 50)

    try:
        pdf_count, image_count = process_pdf_files(
            folder_path=folder_path,
            dpi=dpi,
            delete_after=delete_after,
            num_workers=num_workers
        )

        log.info("🎉 Conversion completed successfully!")
        log.info("📊 Results: %d PDFs processed, %d images generated", pdf_count, image_count)

    except KeyboardInterrupt:
        log.error("❌ Conversion interrupted by user.")
    except Exception as e:
        log.error("❌ An error occurred during conversion: %s", e)

    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
