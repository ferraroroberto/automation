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
"""

import os
import sys
import time
import logging
import threading
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
    from tkinter import filedialog, messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("⚠️  tkinter not available. Will use command line input for folder selection.")

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
    print("\n📁 Folder Selection")
    print("=" * 50)
    
    while True:
        folder_input = input("Enter the path to the folder containing PDF files: ").strip()
        
        if not folder_input:
            print("❌ Please enter a valid folder path.")
            continue
        
        folder_path = Path(folder_input)
        
        if not folder_path.exists():
            print(f"❌ Folder does not exist: {folder_path}")
            continue
        
        if not folder_path.is_dir():
            print(f"❌ Path is not a directory: {folder_path}")
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

def main():
    """Main application function."""
    print("🚀 PDF to Images Converter")
    print("=" * 50)
    
    # Setup logging
    setup_logging()
    
    # Select folder
    folder_path = None
    if TKINTER_AVAILABLE:
        print("📁 Using GUI folder selection...")
        folder_path = select_folder_gui()
        if not folder_path:
            print("❌ No folder selected. Exiting...")
            input("Press Enter to exit...")
            return
    else:
        folder_path = select_folder_cli()
        if not folder_path:
            print("❌ No folder selected. Exiting...")
            return
    
    print(f"✅ Selected folder: {folder_path}")
    
    # Ask about DPI
    print("\n⚙️  Configuration")
    print("=" * 50)
    while True:
        dpi_input = input(f"Enter DPI for image conversion (default: {DEFAULT_DPI}): ").strip()
        if not dpi_input:
            dpi = DEFAULT_DPI
            break
        try:
            dpi = int(dpi_input)
            if dpi <= 0:
                print("❌ DPI must be a positive number.")
                continue
            break
        except ValueError:
            print("❌ Please enter a valid number.")
    
    # Ask about deleting PDFs
    while True:
        delete_input = input("Delete source PDF files after conversion? (y/n, default: n): ").strip().lower()
        if not delete_input:
            delete_after = False
            break
        if delete_input in ['y', 'yes']:
            delete_after = True
            break
        elif delete_input in ['n', 'no']:
            delete_after = False
            break
        else:
            print("❌ Please enter 'y' or 'n'.")
    
    # Ask about parallel processing
    while True:
        workers_input = input(f"Number of parallel workers (default: {DEFAULT_WORKERS}): ").strip()
        if not workers_input:
            num_workers = DEFAULT_WORKERS
            break
        try:
            num_workers = int(workers_input)
            if num_workers <= 0:
                print("❌ Number of workers must be a positive number.")
                continue
            break
        except ValueError:
            print("❌ Please enter a valid number.")
    
    # Confirm settings
    print(f"\n📋 Settings Summary")
    print("=" * 50)
    print(f"📁 Folder: {folder_path}")
    print(f"🎯 DPI: {dpi}")
    print(f"🗑️  Delete PDFs after conversion: {'Yes' if delete_after else 'No'}")
    print(f"⚡ Parallel workers: {num_workers}")
    
    # Ask for confirmation
    while True:
        confirm = input("\nProceed with conversion? (y/n): ").strip().lower()
        if confirm in ['y', 'yes']:
            break
        elif confirm in ['n', 'no']:
            print("❌ Conversion cancelled.")
            input("Press Enter to exit...")
            return
        else:
            print("❌ Please enter 'y' or 'n'.")
    
    # Start conversion
    print(f"\n🚀 Starting PDF conversion...")
    print("=" * 50)
    
    try:
        pdf_count, image_count = process_pdf_files(
            folder_path=folder_path,
            dpi=dpi,
            delete_after=delete_after,
            num_workers=num_workers
        )
        
        print(f"\n🎉 Conversion completed successfully!")
        print(f"📊 Results: {pdf_count} PDFs processed, {image_count} images generated")
        
    except KeyboardInterrupt:
        print("\n❌ Conversion interrupted by user.")
    except Exception as e:
        print(f"\n❌ An error occurred during conversion: {e}")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
