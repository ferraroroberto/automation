#!/usr/bin/env python3
"""
instagram_formatter_gui.py - GUI module for Instagram image formatting

This module provides a tkinter-based graphical user interface for the
Instagram image formatter.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import logging
import json
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
import sys
import os

# Import the core module
try:
    from instagram_formatter import InstagramFormatter, ProcessingResult, parse_color
except ImportError:
    # If running as standalone, try to import from same directory
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from instagram_formatter import InstagramFormatter, ProcessingResult, parse_color


class TextHandler(logging.Handler):
    """Custom logging handler that writes to a tkinter Text widget"""
    
    def __init__(self, text_widget, queue_obj):
        super().__init__()
        self.text_widget = text_widget
        self.queue = queue_obj
    
    def emit(self, record):
        msg = self.format(record)
        # Put message in queue to be processed by GUI thread
        self.queue.put(('log', msg))


class InstagramFormatterGUI:
    """GUI application for Instagram image formatting"""
    
    CONFIG_FILE = 'instagram_formatter_config.json'
    
    def __init__(self, root):
        self.root = root
        self.root.title("Instagram Image Formatter")
        self.root.geometry("800x600")
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Load configuration
        self.config = self.load_config()
        
        # Variables with config defaults
        self.source_folder = tk.StringVar(value=self.config.get('source_folder', ''))
        self.dest_folder = tk.StringVar(value=self.config.get('destination_folder', ''))
        self.aspect_ratio = tk.StringVar(value=self.config.get('aspect_ratio', '3:4'))
        self.bg_color = tk.StringVar(value=self.config.get('background_color', ''))
        self.processing = False
        self.process_thread = None
        
        # Queue for thread communication
        self.queue = queue.Queue()
        
        # Setup UI
        self.setup_ui()
        
        # Setup formatter with custom logger
        self.setup_formatter()
        
        # Start queue monitoring
        self.root.after(100, self.process_queue)
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from JSON file"""
        config_path = Path(__file__).parent / self.CONFIG_FILE
        
        # Default configuration
        default_config = {
            'source_folder': '',
            'destination_folder': '',
            'aspect_ratio': '3:4',
            'background_color': ''
        }
        
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                # Merge with defaults to ensure all keys exist
                default_config.update(config)
                return default_config
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Could not load config file: {e}")
        
        return default_config
    
    def save_config(self):
        """Save current settings to configuration file"""
        config = {
            'source_folder': self.source_folder.get(),
            'destination_folder': self.dest_folder.get(),
            'aspect_ratio': self.aspect_ratio.get(),
            'background_color': self.bg_color.get()
        }
        
        config_path = Path(__file__).parent / self.CONFIG_FILE
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            print(f"Warning: Could not save config file: {e}")
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Instagram Image Formatter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Source folder selection
        ttk.Label(main_frame, text="Source Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        source_entry = ttk.Entry(main_frame, textvariable=self.source_folder, width=50)
        source_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_source).grid(row=1, column=2, pady=5)
        
        # Destination folder selection
        ttk.Label(main_frame, text="Destination Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        dest_entry = ttk.Entry(main_frame, textvariable=self.dest_folder, width=50)
        dest_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_dest).grid(row=2, column=2, pady=5)
        
        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        settings_frame.columnconfigure(1, weight=1)
          # Aspect ratio
        ttk.Label(settings_frame, text="Aspect Ratio:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ratio_frame = ttk.Frame(settings_frame)
        ratio_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Aspect ratio dropdown
        ratio_combo = ttk.Combobox(ratio_frame, textvariable=self.aspect_ratio, width=15)
        ratio_combo['values'] = ['3:4', '4:5', '9:16', '1:1', '16:9']
        ratio_combo.grid(row=0, column=0, sticky=tk.W)
        
        ttk.Label(ratio_frame, text="(Instagram: 4:5, Stories: 9:16)").grid(row=0, column=1, padx=(10, 0))
        
        # Background color
        ttk.Label(settings_frame, text="Background Color:").grid(row=1, column=0, sticky=tk.W, pady=5)
        color_frame = ttk.Frame(settings_frame)
        color_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        color_entry = ttk.Entry(color_frame, textvariable=self.bg_color, width=20)
        color_entry.grid(row=0, column=0, sticky=tk.W)
        ttk.Label(color_frame, text="(Optional: #RRGGBB or R,G,B)").grid(row=0, column=1, padx=(10, 0))
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Process Images", 
                                     command=self.process_images, style='Accent.TButton')
        self.process_btn.grid(row=4, column=0, columnspan=3, pady=20)
          # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(2, weight=1)  # Only the log area should expand
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Status label below progress bar - fixed height, left aligned
        self.status_label = ttk.Label(progress_frame, text="Ready to process images", anchor='w')
        self.status_label.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Log text area - this will expand when window is resized
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure tags for log formatting
        self.log_text.tag_config('INFO', foreground='black')
        self.log_text.tag_config('WARNING', foreground='orange')
        self.log_text.tag_config('ERROR', foreground='red')
        self.log_text.tag_config('SUCCESS', foreground='green')
    
    def setup_formatter(self):
        """Setup the image formatter with custom logger"""
        # Create logger
        logger = logging.getLogger('InstagramFormatter')
        logger.setLevel(logging.INFO)
        
        # Add our custom handler
        text_handler = TextHandler(self.log_text, self.queue)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                    datefmt='%H:%M:%S')
        text_handler.setFormatter(formatter)
        logger.addHandler(text_handler)
        
        # Create formatter instance
        self.formatter = InstagramFormatter(logger)
    
    def browse_source(self):
        """Browse for source folder"""
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)
    
    def browse_dest(self):
        """Browse for destination folder"""
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_folder.set(folder)
    
    def validate_inputs(self) -> Optional[str]:
        """Validate user inputs and return error message if invalid"""
        if not self.source_folder.get():
            return "Please select a source folder"
        
        if not self.dest_folder.get():
            return "Please select a destination folder"
        
        if not Path(self.source_folder.get()).exists():
            return "Source folder does not exist"
        
        if not self.aspect_ratio.get():
            return "Please specify an aspect ratio"
        
        # Validate aspect ratio
        try:
            self.formatter.parse_aspect_ratio(self.aspect_ratio.get())
        except ValueError:
            return "Invalid aspect ratio format"
        
        # Validate color if provided
        if self.bg_color.get():
            try:
                parse_color(self.bg_color.get())
            except ValueError:
                return "Invalid color format. Use #RRGGBB or R,G,B"
        
        return None
    
    def process_images(self):
        """Start processing images in a separate thread"""
        if self.processing:
            messagebox.showwarning("Processing", "Already processing images!")
            return
        
        # Validate inputs
        error = self.validate_inputs()
        if error:
            messagebox.showerror("Validation Error", error)
            return
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Start processing
        self.processing = True
        self.process_btn.config(state='disabled')
        self.progress_var.set(0)
        self.status_label.config(text="Processing...")
        
        # Start processing thread
        self.process_thread = threading.Thread(target=self.process_thread_func)
        self.process_thread.start()
    
    def process_thread_func(self):
        """Function to run in processing thread"""
        try:
            # Parse background color
            bg_color = None
            if self.bg_color.get():
                bg_color = parse_color(self.bg_color.get())
            
            # Process images
            result = self.formatter.process_folder(
                self.source_folder.get(),
                self.dest_folder.get(),
                self.aspect_ratio.get(),
                bg_color,
                progress_callback=self.progress_callback
            )
              # Send completion message
            self.queue.put(('complete', result))
            
        except Exception as e:
            self.queue.put(('error', str(e)))
    
    def progress_callback(self, current: int, total: int, message: str):
        """Callback for progress updates from processing thread"""
        progress = (current / total) * 100 if total > 0 else 0
        self.queue.put(('progress', (progress, f"{current}/{total} - {message}")))
    
    def process_queue(self):
        """Process messages from the queue"""
        try:
            while True:
                msg_type, msg_data = self.queue.get_nowait()
                
                if msg_type == 'log':
                    # Add log message
                    self.log_text.insert(tk.END, msg_data + '\n')
                    self.log_text.see(tk.END)
                    
                elif msg_type == 'progress':
                    # Update progress
                    progress, status = msg_data
                    self.progress_var.set(progress)
                    self.status_label.config(text=status)
                    
                elif msg_type == 'complete':
                    # Processing complete
                    result: ProcessingResult = msg_data
                    self.processing = False
                    self.process_btn.config(state='normal')
                    self.progress_var.set(100)
                    
                    # Save current settings to config
                    self.save_config()
                    
                    # Show summary
                    summary = (f"\nProcessing Complete!\n"
                             f"Total images: {result.total_images}\n"
                             f"Successful: {result.successful}\n"
                             f"Failed: {result.failed}\n"
                             f"Time elapsed: {result.elapsed_time:.2f} seconds")
                    
                    self.log_text.insert(tk.END, summary, 'SUCCESS')
                    self.log_text.see(tk.END)
                    self.status_label.config(text="Processing complete!")
                    
                    # Show message box
                    if result.failed == 0:
                        messagebox.showinfo("Success", 
                                          f"Successfully processed {result.successful} images!")
                    else:
                        messagebox.showwarning("Completed with errors", 
                                             f"Processed {result.successful} images.\n"
                                             f"{result.failed} images failed.")
                    
                elif msg_type == 'error':
                    # Error occurred
                    self.processing = False
                    self.process_btn.config(state='normal')
                    self.status_label.config(text="Error occurred!")
                    messagebox.showerror("Processing Error", msg_data)
                    
        except queue.Empty:
            pass
        
        # Schedule next check
        self.root.after(100, self.process_queue)


def main():
    """Main entry point for GUI application"""
    root = tk.Tk()
    app = InstagramFormatterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()