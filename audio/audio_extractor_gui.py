import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import logging
from pathlib import Path
from audio_extractor_core import AudioExtractor

class QueueHandler(logging.Handler):
    """Custom logging handler that puts messages in a queue."""
    
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
    
    def emit(self, record):
        self.log_queue.put(self.format(record))

class AudioExtractorGUI:
    """GUI application for audio extraction."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Audio Track Extractor")
        self.root.geometry("1200x600")
        
        # Initialize audio extractor
        self.extractor = AudioExtractor()
        
        # Queue for thread-safe logging
        self.log_queue = queue.Queue()
        
        # Setup logging
        self.setup_logging()
        
        # Selected files
        self.selected_files = []
        
        # Create GUI
        self.create_widgets()
        
        # Start log polling
        self.poll_log_queue()
    
    def setup_logging(self):
        """Setup logging to display in the GUI."""
        # Create queue handler
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        # Get root logger
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)
        logger.addHandler(queue_handler)
    
    def create_widgets(self):
        """Create GUI widgets."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # Select files button
        self.select_btn = ttk.Button(file_frame, text="Select Video Files", command=self.select_files)
        self.select_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Selected files label
        self.files_label = ttk.Label(file_frame, text="No files selected")
        self.files_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Extract unified checkbox
        self.extract_unified_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Extract unified audio (all tracks)", 
                       variable=self.extract_unified_var).grid(row=0, column=0, sticky=tk.W)
        
        # Extract individual checkbox
        self.extract_individual_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Extract individual tracks", 
                       variable=self.extract_individual_var).grid(row=1, column=0, sticky=tk.W)
        
        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Process Log", padding="10")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Process button
        self.process_btn = ttk.Button(button_frame, text="Extract Audio", command=self.process_files)
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear log button
        self.clear_btn = ttk.Button(button_frame, text="Clear Log", command=self.clear_log)
        self.clear_btn.pack(side=tk.LEFT)

        # Exit button
        self.exit_btn = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        self.exit_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def select_files(self):
        """Open file dialog to select video files."""
        files = filedialog.askopenfilenames(
            title="Select video files",
            filetypes=[('Video files', '*.mkv *.mp4'), ('All files', '*.*')]
        )
        
        if files:
            self.selected_files = [Path(f) for f in files]
            file_names = [f.name for f in self.selected_files]
            self.files_label.config(text=f"{len(self.selected_files)} file(s): {', '.join(file_names[:3])}{'...' if len(file_names) > 3 else ''}")
            self.process_btn.config(state=tk.NORMAL)
        else:
            self.selected_files = []
            self.files_label.config(text="No files selected")
            self.process_btn.config(state=tk.DISABLED)
    
    def process_files(self):
        """Process selected files in a separate thread."""
        if not self.selected_files:
            messagebox.showwarning("No files", "Please select video files first.")
            return
        
        # Check FFmpeg
        if not self.extractor.check_ffmpeg():
            messagebox.showerror("FFmpeg not found", 
                               "FFmpeg is not installed or not in PATH.\n"
                               "Please install FFmpeg to use this application.")
            return
        
        # Disable buttons during processing
        self.select_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)
        self.progress.start(10)
        
        # Process files in a separate thread
        thread = threading.Thread(target=self._process_files_thread)
        thread.daemon = True
        thread.start()
    
    def _process_files_thread(self):
        """Process files in a separate thread."""
        try:
            # Process each file
            for video_file in self.selected_files:
                self.extractor.process_video_file(video_file)
            
            # Show completion message
            self.root.after(0, self._processing_complete)
        except Exception as e:
            self.root.after(0, self._processing_error, str(e))
    
    def _processing_complete(self):
        """Called when processing is complete."""
        self.progress.stop()
        self.select_btn.config(state=tk.NORMAL)
        self.process_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Complete", f"Successfully processed {len(self.selected_files)} file(s).")
    
    def _processing_error(self, error_msg):
        """Called when processing encounters an error."""
        self.progress.stop()
        self.select_btn.config(state=tk.NORMAL)
        self.process_btn.config(state=tk.NORMAL)
        messagebox.showerror("Error", f"An error occurred during processing:\n{error_msg}")
    
    def clear_log(self):
        """Clear the log text area."""
        self.log_text.delete(1.0, tk.END)
    
    def poll_log_queue(self):
        """Poll the log queue and update the log text area."""
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, msg + '\n')
                self.log_text.see(tk.END)
            except queue.Empty:
                break
        
        # Schedule next poll
        self.root.after(100, self.poll_log_queue)

def main():
    """Main function to run the GUI application."""
    root = tk.Tk()
    app = AudioExtractorGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
