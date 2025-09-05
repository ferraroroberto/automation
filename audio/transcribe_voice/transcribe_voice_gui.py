"""
GUI module for the voice transcription application.
This module provides a graphical interface for the transcription functionality.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import threading
import os
import sys
import time
import pyperclip
import tempfile
import scipy.io.wavfile as wav
from typing import Optional

# Add CUDA check
try:
    import torch
    CUDA_AVAILABLE = torch.cuda.is_available()
except ImportError:
    CUDA_AVAILABLE = False

# Import core transcription functionality
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from transcribe_voice_core import (
    TranscriptionConfig,
    AudioRecorder,
    Transcriber,
    transcribe_file
)

class TranscriptionGUI:
    """Main GUI application for voice transcription."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Voice Transcription")
        self.root.geometry("400x400")
        self.root.configure(background='#2E2E2E')
        self.root.resizable(False, False)
        
        # Configuration
        self.config = TranscriptionConfig()
        self.recorder = None
        self.transcriber = None
        
        # GUI state
        self.recording_started = False
        self.update_queue = queue.Queue()
        self.gui_should_stop = False
        
        # GUI variables
        self.selected_language = tk.StringVar(value="Spanish")
        self.selected_model_size = tk.StringVar(value="medium")
        self.progress_var = tk.DoubleVar()
        self.level_var = tk.DoubleVar()
        
        # Show CUDA status in the GUI or console
        if CUDA_AVAILABLE:
            print("✅ CUDA is available. Will use GPU for transcription.")
        else:
            print("⚠️  CUDA not available. Will use CPU (slower).")
        
        # Initialize GUI
        self._setup_styles()
        self._create_widgets()
        self._center_window()
        
        # Set up window close handler
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
    def _setup_styles(self):
        """Configure ttk styles."""
        s = ttk.Style()
        s.theme_use('default')
        s.configure("Horizontal.TProgressbar", thickness=20)
        s.configure("Green.Horizontal.TProgressbar", foreground='green', background='green')
        s.configure("Yellow.Horizontal.TProgressbar", foreground='#FFA500', background='#FFA500')
        s.configure("Orange.Horizontal.TProgressbar", foreground='orange', background='orange')
        s.configure("Red.Horizontal.TProgressbar", foreground='red', background='red')
        s.configure("TButton", padding=10)
        
    def _create_widgets(self):
        """Create and layout all widgets."""
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create both views
        self._create_main_view()
        self._create_recording_view()
        
        # Start with main view
        self._show_main_view()
        
    def _create_main_view(self):
        """Create the main selection view."""
        self.main_view_frame = ttk.Frame(self.main_frame)
        
        # Title
        title_label = ttk.Label(self.main_view_frame, text="Voice Transcription", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Language selection
        language_frame = ttk.Frame(self.main_view_frame)
        language_frame.pack(fill=tk.X, pady=5)
        
        language_label = ttk.Label(language_frame, text="Select Language:")
        language_label.pack(side=tk.LEFT, padx=(0, 10))
        
        language_combo = ttk.Combobox(language_frame, textvariable=self.selected_language, 
                                     state="readonly", width=15)
        language_combo['values'] = ('Spanish', 'English')
        language_combo.pack(side=tk.LEFT)
        
        # Model size selection
        model_size_frame = ttk.Frame(self.main_view_frame)
        model_size_frame.pack(fill=tk.X, pady=5)
        
        model_size_label = ttk.Label(model_size_frame, text="Select Model Size:")
        model_size_label.pack(side=tk.LEFT, padx=(0, 10))
        
        model_size_combo = ttk.Combobox(model_size_frame, textvariable=self.selected_model_size, 
                                       state="readonly", width=15)
        model_size_combo['values'] = ('tiny', 'base', 'small', 'medium', 'large')
        model_size_combo.pack(side=tk.LEFT)
        
        # Help text
        help_text = ttk.Label(self.main_view_frame, 
                             text="Spanish: Transcribe & Translate | English: Transcribe only", 
                             font=("Arial", 9), foreground="#888888")
        help_text.pack(pady=5)
        
        # Separator
        separator = ttk.Separator(self.main_view_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=10)
        
        # Options
        options_label = ttk.Label(self.main_view_frame, text="Choose an option:", 
                                 font=("Arial", 12))
        options_label.pack(pady=(0, 5))
        
        # Button container for even spacing
        button_container = ttk.Frame(self.main_view_frame)
        button_container.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Buttons with equal spacing
        start_button = ttk.Button(button_container, text="🎤 Start Recording", 
                                 command=self._start_recording)
        start_button.pack(pady=10, fill=tk.X)
        
        file_button = ttk.Button(button_container, text="📁 Select Audio File", 
                                command=self._select_audio_file)
        file_button.pack(pady=10, fill=tk.X)
        
        # Exit button
        exit_button = ttk.Button(button_container, text="Exit", 
                                command=self._on_close)
        exit_button.pack(pady=10, fill=tk.X)
        
    def _create_recording_view(self):
        """Create the recording view."""
        self.recording_view_frame = ttk.Frame(self.main_frame)
        
        # Time remaining
        self.time_label = ttk.Label(self.recording_view_frame, 
                                   text=f"Time remaining: {self.config.record_seconds}s")
        self.time_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(self.recording_view_frame, 
                                           variable=self.progress_var, length=360)
        self.progress_bar.pack(fill=tk.X, pady=(5, 15))
        
        # Audio level
        self.level_label = ttk.Label(self.recording_view_frame, text="Audio level: 0.00")
        self.level_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Level bar
        self.level_bar = ttk.Progressbar(self.recording_view_frame, 
                                        variable=self.level_var, 
                                        style="Green.Horizontal.TProgressbar", 
                                        length=360)
        self.level_bar.pack(fill=tk.X, pady=(5, 15))
        
        # Stop button
        button_frame = ttk.Frame(self.recording_view_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        stop_button = ttk.Button(button_frame, text="Stop Recording", 
                                command=self._stop_recording)
        stop_button.pack(pady=5, ipady=5, fill=tk.X)
        
    def _show_main_view(self):
        """Show the main selection view."""
        self.recording_view_frame.pack_forget()
        self.main_view_frame.pack(fill=tk.BOTH, expand=True)
        
    def _show_recording_view(self):
        """Show the recording view."""
        self.main_view_frame.pack_forget()
        self.recording_view_frame.pack(fill=tk.BOTH, expand=True)
        
    def _center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def _on_close(self):
        """Handle window close event."""
        self.gui_should_stop = True
        if self.recorder:
            self.recorder.key_pressed = True
            self.recorder.cleanup_ffmpeg_copy()
        self.root.quit()
        self.root.destroy()
        
    def _start_recording(self):
        """Start the recording process."""
        # Update configuration
        self.config.language = self.selected_language.get()
        self.config.model_size = self.selected_model_size.get()
        self.config.translate = (self.config.language == "Spanish")
        
        # Reset progress bars
        self.progress_var.set(100)
        self.level_var.set(0)
        
        # Set device in config if possible
        self.config.device = "cuda" if CUDA_AVAILABLE else "cpu"
        
        # Show recording view
        self._show_recording_view()
        
        # Start recording in a separate thread
        self.recording_thread = threading.Thread(target=self._recording_worker)
        self.recording_thread.daemon = True
        self.recording_thread.start()
        
        # Start update checking
        self.root.after(100, self._check_update_queue)
        
    def _recording_worker(self):
        """Worker thread for recording."""
        try:
            print("Starting recording worker...", flush=True)
            # Initialize recorder
            self.recorder = AudioRecorder(self.config)
            print("Recorder initialized.", flush=True)
            
            # Check dependencies
            if not self.recorder.check_dependencies():
                print("Dependencies not available.", flush=True)
                self.update_queue.put(("error", "Required dependencies not available."))
                return
            
            # Get audio devices
            input_devices = self.recorder.get_audio_devices()
            print(f"Input devices: {input_devices}", flush=True)
            if not input_devices:
                self.update_queue.put(("error", "No input devices found!"))
                return
            
            # Select microphone
            selected_device = self.recorder.select_microphone(input_devices, verbose=False)
            print(f"Selected device: {selected_device}", flush=True)
            
            # Progress callback
            def progress_callback(remaining, level):
                self.update_queue.put(("progress", (remaining, level)))
            
            # Record audio
            print("Starting audio recording...", flush=True)
            recording = self.recorder.record_audio(selected_device, progress_callback)
            print("Audio recording finished.", flush=True)
            
            if recording is None or len(recording) == 0:
                self.update_queue.put(("info", "Recording too short, nothing to transcribe."))
                self.update_queue.put(("return_to_main", None))
                return
            
            # Save to temporary file
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmpfile:
                wav.write(tmpfile.name, 16000, recording)
                audio_path = tmpfile.name
            print(f"Audio saved to {audio_path}", flush=True)
            
            # Show transcribing status
            self.update_queue.put(("status", "Loading model and transcribing..."))
            print("Loading model...", flush=True)
            
            # Initialize transcriber
            if not self.transcriber:
                self.transcriber = Transcriber(self.config)
                self.transcriber.load_model(self.recorder)
            print("Model loaded.", flush=True)
            
            # Transcribe
            print("Starting transcription...", flush=True)
            text = self.transcriber.transcribe_audio(audio_path)
            print("Transcription finished.", flush=True)
            
            # Clean up temp file
            if self.config.clean_temp and os.path.exists(audio_path):
                try:
                    os.remove(audio_path)
                except Exception as cleanup_exc:
                    print(f"[WARNING] Could not remove temp file: {cleanup_exc}", flush=True)
            
            # Show result
            self.update_queue.put(("result", text))
            print("Result put in queue.", flush=True)
            
        except Exception as e:
            import traceback
            print("[EXCEPTION] Exception in _recording_worker:", flush=True)
            traceback.print_exc()
            self.update_queue.put(("error", str(e)))
            
    def _stop_recording(self):
        """Stop the current recording."""
        if self.recorder:
            self.recorder.key_pressed = True
            
    def _select_audio_file(self):
        """Select and transcribe an audio file."""
        # Update configuration
        self.config.language = self.selected_language.get()
        self.config.model_size = self.selected_model_size.get()
        self.config.translate = (self.config.language == "Spanish")
        
        # Open file dialog
        filetypes = (
            ('Audio files', '*.mp3 *.wav *.m4a *.flac *.ogg *.mp4 *.avi *.mkv *.mov *.webm'),
            ('All files', '*.*')
        )
        
        filename = filedialog.askopenfilename(
            title='Select audio file',
            initialdir=os.path.expanduser('~'),
            filetypes=filetypes
        )
        
        if filename:
            # Show status window
            self._show_status_window("Transcribing", 
                                   f"File: {os.path.basename(filename)}\n"
                                   "Loading model and transcribing...\n"
                                   "This may take a moment.")
            
            # Start transcription in a separate thread
            transcribe_thread = threading.Thread(
                target=self._transcribe_file_worker,
                args=(filename,)
            )
            transcribe_thread.daemon = True
            transcribe_thread.start()
            # Ensure the update queue is checked
            self.root.after(100, self._check_update_queue)
            
    def _transcribe_file_worker(self, filename):
        """Worker thread for file transcription."""
        try:
            print(f"Starting file transcription for {filename}", flush=True)
            text = transcribe_file(filename, self.config)
            print("File transcription finished.", flush=True)
            self.update_queue.put(("result", text))
        except Exception as e:
            import traceback
            print("[EXCEPTION] Exception in _transcribe_file_worker:", flush=True)
            traceback.print_exc()
            self.update_queue.put(("error", str(e)))
            
    def _check_update_queue(self):
        """Check for updates from worker threads."""
        if self.gui_should_stop:
            return
            
        try:
            while True:
                update_type, data = self.update_queue.get_nowait()
                
                if update_type == "progress":
                    self._update_progress(*data)
                elif update_type == "status":
                    self._show_status_window("Transcribing", data)
                elif update_type == "result":
                    self._hide_status_window()
                    try:
                        print("Showing result window...", flush=True)
                        self._show_result(data)
                    except Exception as e:
                        import traceback
                        print("[EXCEPTION] Exception in _show_result:", flush=True)
                        traceback.print_exc()
                elif update_type == "error":
                    self._hide_status_window()
                    messagebox.showerror("Error", data)
                    self._show_main_view()
                elif update_type == "info":
                    self._hide_status_window()
                    messagebox.showinfo("Info", data)
                elif update_type == "return_to_main":
                    self._show_main_view()
                self.update_queue.task_done()
        except queue.Empty:
            pass
            
        # Schedule next check
        if not self.gui_should_stop:
            self.root.after(100, self._check_update_queue)
            
    def _update_progress(self, remaining, level):
        """Update progress bars during recording."""
        # Update progress bar
        self.progress_var.set(100 * remaining / self.config.record_seconds)
        
        # Update level indicator
        level_percentage = min(100, level * 100)
        self.level_var.set(level_percentage)
        
        # Update color based on level
        if level < 0.01:
            self.level_bar.config(style="Red.Horizontal.TProgressbar")
            level_text = "VERY LOW AUDIO!"
        elif level < 0.05:
            self.level_bar.config(style="Yellow.Horizontal.TProgressbar")
            level_text = "Low volume"
        elif level > 0.8:
            self.level_bar.config(style="Orange.Horizontal.TProgressbar")
            level_text = "TOO LOUD!"
        else:
            self.level_bar.config(style="Green.Horizontal.TProgressbar")
            level_text = "Good level"
            
        # Update labels
        self.time_label.config(text=f"Time remaining: {remaining}s")
        self.level_label.config(text=f"Audio level: {level:.2f} - {level_text}")
        
        # Auto-stop if time is up
        if remaining <= 0:
            self._stop_recording()
            
    def _show_status_window(self, title, message):
        """Show a status window."""
        self.status_window = tk.Toplevel(self.root)
        self.status_window.title(title)
        self.status_window.geometry("400x150")
        self.status_window.resizable(False, False)
        
        # Center the window
        self.status_window.update_idletasks()
        x = (self.status_window.winfo_screenwidth() // 2) - 200
        y = (self.status_window.winfo_screenheight() // 2) - 75
        self.status_window.geometry(f"+{x}+{y}")
        
        # Add content
        frame = ttk.Frame(self.status_window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        label = ttk.Label(frame, text=message, wraplength=350)
        label.pack()
        
        self.status_window.update()
        
    def _hide_status_window(self):
        """Hide the status window if it exists."""
        if hasattr(self, 'status_window') and self.status_window.winfo_exists():
            self.status_window.destroy()
            
    def _show_result(self, text):
        """Show transcription result in a new window."""
        # Create result window
        result_window = tk.Toplevel(self.root)
        result_window.title("Transcription Result")
        result_window.geometry("600x400")
        
        # Center the window
        result_window.update_idletasks()
        x = (result_window.winfo_screenwidth() // 2) - 300
        y = (result_window.winfo_screenheight() // 2) - 200
        result_window.geometry(f"+{x}+{y}")
        
        # Create frame
        frame = ttk.Frame(result_window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Label(frame, text="Transcription Result", 
                          font=("Arial", 14, "bold"))
        header.pack(pady=(0, 10))
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD, 
                             yscrollcommand=scrollbar.set)
        text_widget.insert(tk.END, text)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=text_widget.yview)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # Copy button
        def copy_text():
            pyperclip.copy(text)
            copy_button.config(text="✓ Copied!")
            
        copy_button = ttk.Button(button_frame, text="Copy to Clipboard", 
                               command=copy_text)
        copy_button.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.X, expand=True)
        
        # Close button
        def close_result():
            result_window.destroy()
            self._show_main_view()
        
        close_button = ttk.Button(button_frame, text="Close", 
                                command=close_result)
        close_button.pack(side=tk.RIGHT, padx=5, pady=5, fill=tk.X, expand=True)
        
        # Handle window close (X button)
        result_window.protocol("WM_DELETE_WINDOW", close_result)
        
        # Copy to clipboard automatically
        pyperclip.copy(text)
        
    def run(self):
        """Run the GUI application."""
        self.root.mainloop()
        
        # Cleanup
        if self.recorder:
            self.recorder.cleanup_ffmpeg_copy()


def main():
    """Main entry point for the GUI application."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Voice Transcription GUI")
    parser.add_argument("--launch_from_elgato", action="store_true",
                       help="Launch with GUI interface for Elgato Stream Deck")
    args = parser.parse_args()
    
    # Create and run the GUI
    app = TranscriptionGUI()
    
    try:
        app.run()
    except KeyboardInterrupt:
        print("\n⚠️  Interrupted by user")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        messagebox.showerror("Error", f"Unexpected error: {e}")
    finally:
        # Ensure cleanup
        if app.recorder:
            app.recorder.cleanup_ffmpeg_copy()


if __name__ == "__main__":
    main()