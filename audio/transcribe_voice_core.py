"""
Core module for audio recording and transcription functionality.
This module can be used standalone from the command line or imported by the GUI module.
"""

import sys
import os
import warnings
import subprocess
import shutil
import whisper
import sounddevice as sd
import scipy.io.wavfile as wav
import tempfile
import pyperclip
import numpy as np
import time
import threading
from pynput import keyboard
from dataclasses import dataclass
from typing import Optional, Tuple, List, Callable
import queue

# Suppress specific whisper warnings for cleaner output
def suppress_whisper_warnings():
    """Suppress known whisper warnings that are not critical."""
    warnings.filterwarnings("ignore", message="FP16 is not supported on CPU; using FP32 instead")
    warnings.filterwarnings("ignore", category=UserWarning, module="whisper")

# Call warning suppression early
suppress_whisper_warnings()

@dataclass
class TranscriptionConfig:
    """Configuration for transcription settings."""
    record_seconds: int = 300
    language: str = "Spanish"
    translate: bool = True
    preferred_mics: List[str] = None
    model_size: str = "base"
    clean_temp: bool = True
    
    def __post_init__(self):
        if self.preferred_mics is None:
            self.preferred_mics = [
                "el gato wave XLR (Elgato Wave XLR)",
                "Wave Link Stream (Elgato Wave:XLR)"
            ]

class AudioRecorder:
    """Handles audio recording functionality."""
    
    def __init__(self, config: TranscriptionConfig):
        self.config = config
        self.audio_buffer = []
        self.stop_recording = False
        self.key_pressed = False
        self.recording = None
        self.start_time = None
        self.gpu_available = None
        self.ffmpeg_path = None
        self.ffmpeg_copy_path = None
        self.ffmpeg_already_cleaned = False
        
    def check_gpu_availability(self) -> bool:
        """Check if a CUDA GPU is available. Perform this check lazily."""
        if self.gpu_available is None:
            try:
                import torch
                if torch.cuda.is_available():
                    gpu_count = torch.cuda.device_count()
                    gpu_name = torch.cuda.get_device_name(0) if gpu_count > 0 else "Unknown"
                    print(f"✅ CUDA GPU detected: {gpu_name} (Count: {gpu_count})")
                    self.gpu_available = True
                else:
                    print("⚠️  CUDA not available. Will use CPU (slower).")
                    self.gpu_available = False
            except ImportError:
                print("⚠️  PyTorch not found. Assuming CPU mode.")
                self.gpu_available = False
        return self.gpu_available
    
    def check_dependencies(self) -> bool:
        """Check if all required dependencies are available."""
        # Check GPU availability
        self.check_gpu_availability()
        
        # Check for FFmpeg
        ffmpeg_found = False
        ffmpeg_path = None
        
        try:
            result = subprocess.run(['ffmpeg', '-version'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                ffmpeg_found = True
                ffmpeg_path = shutil.which('ffmpeg')
                print(f"✅ FFmpeg found in system PATH: {ffmpeg_path}")
        except (FileNotFoundError, subprocess.TimeoutExpired, subprocess.SubprocessError):
            pass
        
        if not ffmpeg_found:
            print("❌ FFmpeg executable not found in system PATH.")
            print("Trying alternative solutions...")
            
            # Try to use imageio-ffmpeg
            try:
                import imageio_ffmpeg
                ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
                print(f"✅ FFmpeg binary found via imageio-ffmpeg: {ffmpeg_path}")
                
                # Test if the binary works
                try:
                    result = subprocess.run([ffmpeg_path, '-version'], 
                                          capture_output=True, text=True, timeout=5)
                    if result.returncode == 0:
                        ffmpeg_found = True
                        
                        # Add the directory to PATH
                        ffmpeg_dir = os.path.dirname(ffmpeg_path)
                        current_path = os.environ.get('PATH', '')
                        if ffmpeg_dir not in current_path:
                            os.environ['PATH'] = ffmpeg_dir + os.pathsep + current_path
                            print(f"✅ Added FFmpeg directory to PATH: {ffmpeg_dir}")
                        
                        # Set environment variables
                        os.environ['FFMPEG_BINARY'] = ffmpeg_path
                        os.environ['IMAGEIO_FFMPEG_EXE'] = ffmpeg_path

                        # --- BEGIN: Copy ffmpeg.exe to script directory if not present ---
                        script_dir = os.path.dirname(os.path.abspath(__file__))
                        ffmpeg_copy_path = os.path.join(script_dir, 'ffmpeg.exe')
                        if os.name == 'nt' and not os.path.exists(ffmpeg_copy_path):
                            try:
                                shutil.copy2(ffmpeg_path, ffmpeg_copy_path)
                                print(f"✅ Copied FFmpeg to script directory: {ffmpeg_copy_path}")
                                self.ffmpeg_copy_path = ffmpeg_copy_path
                            except Exception as e:
                                print(f"⚠️  Could not copy FFmpeg to script directory: {e}")
                        # --- END: Copy ffmpeg.exe to script directory if not present ---

                except Exception as e:
                    print(f"⚠️  Could not test FFmpeg binary: {e}")
                    
            except ImportError:
                print("📦 Installing imageio-ffmpeg...")
                try:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'imageio-ffmpeg'])
                    print("✅ imageio-ffmpeg installed successfully.")
                    return self.check_dependencies()  # Retry after installation
                except subprocess.CalledProcessError:
                    print("❌ ERROR: Could not install imageio-ffmpeg")
                    return False
        
        if ffmpeg_path:
            self.ffmpeg_path = ffmpeg_path
            
        return ffmpeg_found
    
    def get_audio_devices(self) -> List[Tuple[int, dict]]:
        """Get list of available audio input devices."""
        devices = sd.query_devices()
        input_devices = []
        
        for i, device in enumerate(devices):
            if device['max_input_channels'] > 0:
                input_devices.append((i, device))
                
        return input_devices
    
    def select_microphone(self, input_devices: List[Tuple[int, dict]], 
                         verbose: bool = True) -> Optional[int]:
        """Select the best available microphone."""
        if not input_devices:
            return None
            
        # Try to find preferred microphone
        for preferred_mic in self.config.preferred_mics:
            for i, device in input_devices:
                if preferred_mic.lower() in device['name'].lower():
                    if verbose:
                        print(f"\nFound preferred microphone: {device['name']}")
                    return i
        
        # Use first available device as fallback
        selected_device = input_devices[0][0]
        if verbose:
            device_info = sd.query_devices(selected_device)
            print(f"\nNo preferred microphone found. Using: {device_info['name']}")
        
        return selected_device
    
    def monitor_audio_levels(self, device: int, fs: int, 
                           progress_callback: Optional[Callable] = None):
        """Monitor audio levels during recording."""
        def callback(indata, frames, time, status):
            if status:
                print(f"Status: {status}")
            self.audio_buffer.append(np.max(np.abs(indata)))
            
            if len(self.audio_buffer) > 50:
                self.audio_buffer.pop(0)
        
        with sd.InputStream(device=device, channels=1, samplerate=fs, callback=callback):
            while not self.stop_recording and not self.key_pressed:
                if self.audio_buffer and self.start_time:
                    level = self.audio_buffer[-1]
                    remaining = int(self.config.record_seconds - (time.time() - self.start_time))
                    remaining = max(0, remaining)
                    
                    if progress_callback:
                        progress_callback(remaining, level)
                    else:
                        # Terminal visualization
                        self._show_terminal_progress(remaining, level)
                
                time.sleep(0.1)
    
    def _show_terminal_progress(self, remaining: int, level: float):
        """Show progress in terminal mode."""
        meter_width = 50
        bars = int(level * meter_width)
        meter = "▓" * bars + "░" * (meter_width - bars)
        
        status = ""
        if level < 0.01:
            status = "⚠️  VERY LOW AUDIO"
        elif level < 0.05:
            status = "⚠️  Low volume"
        elif level > 0.8:
            status = "⚠️  TOO LOUD"
        else:
            status = "✅ Good level"
        
        remaining_str = f"{remaining}s" if remaining > 0 else "Done"
        output = f"Time left: {remaining_str} | Level: [{meter}] {level:.2f} {status}"
        
        sys.stdout.write("\033[F")  # Move up one line
        sys.stdout.write("\033[K")  # Clear the line
        sys.stdout.write(output)
        sys.stdout.write("\n")
        sys.stdout.flush()
    
    def record_audio(self, device: int, 
                    progress_callback: Optional[Callable] = None) -> Optional[np.ndarray]:
        """Record audio from the specified device."""
        self.audio_buffer = []
        self.stop_recording = False
        self.key_pressed = False
        
        fs = 16000  # Whisper prefers 16kHz audio
        
        # Start recording
        self.start_time = time.time()
        self.recording = sd.rec(int(self.config.record_seconds * fs), 
                              samplerate=fs, channels=1, 
                              dtype='float32', device=device)
        
        # Start monitoring in a separate thread
        monitor_thread = threading.Thread(
            target=self.monitor_audio_levels, 
            args=(device, fs, progress_callback)
        )
        monitor_thread.start()
        
        # Wait for recording to complete or be interrupted
        while not self.key_pressed and (time.time() - self.start_time) < self.config.record_seconds:
            time.sleep(0.1)
        
        # Stop recording
        sd.stop()
        
        # Calculate actual recording length
        elapsed = time.time() - self.start_time
        elapsed = min(elapsed, self.config.record_seconds)
        
        # Trim recording to actual length
        frames_recorded = int(elapsed * fs)
        if frames_recorded > 0 and frames_recorded < len(self.recording):
            self.recording = self.recording[:frames_recorded]
        
        # Signal monitoring thread to stop
        self.stop_recording = True
        monitor_thread.join()
        
        # Check if recording contains audio
        if len(self.recording) > 0:
            max_level = np.max(np.abs(self.recording))
            if max_level < 0.01:
                print(f"\n⚠️ WARNING: Very low audio levels detected (max: {max_level:.4f})")
            else:
                print(f"\nPeak audio level: {max_level:.4f}")
            
            # Convert to int16
            return (self.recording * 32767).astype(np.int16)
        
        return None
    
    def cleanup_ffmpeg_copy(self):
        """Clean up the temporary FFmpeg copy if it exists."""
        if self.ffmpeg_already_cleaned:
            return
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_exe_path = os.path.join(script_dir, 'ffmpeg.exe')
        
        target_path = self.ffmpeg_copy_path if self.ffmpeg_copy_path else ffmpeg_exe_path
        
        if os.path.exists(target_path):
            try:
                os.remove(target_path)
                print(f"✅ Cleaned up FFmpeg copy: {target_path}")
                self.ffmpeg_already_cleaned = True
            except Exception as e:
                print(f"⚠️  Could not remove FFmpeg copy: {e}")
        else:
            self.ffmpeg_already_cleaned = True

class Transcriber:
    """Handles transcription functionality using Whisper."""
    
    def __init__(self, config: TranscriptionConfig):
        self.config = config
        self.model = None
        self.device = None
        
    def load_model(self, recorder: AudioRecorder):
        """Load the Whisper model."""
        # Configure FFmpeg paths
        try:
            import imageio_ffmpeg
            ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
            os.environ['IMAGEIO_FFMPEG_EXE'] = ffmpeg_path
            os.environ['FFMPEG_BINARY'] = ffmpeg_path
        except:
            pass
        
        # Determine device
        self.device = "cuda" if recorder.check_gpu_availability() else "cpu"
        if self.device == "cuda":
            print("🚀 Using GPU acceleration (CUDA)")
        else:
            print("⚠️  Using CPU (will use FP32 precision)")
        
        # Load model
        self.model = whisper.load_model(self.config.model_size, device=self.device)
        
    def transcribe_audio(self, audio_path: str) -> str:
        """Transcribe audio file and return text."""
        if not self.model:
            raise ValueError("Model not loaded. Call load_model() first.")
        
        print("Transcribing...")

        # --- Ensure ffmpeg.exe in script dir is in PATH (for Windows) ---
        script_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_exe_path = os.path.join(script_dir, 'ffmpeg.exe')
        if os.name == 'nt' and os.path.exists(ffmpeg_exe_path):
            current_path = os.environ.get('PATH', '')
            if script_dir not in current_path.split(os.pathsep):
                os.environ['PATH'] = script_dir + os.pathsep + current_path
                print(f"🔧 Added script directory to PATH for FFmpeg: {script_dir}")
        # ---------------------------------------------------------------

        # Set transcription options
        transcribe_options = {
            "language": self.config.language,
            "task": "translate" if self.config.translate else "transcribe"
        }
        
        # Use fp16 only if GPU is available
        if self.device == "cuda":
            transcribe_options["fp16"] = True
            print("Using FP16 precision for faster GPU processing")
        else:
            transcribe_options["fp16"] = False
            print("Using FP32 precision (CPU mode)")
        
        # Transcribe with warnings suppressed
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            result = self.model.transcribe(audio_path, **transcribe_options)
        
        return result["text"]

def prompt_for_language() -> Tuple[str, bool]:
    """Prompt the user to select a language in console mode."""
    print("\n=== LANGUAGE SELECTION ===")
    print("Select language (default: Spanish):")
    print("1. Spanish (Transcribe & Translate)")
    print("2. English (Transcribe only)")
    
    choice = input("Enter choice (1/2 or es/en) [1]: ")

    if not choice:
        choice = "1"
    
    if choice in ["1", "es", "spanish", "español", "s"]:
        language = "Spanish"
        translate = True
        print("✅ Selected: Spanish (with translation)")
    elif choice in ["2", "en", "english", "e"]:
        language = "English"
        translate = False
        print("✅ Selected: English (transcription only)")
    else:
        print("⚠️ Invalid choice. Using default: Spanish (with translation)")
        language = "Spanish"
        translate = True
    
    print()
    return language, translate

def run_console_transcription(config: Optional[TranscriptionConfig] = None):
    """Run transcription in console mode."""
    if config is None:
        config = TranscriptionConfig()
    
    # Prompt for language
    config.language, config.translate = prompt_for_language()
    
    # Initialize recorder
    recorder = AudioRecorder(config)
    
    # Check dependencies
    if not recorder.check_dependencies():
        print("\n❌ ERROR: Required dependencies not available.")
        return 1
    
    # Get and display audio devices
    print("\n=== AVAILABLE MICROPHONES ===\n")
    input_devices = recorder.get_audio_devices()
    
    for i, device in input_devices:
        print(f"{i}: {device['name']}")
    
    # Select microphone
    selected_device = recorder.select_microphone(input_devices)
    if selected_device is None:
        print("No input devices found! Cannot continue.")
        return 1
    
    print(f"\nStarting recording (max {config.record_seconds} seconds)...")
    print("\n=== AUDIO LEVEL METER ===\n")
    print("Press any key to stop recording early.")
    print("\n")  # Create a blank line for the meter
    
    # Set up keyboard listener
    def on_press(key):
        recorder.key_pressed = True
        return False
    
    from pynput import keyboard
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    
    # Record audio
    recording = recorder.record_audio(selected_device)
    
    listener.stop()
    
    if recording is None or len(recording) == 0:
        print("Recording too short, nothing to transcribe.")
        recorder.cleanup_ffmpeg_copy()
        return 0
    
    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmpfile:
        wav.write(tmpfile.name, 16000, recording)
        audio_path = tmpfile.name
        print(f"Audio saved temporarily to: {tmpfile.name}")
    
    # Initialize transcriber and load model
    transcriber = Transcriber(config)
    print("Loading model...")
    transcriber.load_model(recorder)
    
    try:
        # Transcribe
        text = transcriber.transcribe_audio(audio_path)
        
        # Output results
        print("\n=== FINAL OUTPUT ===\n")
        print(text)
        pyperclip.copy(text)
        print("\n✅ Text copied to clipboard.")
        
    except Exception as e:
        print(f"❌ ERROR during transcription: {e}")
        return 1
    finally:
        # Clean up
        if config.clean_temp and os.path.exists(audio_path):
            try:
                os.remove(audio_path)
                print(f"Temporary audio file deleted: {audio_path}")
            except Exception as e:
                print(f"Note: Could not delete temporary file: {str(e)}")
        
        recorder.cleanup_ffmpeg_copy()
    
    return 0

def transcribe_file(file_path: str, config: Optional[TranscriptionConfig] = None) -> Optional[str]:
    """Transcribe a single audio file and return the text."""
    if config is None:
        config = TranscriptionConfig()
    
    # Initialize recorder (for dependency checking)
    recorder = AudioRecorder(config)
    
    # Check dependencies
    if not recorder.check_dependencies():
        raise RuntimeError("Required dependencies not available.")
    
    # Initialize transcriber and load model
    transcriber = Transcriber(config)
    transcriber.load_model(recorder)
    
    try:
        # Transcribe
        text = transcriber.transcribe_audio(file_path)
        return text
    finally:
        recorder.cleanup_ffmpeg_copy()

if __name__ == "__main__":
    # Run in console mode when executed directly
    import argparse
    
    parser = argparse.ArgumentParser(description="Voice recording and transcription tool")
    parser.add_argument("--file", type=str, help="Transcribe a specific audio file")
    parser.add_argument("--language", type=str, choices=["Spanish", "English"], 
                       default="Spanish", help="Language for transcription")
    parser.add_argument("--model", type=str, default="large",
                       choices=["tiny", "base", "small", "medium", "large"],
                       help="Whisper model size")
    args = parser.parse_args()
    
    if args.file:
        # File mode
        config = TranscriptionConfig(
            language=args.language,
            translate=(args.language == "Spanish"),
            model_size=args.model
        )
        try:
            text = transcribe_file(args.file, config)
            print("\n=== TRANSCRIPTION ===\n")
            print(text)
            pyperclip.copy(text)
            print("\n✅ Text copied to clipboard.")
        except Exception as e:
            print(f"❌ ERROR: {e}")
            sys.exit(1)
    else:
        # Interactive recording mode
        try:
            exit_code = run_console_transcription()
            sys.exit(exit_code)
        except KeyboardInterrupt:
            print("\n⚠️  Interrupted by user")
            sys.exit(1)