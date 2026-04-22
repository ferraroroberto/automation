"""
Core module for audio recording and transcription functionality.
This module can be used standalone from the command line or imported by the GUI module.
"""

# Standard library imports
import json
import logging
import os
import platform
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import warnings
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Tuple

# Third-party imports
import numpy as np
import pyperclip
import scipy.io.wavfile as wav
import sounddevice as sd
from pynput import keyboard

# Local imports
# import whisper  # Deferred until needed in load_model()

# Set up module-level logger
logger = logging.getLogger(__name__)

def get_machine_name() -> str:
    """Get the machine name (hostname) for device-specific configuration.

    Returns:
        str: The machine name in lowercase.
    """
    try:
        machine_name = platform.node().lower()
        logger.debug(f"🖥️ Detected machine name: {machine_name}")
        return machine_name
    except Exception as e:
        logger.warning(f"⚠️ Could not determine machine name: {e}")
        return "unknown"

# Known CUDA architectures that official PyTorch GPU wheels currently support
TORCH_KNOWN_COMPATIBLE_ARCHES = {
    "sm_50",
    "sm_60",
    "sm_61",
    "sm_70",
    "sm_75",
    "sm_80",
    "sm_86",
    "sm_90",
}

TORCH_CUDA_INDEX_URL = "https://download.pytorch.org/whl/cu121"


def build_torch_install_command() -> List[str]:
    """Return the recommended pip command for installing CUDA-enabled PyTorch."""
    return [
        sys.executable,
        "-m",
        "pip",
        "install",
        "torch",
        "torchvision",
        "torchaudio",
        "--index-url",
        TORCH_CUDA_INDEX_URL,
    ]

# Custom exceptions
class TranscriptionError(Exception):
    """Base exception for transcription-related errors."""
    pass

class DependencyError(TranscriptionError):
    """Raised when required dependencies are not available."""
    pass

class AudioDeviceError(TranscriptionError):
    """Raised when audio device issues occur."""
    pass

class ConfigurationError(TranscriptionError):
    """Raised when configuration issues occur."""
    pass

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
    preferred_mics: Optional[List[str]] = None
    machine_specific_mics: Optional[Dict[str, List[str]]] = None
    model_size: str = "small"
    clean_temp: bool = True
    ffmpeg_path: Optional[str] = None
    log_level: str = "INFO"

    def __post_init__(self) -> None:
        """Initialize default values for optional fields."""
        # Set machine-specific microphones if available
        if self.preferred_mics is None:
            machine_name = get_machine_name()
            if self.machine_specific_mics and machine_name in self.machine_specific_mics:
                self.preferred_mics = self.machine_specific_mics[machine_name]
                logger.info(f"🎙️ Using machine-specific microphones for '{machine_name}': {self.preferred_mics}")
            else:
                # Default fallback microphones
                self.preferred_mics = [
                    "el gato wave XLR (Elgato Wave XLR)",
                    "Wave Link Stream (Elgato Wave:XLR)"
                ]
                logger.info(f"🎙️ Using default microphones (machine: '{machine_name}')")

    @classmethod
    def from_json(cls, config_path: Optional[str] = None) -> "TranscriptionConfig":
        """Load configuration from JSON file.

        Args:
            config_path: Path to config file. If None, uses default location.

        Returns:
            TranscriptionConfig: Loaded configuration instance.
        """
        if config_path is None:
            # Default config path relative to this module
            module_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(module_dir, "config.json")

        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            # Validate required fields
            cls._validate_config(config_data)

            # Convert preferred_mics to list if it's not None
            if 'preferred_mics' in config_data and config_data['preferred_mics'] is not None:
                config_data['preferred_mics'] = list(config_data['preferred_mics'])

            return cls(**config_data)

        except FileNotFoundError:
            logger.warning(f"📂 Config file not found at {config_path}, using defaults")
            return cls()
        except json.JSONDecodeError as e:
            logger.error(f"❌ Invalid JSON in config file: {e}")
            return cls()
        except Exception as e:
            logger.error(f"❌ Error loading config: {e}")
            return cls()

    @staticmethod
    def _validate_config(config_data: Dict[str, Any]) -> None:
        """Validate configuration data.

        Args:
            config_data: Configuration dictionary to validate.

        Raises:
            ValueError: If configuration is invalid.
        """
        # Validate record_seconds
        if 'record_seconds' in config_data:
            if not isinstance(config_data['record_seconds'], int) or config_data['record_seconds'] <= 0:
                raise ValueError("record_seconds must be a positive integer")

        # Validate language
        if 'language' in config_data:
            valid_languages = ["Spanish", "English"]
            if config_data['language'] not in valid_languages:
                raise ValueError(f"language must be one of: {valid_languages}")

        # Validate model_size
        if 'model_size' in config_data:
            valid_models = ["tiny", "base", "small", "medium", "large"]
            if config_data['model_size'] not in valid_models:
                raise ValueError(f"model_size must be one of: {valid_models}")

        # Validate log_level
        if 'log_level' in config_data:
            valid_levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
            if config_data['log_level'] not in valid_levels:
                raise ValueError(f"log_level must be one of: {valid_levels}")

    def save_to_json(self, config_path: Optional[str] = None) -> None:
        """Save current configuration to JSON file.

        Args:
            config_path: Path to save config file. If None, uses default location.
        """
        if config_path is None:
            module_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(module_dir, "config.json")

        config_data = {
            "record_seconds": self.record_seconds,
            "language": self.language,
            "translate": self.translate,
            "preferred_mics": self.preferred_mics,
            "machine_specific_mics": self.machine_specific_mics,
            "model_size": self.model_size,
            "clean_temp": self.clean_temp,
            "ffmpeg_path": self.ffmpeg_path,
            "log_level": self.log_level
        }

        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            logger.info(f"🔧 Configuration saved to {config_path}")
        except Exception as e:
            logger.error(f"❌ Failed to save configuration: {e}")

class AudioRecorder:
    """Handles audio recording functionality."""

    def __init__(self, config: TranscriptionConfig) -> None:
        """Initialize the audio recorder with configuration."""
        self.config = config
        self.audio_buffer: List[float] = []
        self.stop_recording: bool = False
        self.key_pressed: bool = False
        self.recording: Optional[np.ndarray] = None
        self.start_time: Optional[float] = None
        self.gpu_available: Optional[bool] = None
        self.ffmpeg_path: Optional[str] = None
        self.ffmpeg_copy_path: Optional[str] = None
        self.ffmpeg_already_cleaned: bool = False
        self.dependencies_checked: bool = False  # Cache dependency check results
        self.audio_devices_cached: Optional[List[Tuple[int, dict]]] = None   # Cache audio devices
        self.install_prompt_shown: bool = False

    @staticmethod
    def _format_command_for_logging(command: List[str]) -> str:
        """Return a shell-friendly representation of a command list."""
        formatted_parts = []
        for part in command:
            if " " in part or "\t" in part:
                formatted_parts.append(f'"{part}"')
            else:
                formatted_parts.append(part)
        return " ".join(formatted_parts)

    @staticmethod
    def _query_nvidia_smi_capability() -> Optional[str]:
        """Attempt to read the GPU compute capability via nvidia-smi."""
        nvidia_smi = shutil.which("nvidia-smi")
        if not nvidia_smi:
            return None

        try:
            result = subprocess.run(
                [nvidia_smi, "--query-gpu=compute_cap", "--format=csv,noheader"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            if result.returncode != 0:
                return None
            raw_value = result.stdout.strip().splitlines()[0].strip()
            major, minor = raw_value.split(".")
            return f"sm_{int(major)}{int(minor)}"
        except Exception as exc:
            logger.debug(f"⚠️  Could not determine compute capability via nvidia-smi: {exc}")
            return None

    def _has_known_compatible_build(self, capability: Optional[str]) -> bool:
        """Return True when we know a published wheel supports the GPU capability."""
        if not capability:
            return False
        return capability in TORCH_KNOWN_COMPATIBLE_ARCHES

    def _maybe_prompt_torch_install(self, reason: str) -> None:
        """Optionally prompt the user to install a GPU-enabled PyTorch build."""
        if self.install_prompt_shown:
            return

        self.install_prompt_shown = True
        install_command = build_torch_install_command()
        command_display = self._format_command_for_logging(install_command)
        logger.warning(reason)
        logger.info("💡 A compatible CUDA build exists. Run this to install it:")
        logger.info(command_display)

        if not sys.stdin or not sys.stdin.isatty():
            logger.info("Skipping interactive prompt (non-interactive session).")
            return

        choice = input("Install GPU-enabled PyTorch now? [y/N]: ").strip().lower()
        if choice not in {"y", "yes"}:
            logger.info("Keeping current installation. Continuing with CPU execution.")
            return

        try:
            subprocess.check_call(install_command)
            logger.info("✅ PyTorch installation finished. Restart the app to enable GPU acceleration.")
        except subprocess.CalledProcessError as exc:
            logger.error(f"❌ PyTorch installation failed: {exc}")
            logger.info("Continuing with CPU mode for this session.")

    def _handle_gpu_incompatibility(
        self,
        capability: Optional[str],
        compatible_build_available: bool,
        reason: str,
    ) -> None:
        """Log guidance when CUDA execution is not possible."""
        if compatible_build_available:
            self._maybe_prompt_torch_install(reason)
        else:
            if capability:
                logger.warning(
                    f"⚠️  GPU capability {capability} is newer than available PyTorch wheels. "
                    "Falling back to CPU."
                )
            else:
                logger.warning("⚠️  Could not determine a compatible GPU build. Falling back to CPU.")
        
    def check_gpu_availability(self) -> bool:
        """Check if a CUDA GPU is available. Perform this check lazily."""
        if self.gpu_available is not None:
            return self.gpu_available

        try:
            import torch
        except ImportError:
            logger.warning("⚠️  PyTorch not found. Assuming CPU mode.")
            self.gpu_available = False
            return self.gpu_available

        capability_str: Optional[str] = None
        compatibility_reason: Optional[str] = None
        build_arches: List[str] = []
        cpu_only_build = False

        if hasattr(torch.cuda, "get_arch_list"):
            try:
                build_arches = [arch.strip() for arch in torch.cuda.get_arch_list()]
            except Exception as exc:
                logger.debug(f"⚠️  Could not read compiled CUDA architectures: {exc}")

        if torch.cuda.is_available():
            try:
                gpu_count = torch.cuda.device_count()
                gpu_name = torch.cuda.get_device_name(0) if gpu_count > 0 else "Unknown"
                major, minor = torch.cuda.get_device_capability(0)
                capability_str = f"sm_{major}{minor}"
                logger.info(
                    f"✅ CUDA GPU detected: {gpu_name} (Count: {gpu_count}, Capability: {capability_str})"
                )
            except Exception as exc:
                compatibility_reason = f"Unable to inspect GPU details: {exc}"

            # Check for direct compatibility or known compatible architectures
            is_compatible = False
            if capability_str:
                if not build_arches or capability_str in build_arches:
                    is_compatible = True
                else:
                    # Check for known compatible architectures
                    # sm_89 (RTX 40 series Ada) is compatible with sm_86/sm_90 kernels
                    compatibility_map = {
                        'sm_89': ['sm_86', 'sm_90'],  # RTX 40 Ada can use sm_86/sm_90 kernels
                    }
                    compatible_arches = compatibility_map.get(capability_str, [])
                    if any(arch in build_arches for arch in compatible_arches):
                        is_compatible = True
                        logger.info(f"✅ GPU {capability_str} compatible with available kernels")

            if is_compatible:
                self.gpu_available = True
                return True

            if capability_str and build_arches and not is_compatible:
                supported_arches = ", ".join(build_arches) if build_arches else "unknown"
                compatibility_reason = (
                    f"GPU capability {capability_str} is not included in the current PyTorch build "
                    f"(supports: {supported_arches})."
                )
        else:
            torch_cuda_version = getattr(torch.version, "cuda", None)
            if torch_cuda_version is None:
                cpu_only_build = True
                compatibility_reason = "Current PyTorch build is CPU-only."
            else:
                compatibility_reason = (
                    "PyTorch CUDA runtime is unavailable despite a CUDA build being installed."
                )

        if capability_str is None:
            capability_str = self._query_nvidia_smi_capability()

        compatible_build_available = self._has_known_compatible_build(capability_str)
        if not compatible_build_available and cpu_only_build:
            # We could not determine the GPU capability, but a compatible build likely exists.
            compatible_build_available = True
        reason = compatibility_reason or "Unknown GPU compatibility issue."
        self._handle_gpu_incompatibility(capability_str, compatible_build_available, reason)
        self.gpu_available = False
        return self.gpu_available
    
    def check_dependencies(self) -> bool:
        """Check if all required dependencies are available.

        Returns:
            bool: True if all dependencies are available, False otherwise.

        Raises:
            DependencyError: If critical dependencies are missing.
        """
        # Return cached result if already checked
        if self.dependencies_checked:
            return True

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
                logger.info(f"✅ FFmpeg found in system PATH: {ffmpeg_path}")
        except (FileNotFoundError, subprocess.TimeoutExpired, subprocess.SubprocessError):
            pass

        if not ffmpeg_found:
            logger.error("❌ FFmpeg executable not found in system PATH.")
            logger.info("Trying alternative solutions...")

            # Try to use imageio-ffmpeg
            try:
                import imageio_ffmpeg
                ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
                logger.info(f"✅ FFmpeg binary found via imageio-ffmpeg: {ffmpeg_path}")

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
                            logger.info(f"✅ Added FFmpeg directory to PATH: {ffmpeg_dir}")

                        # Set environment variables
                        os.environ['FFMPEG_BINARY'] = ffmpeg_path
                        os.environ['IMAGEIO_FFMPEG_EXE'] = ffmpeg_path

                        # --- BEGIN: Copy ffmpeg.exe to script directory if not present ---
                        script_dir = os.path.dirname(os.path.abspath(__file__))
                        ffmpeg_copy_path = os.path.join(script_dir, 'ffmpeg.exe')
                        if os.name == 'nt' and not os.path.exists(ffmpeg_copy_path):
                            try:
                                shutil.copy2(ffmpeg_path, ffmpeg_copy_path)
                                logger.info(f"✅ Copied FFmpeg to script directory: {ffmpeg_copy_path}")
                                self.ffmpeg_copy_path = ffmpeg_copy_path
                            except Exception as e:
                                logger.warning(f"⚠️  Could not copy FFmpeg to script directory: {e}")
                        # --- END: Copy ffmpeg.exe to script directory if not present ---

                except Exception as e:
                    logger.warning(f"⚠️  Could not test FFmpeg binary: {e}")

            except ImportError:
                logger.info("📦 Installing imageio-ffmpeg...")
                try:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'imageio-ffmpeg'])
                    logger.info("✅ imageio-ffmpeg installed successfully.")
                    return self.check_dependencies()  # Retry after installation
                except subprocess.CalledProcessError:
                    error_msg = "❌ ERROR: Could not install imageio-ffmpeg"
                    logger.error(error_msg)
                    raise DependencyError(error_msg)

        if not ffmpeg_found:
            error_msg = "FFmpeg is required but not available"
            logger.error(f"❌ {error_msg}")
            raise DependencyError(error_msg)

        if ffmpeg_path:
            self.ffmpeg_path = ffmpeg_path

        # Cache the result
        self.dependencies_checked = True
        return ffmpeg_found
    
    def get_audio_devices(self) -> List[Tuple[int, dict]]:
        """Get list of available audio input devices.

        Returns:
            List[Tuple[int, dict]]: List of tuples containing device index and device info.
        """
        # Return cached result if available
        if self.audio_devices_cached is not None:
            return self.audio_devices_cached

        devices = sd.query_devices()
        input_devices: List[Tuple[int, dict]] = []

        for i, device in enumerate(devices):
            if device['max_input_channels'] > 0:
                input_devices.append((i, device))

        # Cache the result
        self.audio_devices_cached = input_devices
        return input_devices

    def select_microphone(self, input_devices: List[Tuple[int, dict]],
                         verbose: bool = True) -> Optional[int]:
        """Select the best available microphone based on preferences.

        Args:
            input_devices: List of available input devices.
            verbose: Whether to print selection information.

        Returns:
            Optional[int]: Device index of selected microphone, or None if none found.
        """
        if not input_devices:
            return None
            
        # Try to find preferred microphone
        for preferred_mic in self.config.preferred_mics:
            for i, device in input_devices:
                if preferred_mic.lower() in device['name'].lower():
                    if verbose:
                        logger.info(f"🎙️ Found preferred microphone: {device['name']}")
                    return i

        # Use first available device as fallback
        selected_device = input_devices[0][0]
        if verbose:
            device_info = sd.query_devices(selected_device)
            logger.info(f"🎙️ No preferred microphone found. Using: {device_info['name']}")

        return selected_device
    
    def monitor_audio_levels(self, device: int, fs: int,
                           progress_callback: Optional[Callable[[int, float], None]] = None) -> None:
        """Monitor audio levels during recording.

        Args:
            device: Audio device index to monitor.
            fs: Sample rate.
            progress_callback: Optional callback function for progress updates.
        """
        def callback(indata: np.ndarray, frames: int, time_info, status) -> None:
            """Audio stream callback function."""
            if status:
                logger.debug(f"🎤 Audio stream status: {status}")
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
    
    def _show_terminal_progress(self, remaining: int, level: float) -> None:
        """Show progress in terminal mode.

        Args:
            remaining: Seconds remaining in recording.
            level: Current audio level (0.0 to 1.0).
        """
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
                    progress_callback: Optional[Callable[[int, float], None]] = None) -> Optional[np.ndarray]:
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
                logger.warning(f"⚠️  Very low audio levels detected (max: {max_level:.4f})")
            else:
                logger.info(f"🎤 Peak audio level: {max_level:.4f}")

            # Convert to int16
            return (self.recording * 32767).astype(np.int16)
        
        return None
    
    def cleanup_ffmpeg_copy(self) -> None:
        """Clean up the temporary FFmpeg copy if it exists."""
        if self.ffmpeg_already_cleaned:
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_exe_path = os.path.join(script_dir, 'ffmpeg.exe')

        target_path = self.ffmpeg_copy_path if self.ffmpeg_copy_path else ffmpeg_exe_path

        if os.path.exists(target_path):
            try:
                os.remove(target_path)
                logger.info(f"✅ Cleaned up FFmpeg copy: {target_path}")
                self.ffmpeg_already_cleaned = True
            except Exception as e:
                logger.warning(f"⚠️  Could not remove FFmpeg copy: {e}")
        else:
            self.ffmpeg_already_cleaned = True

class Transcriber:
    """Handles transcription functionality using Whisper."""

    def __init__(self, config: TranscriptionConfig) -> None:
        """Initialize the transcriber with configuration.

        Args:
            config: Transcription configuration settings.
        """
        self.config = config
        self.model = None
        self.device: Optional[str] = None
        
    def load_model(self, recorder: AudioRecorder, progress_callback: Optional[Callable[[str], None]] = None) -> None:
        """Load the Whisper model.

        Args:
            recorder: AudioRecorder instance for dependency checking.
            progress_callback: Optional callback for progress updates.
        """
        # Configure FFmpeg paths
        try:
            import imageio_ffmpeg
            ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
            os.environ['IMAGEIO_FFMPEG_EXE'] = ffmpeg_path
            os.environ['FFMPEG_BINARY'] = ffmpeg_path
        except ImportError:
            pass

        # Determine device
        self.device = "cuda" if recorder.check_gpu_availability() else "cpu"
        if self.device == "cuda":
            logger.info("🚀 Using GPU acceleration (CUDA)")
        else:
            logger.warning("⚠️  Using CPU (will use FP32 precision)")

        # Load model
        if progress_callback:
            progress_callback("Initializing model download...")
        logger.info(f"🤖 Downloading/loading Whisper model '{self.config.model_size}'...")

        # Import whisper here (deferred import for faster startup)
        import whisper
        self.model = whisper.load_model(self.config.model_size, device=self.device)
        if progress_callback:
            progress_callback("Model loaded successfully!")

    def transcribe_audio(self, audio_path: str) -> str:
        """Transcribe audio file and return text."""
        if not self.model:
            raise ValueError("Model not loaded. Call load_model() first.")

        logger.info("🎤 Transcribing audio...")

        # --- Ensure ffmpeg.exe in script dir is in PATH (for Windows) ---
        script_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_exe_path = os.path.join(script_dir, 'ffmpeg.exe')
        if os.name == 'nt' and os.path.exists(ffmpeg_exe_path):
            current_path = os.environ.get('PATH', '')
            if script_dir not in current_path.split(os.pathsep):
                os.environ['PATH'] = script_dir + os.pathsep + current_path
                logger.info(f"🔧 Added script directory to PATH for FFmpeg: {script_dir}")
        # ---------------------------------------------------------------

        # Set transcription options
        transcribe_options = {
            "language": self.config.language,
            "task": "translate" if self.config.translate else "transcribe"
        }

        # Use fp16 only if GPU is available
        if self.device == "cuda":
            transcribe_options["fp16"] = True
            logger.info("🚀 Using FP16 precision for faster GPU processing")
        else:
            transcribe_options["fp16"] = False
            logger.info("⚙️ Using FP32 precision (CPU mode)")
        
        # Transcribe with warnings suppressed
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            result = self.model.transcribe(audio_path, **transcribe_options)
        
        return result["text"]

def prompt_for_language() -> Tuple[str, bool]:
    """Prompt the user to select a language in console mode.

    Returns:
        Tuple[str, bool]: Language name and whether to translate.
    """
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
        logger.info("✅ Selected: Spanish (with translation)")
    elif choice in ["2", "en", "english", "e"]:
        language = "English"
        translate = False
        logger.info("✅ Selected: English (transcription only)")
    else:
        logger.warning("⚠️ Invalid choice. Using default: Spanish (with translation)")
        language = "Spanish"
        translate = True

    print()  # Keep this for user input formatting
    return language, translate

def run_console_transcription(config: Optional[TranscriptionConfig] = None) -> int:
    if config is None:
        config = TranscriptionConfig.from_json()
    
    # Prompt for language
    config.language, config.translate = prompt_for_language()
    
    # Initialize recorder
    recorder = AudioRecorder(config)
    
    # Check dependencies
    try:
        recorder.check_dependencies()
    except DependencyError as e:
        logger.error(f"❌ ERROR: {e}")
        return 1

    # Get and display audio devices
    print("\n=== AVAILABLE MICROPHONES ===\n")
    input_devices = recorder.get_audio_devices()

    for i, device in input_devices:
        print(f"{i}: {device['name']}")

    # Select microphone
    selected_device = recorder.select_microphone(input_devices)
    if selected_device is None:
        error_msg = "No suitable microphone found! Cannot continue."
        logger.error(f"❌ {error_msg}")
        raise AudioDeviceError(error_msg)

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
        logger.warning("🎤 Recording too short, nothing to transcribe.")
        recorder.cleanup_ffmpeg_copy()
        return 0

    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmpfile:
        wav.write(tmpfile.name, 16000, recording)
        audio_path = tmpfile.name
        logger.debug(f"📁 Audio saved temporarily to: {tmpfile.name}")

    # Initialize transcriber and load model
    transcriber = Transcriber(config)
    logger.info("🤖 Loading Whisper model...")
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
        logger.error(f"❌ ERROR during transcription: {e}")
        return 1
    finally:
        # Clean up
        if config.clean_temp and os.path.exists(audio_path):
            try:
                os.remove(audio_path)
                logger.debug(f"🗑️ Temporary audio file deleted: {audio_path}")
            except Exception as e:
                logger.warning(f"⚠️ Could not delete temporary file: {str(e)}")

        recorder.cleanup_ffmpeg_copy()
    
    return 0

def transcribe_file(file_path: str, config: Optional[TranscriptionConfig] = None) -> Optional[str]:
    """Transcribe a single audio file and return the text.

    Args:
        file_path: Path to the audio file to transcribe.
        config: Optional transcription configuration.

    Returns:
        Optional[str]: Transcribed text, or None if transcription failed.
    """
    if config is None:
        config = TranscriptionConfig.from_json()
    
    # Initialize recorder (for dependency checking)
    recorder = AudioRecorder(config)
    
    # Check dependencies
    try:
        recorder.check_dependencies()
    except DependencyError as e:
        raise TranscriptionError(f"Dependency check failed: {e}") from e
    
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
    parser.add_argument("--model", type=str, default="small",
                       choices=["tiny", "base", "small", "medium", "large"],
                       help="Whisper model size")
    args = parser.parse_args()
    
    if args.file:
        # File mode
        config = TranscriptionConfig.from_json()
        # Override with command line arguments
        if args.language:
            config.language = args.language
            config.translate = (args.language == "Spanish")
        if args.model:
            config.model_size = args.model
        try:
            text = transcribe_file(args.file, config)
            print("\n=== TRANSCRIPTION ===\n")
            print(text)
            pyperclip.copy(text)
            print("\n✅ Text copied to clipboard.")
        except TranscriptionError as e:
            logger.error(f"❌ Transcription failed: {e}")
            sys.exit(1)
        except Exception as e:
            logger.error(f"❌ Unexpected error: {e}")
            sys.exit(1)
    else:
        # Interactive recording mode
        try:
            config = TranscriptionConfig.from_json()
            exit_code = run_console_transcription(config)
            sys.exit(exit_code)
        except TranscriptionError as e:
            logger.error(f"❌ Transcription failed: {e}")
            sys.exit(1)
        except KeyboardInterrupt:
            logger.warning("⚠️  Interrupted by user")
            sys.exit(1)
        except Exception as e:
            logger.error(f"❌ Unexpected error: {e}")
            sys.exit(1)