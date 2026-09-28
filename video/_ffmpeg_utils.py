"""
Shared FFmpeg / ffprobe helpers: GPU detection and video duration lookup.

Consolidates check_gpu_available() and get_video_duration(), previously
duplicated verbatim across video_reencoder.py and video_trim.py (dedup:
audit issue #64). Uses the narrower, more precise exception list from
video_trim.py's get_video_duration() (subprocess.CalledProcessError /
FileNotFoundError / ValueError) rather than video_reencoder.py's bare
`except Exception`, so an unexpected bug surfaces instead of being
silently swallowed as "no duration".
"""

import subprocess
import sys
from pathlib import Path
from typing import Optional

# Suppress the console window each ffmpeg/ffprobe/nvidia-smi spawn would
# otherwise flash when these helpers run under a console-less GUI parent.
# Repo-root shared helper (scripts run from their own folder, so add the root).
sys.path.append(str(Path(__file__).resolve().parents[1]))
from no_window import NO_WINDOW  # noqa: E402


def check_gpu_available() -> bool:
    """Return True if an NVIDIA GPU with NVENC support is available for ffmpeg."""
    try:
        result = subprocess.run(["nvidia-smi"], capture_output=True, text=True, creationflags=NO_WINDOW)
        if result.returncode != 0:
            return False
        encoders_check = subprocess.run(
            ["ffmpeg", "-encoders"], capture_output=True, text=True, creationflags=NO_WINDOW
        )
        return "h264_nvenc" in encoders_check.stdout
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


def get_video_duration(input_path: str) -> Optional[float]:
    """Get video duration in seconds using ffprobe. Returns float or None on error."""
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error", "-show_entries",
                "format=duration", "-of",
                "default=noprint_wrappers=1:nokey=1", input_path
            ],
            capture_output=True, text=True, check=True, timeout=30,
            creationflags=NO_WINDOW,
        )
        return float(result.stdout.strip())
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return None
