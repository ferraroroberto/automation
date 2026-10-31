"""
Shared FFmpeg / ffprobe helpers: GPU detection, video duration lookup and the
progress-reporting ffmpeg runner.

check_gpu_available(), get_video_duration() and run_ffmpeg_with_progress() are
the single source for video_reencoder.py, video_trim.py and
video_concatenator.py. get_video_duration() narrows
its exception list to subprocess.CalledProcessError / FileNotFoundError /
ValueError rather than catching bare `Exception`, so an unexpected bug
surfaces instead of being silently swallowed as "no duration".
"""

import re
import subprocess
import sys
from pathlib import Path
from typing import Callable, List, Optional

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


# Accepts HH:MM:SS.cc, MM:SS.cc and SS.cc, the forms ffmpeg prints in `time=`.
_TIME_PATTERN = re.compile(
    r"time=(\d+):(\d+):(\d+)\.(\d+)|time=(\d+):(\d+)\.(\d+)|time=(\d+)\.(\d+)"
)


def parse_ffmpeg_time(line: str) -> Optional[float]:
    """Return the elapsed seconds in an ffmpeg `time=` progress line, or None."""
    m = _TIME_PATTERN.search(line)
    if not m:
        return None
    if m.group(1) is not None:
        h, mn, s, cs = (int(m.group(i)) for i in (1, 2, 3, 4))
        return h * 3600 + mn * 60 + s + cs / 100.0
    if m.group(5) is not None:
        mn, s, cs = (int(m.group(i)) for i in (5, 6, 7))
        return mn * 60 + s + cs / 100.0
    s, cs = int(m.group(8)), int(m.group(9))
    return s + cs / 100.0


def run_ffmpeg_with_progress(
    cmd: List[str],
    total_seconds: Optional[float],
    progress_callback: Optional[Callable[[float], None]] = None,
    max_percent: float = 100.0,
) -> int:
    """Run ffmpeg, reporting percent done from its stderr `time=` lines; returns the exit code.

    `max_percent` caps the reported value so a caller that reports its own final
    100% never shows it early. Without a probed duration the percent falls back
    to a rough two percent per elapsed second.
    """
    process = subprocess.Popen(
        cmd,
        stderr=subprocess.PIPE,
        universal_newlines=True,
        bufsize=1,
        creationflags=NO_WINDOW,
    )
    for line in process.stderr:
        if not progress_callback or "time=" not in line:
            continue
        elapsed = parse_ffmpeg_time(line)
        if elapsed is None:
            continue
        percent = 100.0 * elapsed / total_seconds if total_seconds else elapsed * 2
        progress_callback(min(max_percent, percent))
    process.wait()
    return process.returncode
