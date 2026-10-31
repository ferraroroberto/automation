import logging
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from typing import List

import imageio_ffmpeg

# Repo-root shared helper (scripts run from their own folder, so add the root).
sys.path.append(str(Path(__file__).resolve().parents[1]))
from no_window import NO_WINDOW  # noqa: E402

log = logging.getLogger(__name__)


def build_convert_command(ffmpeg_exe: str, ogg_path: str, mp3_path: str) -> List[str]:
    """ffmpeg command that transcodes ``ogg_path`` to a 192k MP3 at ``mp3_path``."""
    return [ffmpeg_exe, "-y", "-i", ogg_path, "-vn", "-b:a", "192k", mp3_path]


def convert_ogg_to_mp3():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select an OGG file", filetypes=[("OGG files", "*.ogg")])

    if not file_path:
        log.info("No file selected.")
        return

    output_file = os.path.splitext(file_path)[0] + ".mp3"

    # imageio_ffmpeg resolves the bundled ffmpeg at runtime, so no ffmpeg or
    # ffprobe has to be on PATH; decoding goes straight through ffmpeg.
    command = build_convert_command(imageio_ffmpeg.get_ffmpeg_exe(), file_path, output_file)
    result = subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", errors="replace", creationflags=NO_WINDOW
    )
    if result.returncode == 0:
        log.info("Conversion successful! Saved as: %s", output_file)
    else:
        log.error("Error during conversion: %s", result.stderr.strip()[-500:])

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    convert_ogg_to_mp3()
