import logging
import os
import tkinter as tk
from tkinter import filedialog
import imageio_ffmpeg
from pydub import AudioSegment

log = logging.getLogger(__name__)

# Resolve ffmpeg at runtime via imageio_ffmpeg so this works on any machine
# regardless of Python version, install location, or ffmpeg build.
_ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()
AudioSegment.converter = _ffmpeg_path
AudioSegment.ffprobe = _ffmpeg_path  # imageio_ffmpeg bundles no separate ffprobe

def convert_ogg_to_mp3():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select an OGG file", filetypes=[("OGG files", "*.ogg")])

    if not file_path:
        log.info("No file selected.")
        return

    output_file = os.path.splitext(file_path)[0] + ".mp3"

    try:
        audio = AudioSegment.from_file(file_path, format="ogg")
        audio.export(output_file, format="mp3", bitrate="192k")
        log.info("Conversion successful! Saved as: %s", output_file)
    except Exception as e:
        log.error("Error during conversion: %s", e)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    convert_ogg_to_mp3()
