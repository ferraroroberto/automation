# source chatGPT > https://chatgpt.com/c/67a4e5d8-22d8-8009-b78b-3d3d7697e06a

import logging
import os
import tkinter as tk
from tkinter import filedialog
from pydub import AudioSegment

log = logging.getLogger(__name__)

# Manually set the correct paths for ffmpeg and ffprobe
ffmpeg_path = r"C:\Users\rober\AppData\Roaming\Python\Python311\site-packages\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v7.1.exe"
ffprobe_path = ffmpeg_path  # `imageio_ffmpeg` does not have a separate ffprobe, so use the same path

AudioSegment.converter = ffmpeg_path
AudioSegment.ffprobe = ffprobe_path

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
