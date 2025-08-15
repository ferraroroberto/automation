# source chatGPT > https://chatgpt.com/c/67a4e5d8-22d8-8009-b78b-3d3d7697e06a (not working, but with command line yes)

import os
import tkinter as tk
from tkinter import filedialog
from pydub import AudioSegment

# Manually set the correct paths for ffmpeg and ffprobe
ffmpeg_path = r"C:\Users\rober\AppData\Roaming\Python\Python311\site-packages\imageio_ffmpeg\binaries\ffmpeg-win-x86_64-v7.1.exe"
ffprobe_path = ffmpeg_path  # `imageio_ffmpeg` does not have a separate ffprobe, so use the same path

AudioSegment.converter = ffmpeg_path
AudioSegment.ffprobe = ffprobe_path

def convert_ogg_to_mp3():
    # Open file picker
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select an OGG file", filetypes=[("OGG files", "*.ogg")])

    if not file_path:
        print("No file selected.")
        return

    # Get output file path
    output_file = os.path.splitext(file_path)[0] + ".mp3"

    # Convert OGG to MP3 using pydub
    try:
        audio = AudioSegment.from_file(file_path, format="ogg")  # Explicitly set format
        audio.export(output_file, format="mp3", bitrate="192k")
        print(f"Conversion successful! Saved as: {output_file}")
    except Exception as e:
        print("Error during conversion:", e)

if __name__ == "__main__":
    convert_ogg_to_mp3()
