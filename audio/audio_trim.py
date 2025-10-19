import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import subprocess
import re
import logging

# Configure module-level logger
logger = logging.getLogger(__name__)

# -------------------------------------------------------
# Function: choose_file
# Purpose:  Open a file dialog to select an audio file
# Returns:  The full path of the selected file
# -------------------------------------------------------
def choose_file() -> str:
    root = tk.Tk()
    root.withdraw()  # Hide the main Tk window
    file_path = filedialog.askopenfilename(
        title="Select audio file",
        filetypes=[("Audio files", "*.mp3 *.wav *.flac *.aac *.ogg *.m4a *.wma *.aiff")]  # Major audio formats
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        exit()
    return file_path

# -------------------------------------------------------
# Function: ask_time
# Purpose:  Ask the user for a time string in mm:ss or hh:mm:ss format
# Arguments:
#   - prompt (str): the message to show in the input dialog
# Returns:  A validated string in mm:ss or hh:mm:ss format
# -------------------------------------------------------
def ask_time(prompt: str) -> str:
    while True:
        input_time = simpledialog.askstring("Input", prompt + " (mm:ss or hh:mm:ss)")
        if input_time is None:
            messagebox.showerror("Error", "Input canceled.")
            exit()
        if re.match(r"^(?:\d{1,2}:\d{2}|\d{1,2}:\d{2}:\d{2})$", input_time):
            return input_time
        else:
            messagebox.showerror("Error", "Invalid format. Use mm:ss or hh:mm:ss.")

# -------------------------------------------------------
# Function: time_to_seconds
# Purpose:  Convert time from mm:ss or hh:mm:ss format to seconds
# Arguments:
#   - time_str (str): time string in mm:ss or hh:mm:ss format
# Returns:  Integer number of seconds
# -------------------------------------------------------
def time_to_seconds(time_str: str) -> int:
    parts = time_str.split(":")
    if len(parts) == 2:
        # mm:ss format
        minutes, seconds = map(int, parts)
        return minutes * 60 + seconds
    elif len(parts) == 3:
        # hh:mm:ss format
        hours, minutes, seconds = map(int, parts)
        return hours * 3600 + minutes * 60 + seconds
    else:
        raise ValueError(f"Invalid time format: {time_str}")

# -------------------------------------------------------
# Function: trim_audio
# Purpose:  Trim the selected audio between start and end time,
#           normalize its volume, and re-encode it for compatibility
# Arguments:
#   - input_path (str): path to the original audio file
#   - start_time (str): start time in mm:ss or hh:mm:ss
#   - end_time (str): end time in mm:ss or hh:mm:ss
# Behavior:
#   - Validates time input
#   - Trims audio using ffmpeg
#   - Applies AAC encoding and audio volume normalization using EBU R128 standard
#   - Saves output with a descriptive filename
# -------------------------------------------------------
def trim_audio(input_path: str, start_time: str, end_time: str) -> None:
    input_dir = os.path.dirname(input_path)
    filename = os.path.basename(input_path)
    name, ext = os.path.splitext(filename)

    # Convert time format → filename format
    start_label = start_time.replace(":", "-")
    end_label = end_time.replace(":", "-")
    output_file = os.path.join(input_dir, f"{name}_trim_{start_label}_{end_label}.m4a")

    # Convert times to seconds
    start_sec = time_to_seconds(start_time)
    end_sec = time_to_seconds(end_time)
    duration = end_sec - start_sec

    if duration <= 0:
        messagebox.showerror("Error", "End time must be after start time.")
        return

    try:
        logger.info("🔄 Starting audio trimming process...")

        # FFmpeg command for audio trimming and normalization
        ffmpeg_cmd = [
            "ffmpeg",
            "-y",                          # Overwrite output file without asking
            "-ss", str(start_sec),        # Start time
            "-i", input_path,             # Input audio
            "-t", str(duration),          # Duration to keep
            "-c:a", "aac",                # AAC audio codec for compatibility
            "-b:a", "192k",               # Audio bitrate
            "-af", "loudnorm",            # Audio filter: EBU R128 normalization
            "-movflags", "+faststart",   # Allows faster streaming and better compatibility
            output_file
        ]

        subprocess.run(ffmpeg_cmd, check=True)

        logger.info("✅ Audio trimming completed successfully")
        messagebox.showinfo("Success",
            f"Trimmed audio saved as:\n{output_file}\n\n"
            f"Duration: {duration} seconds")

    except subprocess.CalledProcessError as e:
        logger.error(f"❌ FFmpeg failed: {e}")
        messagebox.showerror("Error", f"FFmpeg failed:\n{e}")

# -------------------------------------------------------
# Main program entry point
# -------------------------------------------------------
if __name__ == "__main__":
    # Set up basic logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    file = choose_file()
    start = ask_time("Enter START time")
    end = ask_time("Enter END time")
    trim_audio(file, start, end)
