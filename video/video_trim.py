import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import subprocess
import re

# -------------------------------------------------------
# Function: choose_file
# Purpose:  Open a file dialog to select a video file
# Returns:  The full path of the selected file
# -------------------------------------------------------
def choose_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main Tk window
    file_path = filedialog.askopenfilename(
        title="Select video file",
        filetypes=[("Video files", "*.mp4 *.mov *.mkv *.webm *.avi")]  # Supported formats
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        exit()
    return file_path

# -------------------------------------------------------
# Function: ask_time
# Purpose:  Ask the user for a time string in mm:ss format
# Arguments:
#   - prompt (str): the message to show in the input dialog
# Returns:  A validated string in mm:ss format
# -------------------------------------------------------
def ask_time(prompt):
    while True:
        input_time = simpledialog.askstring("Input", prompt + " (mm:ss)")
        if input_time is None:
            messagebox.showerror("Error", "Input canceled.")
            exit()
        if re.match(r"^\d{1,2}:\d{2}$", input_time):
            return input_time
        else:
            messagebox.showerror("Error", "Invalid format. Use mm:ss.")

# -------------------------------------------------------
# Function: time_to_seconds
# Purpose:  Convert time from mm:ss format to seconds
# Arguments:
#   - time_str (str): time string in mm:ss format
# Returns:  Integer number of seconds
# -------------------------------------------------------
def time_to_seconds(time_str):
    minutes, seconds = map(int, time_str.split(":"))
    return minutes * 60 + seconds

# -------------------------------------------------------
# Function: trim_video
# Purpose:  Trim the selected video between start and end time,
#           normalize its audio, and re-encode it for compatibility
# Arguments:
#   - input_path (str): path to the original video file
#   - start_time (str): start time in mm:ss
#   - end_time (str): end time in mm:ss
# Behavior:
#   - Validates time input
#   - Trims video using ffmpeg
#   - Applies H.264 encoding and AAC audio
#   - Normalizes audio volume using EBU R128 standard
#   - Saves output with a descriptive filename
# -------------------------------------------------------
def trim_video(input_path, start_time, end_time):
    input_dir = os.path.dirname(input_path)
    filename = os.path.basename(input_path)
    name, ext = os.path.splitext(filename)

    # Convert mm:ss → mm-ss for the filename
    start_label = start_time.replace(":", "-")
    end_label = end_time.replace(":", "-")
    output_file = os.path.join(input_dir, f"{name}_trim_{start_label}_{end_label}{ext}")

    # Convert times to seconds
    start_sec = time_to_seconds(start_time)
    end_sec = time_to_seconds(end_time)
    duration = end_sec - start_sec

    if duration <= 0:
        messagebox.showerror("Error", "End time must be after start time.")
        return

    try:
        # ffmpeg command with re-encoding and audio normalization
        subprocess.run([
            "ffmpeg",
            "-y",                          # Overwrite output file without asking
            "-ss", str(start_sec),        # Start time (before input = faster)
            "-i", input_path,             # Input video
            "-t", str(duration),          # Duration to keep
            "-c:v", "libx264",            # H.264 video codec for compatibility
            "-preset", "fast",            # Speed/quality tradeoff
            "-crf", "23",                 # Constant Rate Factor (quality setting)
            "-c:a", "aac",                # AAC audio codec
            "-b:a", "192k",               # Audio bitrate
            "-af", "loudnorm",            # Audio filter: EBU R128 normalization
            "-movflags", "+faststart",   # Allows faster streaming and better compatibility
            output_file
        ], check=True)

        messagebox.showinfo("Success", f"Trimmed video saved as:\n{output_file}")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"FFmpeg failed:\n{e}")

# -------------------------------------------------------
# Main program entry point
# -------------------------------------------------------
if __name__ == "__main__":
    file = choose_file()
    start = ask_time("Enter START time")
    end = ask_time("Enter END time")
    trim_video(file, start, end)
