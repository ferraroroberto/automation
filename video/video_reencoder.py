import os
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
import time

def choose_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select video file",
        filetypes=[("Video files", "*.mp4 *.mov *.mkv *.webm *.avi")]
    )
    root.destroy()
    if not file_path:
        print("Error: No file selected.")
        exit()
    return file_path

def get_file_info(file_path):
    size_bytes = os.path.getsize(file_path)
    size_mb = size_bytes / (1024 * 1024)
    ext = os.path.splitext(file_path)[1].lower().replace('.', '')
    return ext, size_mb

def ask_format(formats):
    print(f"Available formats: {', '.join(formats)}")
    while True:
        format_choice = input("Choose output format: ").strip().lower()
        if format_choice in formats:
            return format_choice
        print("Invalid format. Try again.")

def ask_target_size(current_size):
    print(f"Current file size: {current_size:.2f} MB")
    while True:
        size_str = input("Enter desired final size in MB: ").strip()
        try:
            size_mb = float(size_str)
            if size_mb > 0:
                return size_mb
        except ValueError:
            pass
        print("Invalid size. Try again.")

def reencode_video(input_path, output_format, target_size_mb):
    # Get video duration using ffprobe
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error", "-show_entries",
                "format=duration", "-of",
                "default=noprint_wrappers=1:nokey=1", input_path
            ],
            capture_output=True, text=True, check=True
        )
        duration = float(result.stdout.strip())
    except Exception:
        print("Could not determine video duration.")
        exit()

    # Calculate total bitrate needed (in bits/sec)
    total_bitrate = (target_size_mb * 8 * 1024 * 1024) / duration
    audio_bitrate = 128 * 1024
    video_bitrate = int(total_bitrate - audio_bitrate)
    if video_bitrate < 100_000:
        print("Target size too small for reasonable quality.")
        exit()

    input_dir = os.path.dirname(input_path)
    name, _ = os.path.splitext(os.path.basename(input_path))
    output_file = os.path.join(input_dir, f"{name}_reencoded.{output_format}")

    codecs = {
        "mp4":  ("h264_nvenc", "aac"),
        "mkv":  ("h264_nvenc", "aac"),
        "mov":  ("h264_nvenc", "aac"),
        "webm": ("libvpx-vp9", "libopus"),
        "avi":  ("mpeg4", "mp3"),
    }
    vcodec, acodec = codecs.get(output_format, ("h264_nvenc", "aac"))

    cmd = [
        "ffmpeg",
        "-y",
        "-i", input_path,
        "-c:v", vcodec,
        "-b:v", str(video_bitrate),
        "-c:a", acodec,
        "-b:a", "128k",
        "-movflags", "+faststart" if output_format in ("mp4", "mov") else "",
        output_file
    ]
    cmd = [arg for arg in cmd if arg]

    print(f"\nInput file: {input_path}")
    print(f"Output file: {output_file}")
    print(f"Duration: {duration:.2f} seconds")
    print(f"Target size: {target_size_mb:.2f} MB")
    print(f"Video bitrate: {video_bitrate // 1000} kbps, Audio bitrate: 128 kbps")
    print(f"Command: {' '.join(cmd)}\n")

    def run_ffmpeg_with_progress():
        process = subprocess.Popen(
            cmd,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            bufsize=1
        )

        time_pattern = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")
        speed_pattern = re.compile(r"speed=([\d\.x]+)")
        fps_pattern = re.compile(r"fps=(\d+)")
        size_pattern = re.compile(r"size=\s*([\d\.]+)kB")
        bitrate_pattern = re.compile(r"bitrate=\s*([\d\.]+)kbits/s")

        start_time = time.time()
        last_percent = -1
        last_stats = {}

        for line in process.stderr:
            match = time_pattern.search(line)
            if match:
                h, m, s, ms = map(int, match.groups())
                elapsed = h * 3600 + m * 60 + s + ms / 100.0
                percent = int(100 * min(elapsed, duration) / duration)
                elapsed_wall = time.time() - start_time
                if elapsed > 0:
                    est_total = elapsed_wall * duration / elapsed
                    eta = est_total - elapsed_wall
                else:
                    eta = 0
                # Parse more stats
                speed = speed_pattern.search(line)
                fps = fps_pattern.search(line)
                size = size_pattern.search(line)
                bitrate = bitrate_pattern.search(line)
                last_stats = {
                    "speed": speed.group(1) if speed else "",
                    "fps": fps.group(1) if fps else "",
                    "size": size.group(1) if size else "",
                    "bitrate": bitrate.group(1) if bitrate else "",
                }
                if percent != last_percent:
                    print(
                        f"\rProgress: {percent:3d}% "
                        f"({elapsed:.1f}s / {duration:.1f}s) "
                        f"Elapsed: {elapsed_wall:.1f}s "
                        f"ETA: {eta:.1f}s "
                        f"Speed: {last_stats['speed']} "
                        f"FPS: {last_stats['fps']} "
                        f"Size: {last_stats['size']}kB "
                        f"Bitrate: {last_stats['bitrate']}kbits/s",
                        end="", flush=True
                    )
                    last_percent = percent
        process.wait()
        print("\rProgress: 100% ({:.1f}s / {:.1f}s) Elapsed: {:.1f}s ETA: 0.0s".format(duration, duration, time.time() - start_time))

        return process.returncode

    try:
        retcode = run_ffmpeg_with_progress()
        if retcode == 0:
            final_size = os.path.getsize(output_file) / (1024 * 1024)
            print(f"\nSuccess: Re-encoded video saved as:\n{output_file}\nFinal size: {final_size:.2f} MB")
        else:
            print("\nFFmpeg failed.")
    except Exception as e:
        print(f"\nFFmpeg failed: {e}")

if __name__ == "__main__":
    formats = ["mp4", "mkv", "webm", "avi", "mov"]
    file = choose_file()
    current_format, current_size = get_file_info(file)
    print(f"Format: {current_format}\nSize: {current_size:.2f} MB")
    out_format = ask_format(formats)
    target_size = ask_target_size(current_size)
    reencode_video(file, out_format, target_size)
