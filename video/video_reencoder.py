import json
import os
import re
import subprocess
import time
import tkinter as tk
from tkinter import filedialog

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

def get_audio_tracks_info(input_path):
    """Get information about audio tracks in the video file."""
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error", "-select_streams", "a:0",
                "-show_entries", "stream=index,codec_name,channels,bit_rate",
                "-of", "json", input_path
            ],
            capture_output=True, text=True, check=True
        )
        data = json.loads(result.stdout)

        tracks = []
        if 'streams' in data:
            for stream in data['streams']:
                track_info = {
                    'index': stream.get('index'),
                    'codec': stream.get('codec_name', 'unknown'),
                    'channels': stream.get('channels', 2),
                    'bitrate': stream.get('bit_rate', 'unknown')
                }
                tracks.append(track_info)

        # Also check for additional streams
        result_all = subprocess.run(
            [
                "ffprobe", "-v", "error", "-select_streams", "a",
                "-show_entries", "stream=index,codec_name,channels,bit_rate",
                "-of", "json", input_path
            ],
            capture_output=True, text=True, check=True
        )
        data_all = json.loads(result_all.stdout)
        if 'streams' in data_all and len(data_all['streams']) > len(tracks):
            for stream in data_all['streams'][len(tracks):]:
                track_info = {
                    'index': stream.get('index'),
                    'codec': stream.get('codec_name', 'unknown'),
                    'channels': stream.get('channels', 2),
                    'bitrate': stream.get('bit_rate', 'unknown')
                }
                tracks.append(track_info)

        return tracks
    except Exception as e:
        print(f"Could not get audio track information: {e}")
        return []

def ask_audio_track_choice(audio_tracks):
    """Ask user which audio track to use or if to merge all."""
    if not audio_tracks:
        return None

    if len(audio_tracks) == 1:
        print(f"Found 1 audio track: {audio_tracks[0]['codec']} ({audio_tracks[0]['channels']}ch)")
        return [0]  # Single track, no choice needed

    print(f"\nFound {len(audio_tracks)} audio tracks:")
    for i, track in enumerate(audio_tracks):
        codec = track['codec']
        channels = track['channels']
        bitrate = track.get('bitrate', 'unknown')
        print(f"  {i+1}. Track {track['index']}: {codec} ({channels}ch, {bitrate} bps)")

    print(f"  {len(audio_tracks)+1}. Merge all tracks (play simultaneously)")

    while True:
        try:
            choice = input(f"Choose track (1-{len(audio_tracks)+1}): ").strip()
            choice_num = int(choice)

            if 1 <= choice_num <= len(audio_tracks):
                return [choice_num - 1]  # Single track selection
            elif choice_num == len(audio_tracks) + 1:
                return list(range(len(audio_tracks)))  # All tracks for merging
            else:
                print("Invalid choice.")
        except ValueError:
            print("Please enter a number.")

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

    # Get audio tracks information
    audio_tracks = get_audio_tracks_info(input_path)
    selected_tracks = ask_audio_track_choice(audio_tracks)
    if selected_tracks is None:
        print("No audio tracks found.")
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

    # Build FFmpeg command with track selection/merging
    cmd = ["ffmpeg", "-y"]

    # Add video input
    cmd.extend(["-i", input_path])

    # Select video stream (usually stream 0)
    cmd.extend(["-map", "0:v"])

    # Handle audio track selection/merging
    if len(selected_tracks) == 1:
        # Single track selection
        track_index = audio_tracks[selected_tracks[0]]['index']
        cmd.extend(["-map", f"0:{track_index}"])
        print(f"Using audio track {track_index}")
    else:
        # Multiple tracks - merge them
        print(f"Merging {len(selected_tracks)} audio tracks simultaneously")

        # Add audio mixing filter for multiple tracks
        filter_parts = []
        for i, track_idx in enumerate(selected_tracks):
            track_index = audio_tracks[track_idx]['index']
            filter_parts.append(f"[0:{track_index}]")

        filter_input = "".join(filter_parts)
        filter_complex = f"{filter_input}amix=inputs={len(selected_tracks)}:duration=longest[aout]"
        cmd.extend(["-filter_complex", filter_complex, "-map", "[aout]"])

    # Add video codec and bitrate
    cmd.extend(["-c:v", vcodec, "-b:v", str(video_bitrate)])

    # Add audio codec and bitrate
    cmd.extend(["-c:a", acodec, "-b:a", "128k"])

    # Add movflags for MP4/MOV
    if output_format in ("mp4", "mov"):
        cmd.extend(["-movflags", "+faststart"])

    # Add output file
    cmd.append(output_file)

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
            stdout=subprocess.PIPE,
            universal_newlines=True,
            bufsize=1
        )

        time_pattern = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")
        speed_pattern = re.compile(r"speed=([\d\.x]+)")
        fps_pattern = re.compile(r"fps=(\d+)")
        size_pattern = re.compile(r"size=\s*([\d\.]+)kB")
        bitrate_pattern = re.compile(r"bitrate=\s*([\d\.]+)kbits/s")
        error_pattern = re.compile(r"error|Error|ERROR")

        start_time = time.time()
        last_percent = -1
        last_stats = {}
        error_lines = []

        for line in process.stderr:
            # Check for errors
            if error_pattern.search(line) and not any(skip in line.lower() for skip in ["time=", "speed=", "fps=", "size=", "bitrate="]):
                error_lines.append(line.strip())

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

        # Show final progress
        elapsed_wall = time.time() - start_time
        if last_percent >= 0:
            print(
                f"\rProgress: {last_percent:3d}% "
                f"({duration:.1f}s / {duration:.1f}s) "
                f"Elapsed: {elapsed_wall:.1f}s "
                f"ETA: 0.0s "
                f"Speed: {last_stats.get('speed', '')} "
                f"FPS: {last_stats.get('fps', '')} "
                f"Size: {last_stats.get('size', '')}kB "
                f"Bitrate: {last_stats.get('bitrate', '')}kbits/s"
            )
        else:
            print(f"\rProcess completed in {elapsed_wall:.1f}s")

        # Show any errors that occurred
        if error_lines:
            print("\nFFmpeg errors encountered:")
            for error in error_lines[-5:]:  # Show last 5 errors
                print(f"  {error}")

        return process.returncode

    try:
        retcode = run_ffmpeg_with_progress()
        if retcode == 0:
            final_size = os.path.getsize(output_file) / (1024 * 1024)
            print(f"\nSuccess: Re-encoded video saved as:\n{output_file}\nFinal size: {final_size:.2f} MB")
        else:
            print(f"\nFFmpeg failed with return code: {retcode}")
            print("This could be due to:")
            print("- Invalid FFmpeg command syntax")
            print("- Unsupported codec combination")
            print("- Insufficient disk space")
            print("- File permission issues")
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
