import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from urllib.parse import urlparse, parse_qs
import subprocess
from tqdm import tqdm
import yt_dlp

# Your video list
video_urls = [
    "https://youtu.be/1AT5klu_yAQ",
    "https://youtu.be/6XvC13Zfxgk",
    "https://youtu.be/6NHySKiUJfs",
    "https://youtu.be/uJKDipbCzdc",
    "https://youtu.be/LhoLuui9gX8",
    "https://youtu.be/k9Z9z_Fd3Bw",
    "https://youtu.be/OXnuPK1obWI",
    "https://youtu.be/rrkrvAUbU9Y",
    "https://youtu.be/YyXRYgjQXX0",
    "https://youtu.be/ydeGSuFlbDA",
    "https://youtu.be/40TumEHQk8A"
]

# Extract timestamp
def extract_start_time(url):
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    t = query.get("t", [0])[0]
    try:
        return int(t)
    except:
        match = re.search(r"[?&]t=(\d+)", url)
        return int(match.group(1)) if match else 0

# Select folder
def choose_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select Download Folder")
    if not folder:
        messagebox.showerror("Error", "No folder selected. Exiting.")
        exit()
    return folder

# Download using yt-dlp
def download_video(url, output_path, start_time):
    # Set output template and options
    ydl_opts = {
        'outtmpl': os.path.join(output_path, '%(title).70s.%(ext)s'),
        'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/mp4',
        'noplaylist': True,
        'quiet': True,
        'progress_hooks': [progress_hook],
        'postprocessors': [{
            'key': 'FFmpegVideoRemuxer',
            'preferedformat': 'mp4'
        }]
    }

    global pbar
    pbar = tqdm(total=100, desc="Starting", position=0)

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            filename = ydl.prepare_filename(info)

        pbar.close()

        # Rename if start time exists
        if start_time > 0:
            base, ext = os.path.splitext(filename)
            new_name = f"{base}_start@{start_time}s{ext}"
            os.rename(filename, new_name)

    except Exception as e:
        pbar.close()
        print(f"Failed to download {url}: {e}")

# Update progress bar
def progress_hook(d):
    if d['status'] == 'downloading':
        percent = d.get('_percent_str', '0.0%').strip().replace('%', '')
        try:
            pbar.n = float(percent)
            pbar.refresh()
        except:
            pass
    elif d['status'] == 'finished':
        pbar.n = 100
        pbar.refresh()
        pbar.set_description("Download complete")

# Run all
if __name__ == "__main__":
    folder = choose_folder()
    for url in video_urls:
        start_time = extract_start_time(url)
        download_video(url, folder, start_time)
    print("\n✅ All downloads complete.")
