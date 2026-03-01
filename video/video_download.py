"""
Unified Video Downloader - GUI application for downloading videos from multiple sources.
Combines: YouTube (yt-dlp), HLS/M3U8 (ffmpeg), and Direct URL (HTTP with resume).
"""

import os
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from urllib.parse import urlparse, parse_qs

import requests
import yt_dlp


# -----------------------------------------------------------------------------
# YouTube download (yt-dlp)
# -----------------------------------------------------------------------------

def extract_start_time(url):
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    t = query.get("t", [0])[0]
    try:
        return int(t)
    except Exception:
        match = re.search(r"[?&]t=(\d+)", url)
        return int(match.group(1)) if match else 0


def _yt_progress_hook(d, status_callback, progress_callback):
    if d['status'] == 'downloading':
        downloaded = d.get('downloaded_bytes') or 0
        total = d.get('total_bytes') or d.get('total_bytes_estimate')
        try:
            if total and total > 0:
                value = min(100.0, 100.0 * downloaded / total)
            else:
                value = 0.0
            if progress_callback and value >= 0:
                progress_callback(value)
        except (TypeError, ZeroDivisionError):
            pass
    elif d['status'] == 'finished':
        if progress_callback:
            progress_callback(100)


def download_youtube(url, output_path, start_time=0, status_callback=None, progress_callback=None):
    ydl_opts = {
        'outtmpl': os.path.join(output_path, '%(title).70s.%(ext)s'),
        'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/mp4',
        'noplaylist': True,
        'quiet': True,
        'progress_hooks': [
            lambda d: _yt_progress_hook(d, status_callback, progress_callback)
        ],
        'postprocessors': [{
            'key': 'FFmpegVideoRemuxer',
            'preferedformat': 'mp4'
        }]
    }
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            filename = ydl.prepare_filename(info)
        if progress_callback:
            progress_callback(100)
        if start_time > 0:
            base, ext = os.path.splitext(filename)
            new_name = f"{base}_start@{start_time}s{ext}"
            os.rename(filename, new_name)
        if status_callback:
            status_callback(f"Done: {url}")
    except Exception as e:
        if progress_callback:
            progress_callback(0)
        if status_callback:
            status_callback(f"Failed {url}: {e}")
        else:
            raise


# -----------------------------------------------------------------------------
# HLS/M3U8 download (ffmpeg)
# -----------------------------------------------------------------------------

def download_hls(m3u8_url, output_path, filename, headers=None, status_callback=None, progress_callback=None):
    import subprocess
    output_file = os.path.join(output_path, filename)
    command = ["ffmpeg", "-y"]
    if headers:
        for key, value in headers.items():
            command.extend(["-headers", f"{key}: {value}"])
    command.extend(["-i", m3u8_url, "-c", "copy", output_file])
    try:
        if status_callback:
            status_callback(f"Downloading HLS: {m3u8_url[:50]}...")
        # ffmpeg doesn't easily expose progress; run and signal 100% on success
        subprocess.run(command, check=True)
        if progress_callback:
            progress_callback(100)
        if status_callback:
            status_callback(f"Done: {m3u8_url}")
    except subprocess.CalledProcessError as e:
        if progress_callback:
            progress_callback(0)
        if status_callback:
            status_callback(f"Failed: {e}")
        else:
            raise


# -----------------------------------------------------------------------------
# Direct URL download (requests, with resume)
# -----------------------------------------------------------------------------

def download_direct_url(url, output_path, filename, status_callback=None, progress_callback=None):
    output_file = os.path.join(output_path, filename)
    existing_file_size = 0
    while True:
        headers = {}
        if os.path.exists(output_file):
            existing_file_size = os.path.getsize(output_file)
            headers['Range'] = f'bytes={existing_file_size}-'
        try:
            with requests.get(url, stream=True, headers=headers, timeout=30) as response:
                response.raise_for_status()
                total_size = int(response.headers.get('content-length', 0)) + existing_file_size
                if total_size == 0:
                    total_size = existing_file_size
                block_size = 8192
                if status_callback:
                    status_callback(f"Downloading: {url[:50]}...")
                with open(output_file, 'ab') as f:
                    for data in response.iter_content(block_size):
                        f.write(data)
                        if total_size > 0 and progress_callback:
                            progress_callback(min(100.0, 100.0 * f.tell() / total_size))
                if total_size > 0 and os.path.getsize(output_file) >= total_size:
                    if progress_callback:
                        progress_callback(100)
                    if status_callback:
                        status_callback(f"Done: {url}")
                    return
        except requests.exceptions.RequestException as e:
            if status_callback:
                status_callback(f"Failed: {e}")
            raise
        if status_callback:
            status_callback("Incomplete download, retrying...")
        time.sleep(5)


# -----------------------------------------------------------------------------
# URL type detection
# -----------------------------------------------------------------------------

def is_youtube_url(url):
    return 'youtube.com' in url or 'youtu.be' in url


def is_hls_url(url):
    return '.m3u8' in url.lower()


def detect_url_type(url):
    if is_youtube_url(url):
        return 'youtube'
    if is_hls_url(url):
        return 'hls'
    return 'direct'


# -----------------------------------------------------------------------------
# GUI Application
# -----------------------------------------------------------------------------

class VideoDownloaderApp:
    DEFAULT_HLS_HEADERS = {
        "Referer": "https://www.linkedin.com/",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Video Downloader")
        self.root.geometry("750x550")
        self.root.minsize(600, 450)

        self.folder_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.mode_var = tk.StringVar(value="auto")
        self.downloading = False

        self._build_ui()

    def _build_ui(self):
        # Mode selector
        mode_frame = tk.Frame(self.root)
        mode_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Label(mode_frame, text="Mode:").pack(side=tk.LEFT, padx=(0, 8))
        for val, label in [
            ("auto", "Auto-detect"),
            ("youtube", "YouTube"),
            ("hls", "HLS/M3U8"),
            ("direct", "Direct URL"),
        ]:
            tk.Radiobutton(mode_frame, text=label, variable=self.mode_var, value=val).pack(side=tk.LEFT, padx=(0, 12))

        # Mode explanation
        help_text = (
            "Auto-detect: picks method from URL (youtube.com → yt-dlp, .m3u8 → ffmpeg, else → direct HTTP). "
            "YouTube: force yt-dlp (supports ?t= start time). HLS/M3U8: for live streams (LinkedIn, etc). "
            "Direct URL: for direct .mp4 links with resume support."
        )
        tk.Label(self.root, text=help_text, wraplength=700, fg="gray", font=("", 9), justify=tk.LEFT).pack(
            fill=tk.X, padx=10, pady=(0, 5), anchor="w"
        )

        # Notebook with tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Tab 1: URLs (shared)
        url_tab = tk.Frame(self.notebook, padx=8, pady=8)
        self.notebook.add(url_tab, text="URLs")
        tk.Label(url_tab, text="Paste one URL per line. Supports YouTube, HLS (.m3u8), or direct video links.", fg="gray", font=("", 9)).pack(anchor="w")
        url_frame = tk.LabelFrame(url_tab, text="URLs", padx=8, pady=8)
        url_frame.pack(fill=tk.BOTH, expand=True)
        self.url_text = scrolledtext.ScrolledText(url_frame, height=12, wrap=tk.WORD, font=("Consolas", 10))
        self.url_text.pack(fill=tk.BOTH, expand=True)

        # Tab 2: HLS options
        hls_tab = tk.Frame(self.notebook, padx=8, pady=8)
        self.notebook.add(hls_tab, text="HLS Options")
        tk.Label(
            hls_tab, text="Used for HLS/M3U8 streams (e.g. LinkedIn). If the server checks the origin, set Referer (e.g. https://www.linkedin.com/) and User-Agent. Requires ffmpeg.", fg="gray", wraplength=650, font=("", 9), justify=tk.LEFT
        ).pack(anchor="w")
        hls_frame = tk.LabelFrame(hls_tab, text="Custom headers", padx=8, pady=8)
        hls_frame.pack(fill=tk.X)
        self.hls_referer = tk.Entry(hls_frame, width=60)
        self.hls_referer.insert(0, self.DEFAULT_HLS_HEADERS.get("Referer", ""))
        tk.Label(hls_frame, text="Referer:").grid(row=0, column=0, sticky="w", pady=2)
        self.hls_referer.grid(row=0, column=1, sticky="ew", pady=2, padx=(8, 0))
        self.hls_user_agent = tk.Entry(hls_frame, width=60)
        self.hls_user_agent.insert(0, self.DEFAULT_HLS_HEADERS.get("User-Agent", ""))
        tk.Label(hls_frame, text="User-Agent:").grid(row=1, column=0, sticky="w", pady=2)
        self.hls_user_agent.grid(row=1, column=1, sticky="ew", pady=2, padx=(8, 0))
        hls_frame.columnconfigure(1, weight=1)

        # Folder section
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(folder_frame, text="Download folder (required):").pack(side=tk.LEFT, padx=(0, 8))
        self.folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path, width=50)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(folder_frame, text="Browse...", command=self._browse_folder).pack(side=tk.LEFT)

        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        self.download_btn = tk.Button(
            btn_frame, text="Download", command=self._start_download,
            bg="#4CAF50", fg="white", font=("", 10, "bold"), cursor="hand2",
            padx=20, pady=5
        )
        self.download_btn.pack(side=tk.LEFT, padx=(0, 8))

        # Progress
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X)
        self.progress_label = tk.Label(progress_frame, text="0%", fg="gray")
        self.progress_label.pack(anchor="e")

        self.status_label = tk.Label(self.root, textvariable=self.status_var, fg="gray", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select Download Folder")
        if folder:
            self.folder_path.set(folder)

    def _update_status(self, text):
        self.status_var.set(text)
        self.root.update_idletasks()

    def _update_progress(self, percent):
        self.progress_bar['value'] = percent
        self.progress_label.config(text=f"{percent:.1f}%")
        self.root.update_idletasks()

    def _get_hls_headers(self):
        referer = self.hls_referer.get().strip()
        ua = self.hls_user_agent.get().strip()
        h = {}
        if referer:
            h["Referer"] = referer
        if ua:
            h["User-Agent"] = ua
        return h if h else None

    def _download_one(self, url, folder, mode, index, total, status_cb, progress_cb):
        resolved = mode if mode != "auto" else detect_url_type(url)
        if resolved == "youtube":
            start = extract_start_time(url)
            download_youtube(url, folder, start, status_cb, progress_cb)
        elif resolved == "hls":
            name = "downloaded_video.mp4" if total == 1 and len(url) <= 80 else f"hls_{index}_{abs(hash(url)) % 10000}.mp4"
            download_hls(url, folder, name, self._get_hls_headers(), status_cb, progress_cb)
        else:
            name = os.path.basename(urlparse(url).path) or "downloaded_video.mp4"
            if not name or name == "/":
                name = f"direct_{index}_{abs(hash(url)) % 10000}.mp4"
            elif total > 1:
                base, ext = os.path.splitext(name)
                name = f"{base}_{index}{ext}"
            download_direct_url(url, folder, name, status_cb, progress_cb)

    def _start_download(self):
        urls_text = self.url_text.get("1.0", tk.END)
        urls = [u.strip() for u in urls_text.strip().splitlines() if u.strip()]
        if not urls:
            messagebox.showwarning("No URLs", "Please enter at least one URL.")
            return
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("Invalid Folder", "Please select a valid download folder.")
            return

        self.download_btn.config(state=tk.DISABLED)
        self._update_progress(0)
        mode = self.mode_var.get()

        def status_cb(msg):
            self.root.after(0, lambda m=msg: self._update_status(m))

        def progress_cb(pct):
            self.root.after(0, lambda p=pct: self._update_progress(p))

        def on_complete():
            self.root.after(0, self._download_complete)

        def worker():
            for i, url in enumerate(urls, 1):
                status_cb(f"Downloading {i}/{len(urls)}: {url[:60]}...")
                progress_cb(0)
                try:
                    self._download_one(url, folder, mode, i, len(urls), status_cb, progress_cb)
                except Exception as e:
                    status_cb(f"Failed {url}: {e}")
            on_complete()

        threading.Thread(target=worker, daemon=True).start()

    def _download_complete(self):
        self.download_btn.config(state=tk.NORMAL)
        self._update_progress(100)
        self._update_status("All downloads complete.")
        messagebox.showinfo("Done", "All downloads complete.")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = VideoDownloaderApp()
    app.run()
