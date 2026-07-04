"""
Video Re-encoder - Re-encode videos to a target file size with FFmpeg.
Supports both GUI and command-line interfaces. Uses GPU (NVENC) when available.
"""

import os
import re
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from _ffmpeg_utils import check_gpu_available, get_video_duration


def get_file_info(file_path):
    size_bytes = os.path.getsize(file_path)
    size_mb = size_bytes / (1024 * 1024)
    ext = os.path.splitext(file_path)[1].lower().replace(".", "")
    return ext, size_mb


def reencode_video(
    input_path,
    output_format,
    target_size_mb,
    progress_callback=None,
    status_callback=None,
):
    duration = get_video_duration(input_path)
    if duration is None or duration <= 0:
        if status_callback:
            status_callback("Could not determine video duration.")
        return False

    total_bitrate = (target_size_mb * 8 * 1024 * 1024) / duration
    audio_bitrate = 128 * 1024
    video_bitrate = int(total_bitrate - audio_bitrate)
    if video_bitrate < 100_000:
        if status_callback:
            status_callback("Target size too small for reasonable quality.")
        return False

    input_dir = os.path.dirname(input_path)
    name, _ = os.path.splitext(os.path.basename(input_path))
    output_file = os.path.join(input_dir, f"{name}_reencoded.{output_format}")

    use_gpu = check_gpu_available()
    codecs = {
        "mp4": ("h264_nvenc" if use_gpu else "libx264", "aac"),
        "mkv": ("h264_nvenc" if use_gpu else "libx264", "aac"),
        "mov": ("h264_nvenc" if use_gpu else "libx264", "aac"),
        "webm": ("libvpx-vp9", "libopus"),
        "avi": ("mpeg4", "mp3"),
    }
    vcodec, acodec = codecs.get(output_format, ("libx264", "aac"))

    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-c:v", vcodec,
        "-b:v", str(video_bitrate),
        "-c:a", acodec,
        "-b:a", "128k",
        "-movflags", "+faststart" if output_format in ("mp4", "mov") else "",
        output_file,
    ]
    cmd = [arg for arg in cmd if arg]

    time_pattern = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")

    def run():
        process = subprocess.Popen(
            cmd,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            bufsize=1,
        )
        for line in process.stderr:
            if progress_callback and "time=" in line:
                m = time_pattern.search(line)
                if m:
                    h, mn, s, ms = map(int, m.groups())
                    elapsed = h * 3600 + mn * 60 + s + ms / 100.0
                    pct = min(100.0, 100.0 * elapsed / duration)
                    progress_callback(pct)
        process.wait()
        if process.returncode == 0:
            if progress_callback:
                progress_callback(100.0)
            if status_callback:
                status_callback(f"Done: {output_file}")
            return True
        if status_callback:
            status_callback("FFmpeg failed.")
        return False

    try:
        return run()
    except FileNotFoundError:
        if status_callback:
            status_callback("FFmpeg not found. Add it to your PATH.")
        return False
    except Exception as e:
        if status_callback:
            status_callback(str(e))
        return False


# -----------------------------------------------------------------------------
# GUI Application
# -----------------------------------------------------------------------------

class VideoReencoderApp:
    FORMATS = ["mp4", "mkv", "webm", "avi", "mov"]

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Video Re-encoder")
        self.root.geometry("520x380")
        self.root.minsize(450, 320)

        self.input_path = tk.StringVar()
        self.output_format = tk.StringVar(value="mp4")
        self.target_size_mb = tk.StringVar(value="50")
        self.status_var = tk.StringVar(value="Select a video file")
        self.current_size_mb = 0.0
        self.processing = False

        self._build_ui()

    def _build_ui(self):
        # File
        file_frame = tk.LabelFrame(self.root, text="Input video", padx=10, pady=8)
        file_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Entry(file_frame, textvariable=self.input_path, width=55, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(file_frame, text="Browse...", command=self._browse).pack(side=tk.LEFT)

        # Size info
        self.size_label = tk.Label(self.root, text="", fg="gray")
        self.size_label.pack(anchor="w", padx=10, pady=2)

        # Options
        opt_frame = tk.LabelFrame(self.root, text="Output settings", padx=10, pady=8)
        opt_frame.pack(fill=tk.X, padx=10, pady=5)
        opt_frame.columnconfigure(1, weight=1)

        tk.Label(opt_frame, text="Format:").grid(row=0, column=0, sticky="w", pady=2)
        fmt_combo = ttk.Combobox(opt_frame, textvariable=self.output_format, values=self.FORMATS, width=10, state="readonly")
        fmt_combo.grid(row=0, column=1, sticky="w", padx=(8, 16), pady=2)

        tk.Label(opt_frame, text="Target size (MB):").grid(row=1, column=0, sticky="w", pady=2)
        tk.Entry(opt_frame, textvariable=self.target_size_mb, width=12).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=2)

        # Process
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)
        self.process_btn = tk.Button(
            btn_frame, text="Re-encode", command=self._start_process,
            bg="#4CAF50", fg="white", font=("", 10, "bold"), cursor="hand2", padx=20, pady=5
        )
        self.process_btn.pack(side=tk.LEFT)

        # Progress
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, maximum=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X)
        self.progress_label = tk.Label(progress_frame, text="0%", fg="gray")
        self.progress_label.pack(anchor="e")

        self.status_label = tk.Label(self.root, textvariable=self.status_var, fg="gray", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select video file",
            filetypes=[("Video files", "*.mp4 *.mov *.mkv *.webm *.avi")],
        )
        if path:
            self.input_path.set(path)
            ext, size_mb = get_file_info(path)
            self.current_size_mb = size_mb
            self.size_label.config(text=f"Current: {ext.upper()} | {size_mb:.2f} MB")
            self.status_var.set("Ready")

    def _update_progress(self, percent):
        self.progress_bar["value"] = percent
        self.progress_label.config(text=f"{percent:.1f}%")
        self.root.update_idletasks()

    def _start_process(self):
        path = self.input_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("No file", "Select a video file first.")
            return
        try:
            size_mb = float(self.target_size_mb.get())
        except ValueError:
            messagebox.showerror("Invalid size", "Enter a valid target size (MB).")
            return
        if size_mb <= 0:
            messagebox.showerror("Invalid size", "Target size must be positive.")
            return
        if self.processing:
            return
        self.processing = True
        self.process_btn.config(state=tk.DISABLED)
        self._update_progress(0)

        def worker():
            def status_cb(msg):
                self.root.after(0, lambda m=msg: self.status_var.set(m))

            def progress_cb(pct):
                self.root.after(0, lambda p=pct: self._update_progress(p))

            ok = reencode_video(
                path,
                self.output_format.get(),
                size_mb,
                progress_callback=progress_cb,
                status_callback=status_cb,
            )
            self.root.after(0, lambda: self._process_complete(ok))

        threading.Thread(target=worker, daemon=True).start()

    def _process_complete(self, success):
        self.processing = False
        self.process_btn.config(state=tk.NORMAL)
        self._update_progress(100 if success else 0)
        if success:
            self.status_var.set("Done!")
            messagebox.showinfo("Success", "Re-encoded video saved in the same folder.")
        else:
            messagebox.showerror("Error", "Re-encoding failed. Check status for details.")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = VideoReencoderApp()
    app.run()
