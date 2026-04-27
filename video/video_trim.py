"""
Video Trimmer - Trim videos with FFmpeg (GPU accelerated if available).
Supports both GUI and command-line interfaces.
"""

import argparse
import logging
import os
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

log = logging.getLogger(__name__)

# -------------------------------------------------------
# Time parsing: mm:ss, mm:ss.s, hh:mm:ss, hh:mm:ss.s
# -------------------------------------------------------

def time_to_seconds(time_str):
    """Convert time string to float seconds. Supports mm:ss, mm:ss.s, hh:mm:ss, hh:mm:ss.s"""
    time_str = str(time_str).strip()
    if not time_str:
        raise ValueError("Empty time string")
    parts = time_str.replace(",", ".").split(":")
    if len(parts) == 2:
        m, s = parts[0], parts[1]
        try:
            minutes = int(m)
            seconds = float(s)
        except ValueError:
            raise ValueError(f"Invalid time format: {time_str}")
        return minutes * 60 + seconds
    elif len(parts) == 3:
        h, m, s = parts[0], parts[1], parts[2]
        try:
            hours = int(h)
            minutes = int(m)
            seconds = float(s)
        except ValueError:
            raise ValueError(f"Invalid time format: {time_str}")
        return hours * 3600 + minutes * 60 + seconds
    else:
        raise ValueError(f"Invalid time format: {time_str}")


def seconds_to_time_str(seconds, with_fraction=True):
    """Convert float seconds to mm:ss.s or hh:mm:ss.s string"""
    seconds = max(0, float(seconds))
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    if hours > 0:
        fmt = f"{hours}:{minutes:02d}:{secs:05.2f}" if with_fraction else f"{hours}:{minutes:02d}:{int(secs):02d}"
    else:
        fmt = f"{minutes}:{secs:05.2f}" if with_fraction else f"{minutes}:{int(secs):02d}"
    return fmt.rstrip("0").rstrip(".").rstrip(":")


# -------------------------------------------------------
# FFprobe: get video duration
# -------------------------------------------------------

def get_video_duration(input_path):
    """Get video duration in seconds using ffprobe. Returns float or None on error."""
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error", "-show_entries",
                "format=duration", "-of",
                "default=noprint_wrappers=1:nokey=1", input_path
            ],
            capture_output=True, text=True, check=True, timeout=30
        )
        return float(result.stdout.strip())
    except (subprocess.CalledProcessError, FileNotFoundError, ValueError):
        return None


# -------------------------------------------------------
# GPU check
# -------------------------------------------------------

def check_gpu_available():
    try:
        result = subprocess.run(["nvidia-smi"], capture_output=True, text=True)
        if result.returncode != 0:
            return False
        encoders_check = subprocess.run(["ffmpeg", "-encoders"], capture_output=True, text=True)
        return "h264_nvenc" in encoders_check.stdout
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


# -------------------------------------------------------
# Core trim: runs ffmpeg with progress callback
# -------------------------------------------------------

def trim_video(input_path, start_sec, end_sec, progress_callback=None, status_callback=None, output_path=None):
    """
    Trim video between start_sec and end_sec (float seconds).
    Optional progress_callback(percent), status_callback(message), output_path.
    Returns output file path on success, raises on error.
    """
    input_dir = os.path.dirname(input_path)
    filename = os.path.basename(input_path)
    name, ext = os.path.splitext(filename)

    if output_path is not None:
        output_file = output_path
    else:
        start_label = seconds_to_time_str(start_sec).replace(":", "-").replace(".", "-")
        end_label = seconds_to_time_str(end_sec).replace(":", "-").replace(".", "-")
        output_file = os.path.join(input_dir, f"{name}_trim_{start_label}_{end_label}{ext}")

    duration = end_sec - start_sec
    if duration <= 0:
        raise ValueError("End time must be after start time.")

    use_gpu = check_gpu_available()

    if use_gpu:
        ffmpeg_cmd = [
            "ffmpeg", "-y",
            "-hwaccel", "cuda",
            "-hwaccel_output_format", "cuda",
            "-ss", str(start_sec),
            "-i", input_path,
            "-t", str(duration),
            "-c:v", "h264_nvenc", "-preset", "fast", "-cq", "23",
            "-c:a", "aac", "-b:a", "192k",
            "-af", "loudnorm",
            "-movflags", "+faststart",
            output_file
        ]
    else:
        ffmpeg_cmd = [
            "ffmpeg", "-y",
            "-ss", str(start_sec),
            "-i", input_path,
            "-t", str(duration),
            "-c:v", "libx264", "-preset", "fast", "-crf", "23",
            "-c:a", "aac", "-b:a", "192k",
            "-af", "loudnorm",
            "-movflags", "+faststart",
            output_file
        ]

    # Pattern to parse FFmpeg stderr: time=HH:MM:SS.ms, time=MM:SS.ms, or time=SS.ms
    time_pattern = re.compile(
        r"time=(\d+):(\d+):(\d+)\.(\d+)|time=(\d+):(\d+)\.(\d+)|time=(\d+)\.(\d+)"
    )

    def run_ffmpeg():
        process = subprocess.Popen(
            ffmpeg_cmd,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            bufsize=1
        )
        for line in process.stderr:
            if "time=" in line:
                # Parse time from ffmpeg output
                m = time_pattern.search(line)
                if m:
                    if m.group(1) is not None:  # HH:MM:SS.ms
                        h, mn, s, ms = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                        elapsed = h * 3600 + mn * 60 + s + ms / 100.0
                    elif m.group(5) is not None:  # MM:SS.ms
                        mn, s, ms = int(m.group(5)), int(m.group(6)), int(m.group(7))
                        elapsed = mn * 60 + s + ms / 100.0
                    else:  # SS.ms
                        s, ms = int(m.group(8)), int(m.group(9))
                        elapsed = s + ms / 100.0
                    pct = min(100.0, 100.0 * elapsed / duration)
                    if progress_callback:
                        progress_callback(pct)
        process.wait()
        if process.returncode != 0:
            raise subprocess.CalledProcessError(process.returncode, ffmpeg_cmd)
        if progress_callback:
            progress_callback(100.0)
        return output_file

    return run_ffmpeg()


def cut_middle_video(
    input_path, cut_start_sec, cut_end_sec,
    progress_callback=None, status_callback=None
):
    """
    Cut a middle section from the video and join the remaining parts.
    Removes [cut_start_sec, cut_end_sec], keeps [0, cut_start_sec] + [cut_end_sec, end].
    Returns output file path on success, raises on error.
    """
    input_dir = os.path.dirname(input_path)
    filename = os.path.basename(input_path)
    name, ext = os.path.splitext(filename)
    duration_sec = get_video_duration(input_path)
    if duration_sec is None:
        raise ValueError("Could not detect video duration.")

    cut_start_sec = max(0, float(cut_start_sec))
    cut_end_sec = min(float(cut_end_sec), duration_sec)
    if cut_end_sec <= cut_start_sec:
        raise ValueError("Cut end must be after cut start.")

    start_label = seconds_to_time_str(cut_start_sec).replace(":", "-").replace(".", "-")
    end_label = seconds_to_time_str(cut_end_sec).replace(":", "-").replace(".", "-")
    output_file = os.path.join(input_dir, f"{name}_cut_{start_label}_{end_label}{ext}")

    fd1, temp1 = tempfile.mkstemp(suffix=ext, prefix="video_trim_part1_")
    os.close(fd1)
    fd2, temp2 = tempfile.mkstemp(suffix=ext, prefix="video_trim_part2_")
    os.close(fd2)
    fd_list, list_file = tempfile.mkstemp(suffix=".txt", prefix="video_concat_")
    os.close(fd_list)

    try:
        # Part 1: 0 to cut_start_sec (keep before the cut)
        has_part1 = cut_start_sec > 0
        if has_part1:
            trim_video(input_path, 0, cut_start_sec, output_path=temp1,
                       progress_callback=progress_callback, status_callback=status_callback)
        else:
            temp1 = None

        # Part 2: cut_end_sec to end (keep after the cut)
        has_part2 = cut_end_sec < duration_sec
        if has_part2:
            trim_video(input_path, cut_end_sec, duration_sec, output_path=temp2,
                       progress_callback=progress_callback, status_callback=status_callback)
        else:
            temp2 = None

        # Handle edge cases
        if temp1 is None and temp2 is None:
            raise ValueError("Nothing left after cut.")
        if temp1 is None:
            os.rename(temp2, output_file)
            return output_file
        if temp2 is None:
            os.rename(temp1, output_file)
            return output_file

        # Concatenate
        abs1 = os.path.abspath(temp1).replace("\\", "/")
        abs2 = os.path.abspath(temp2).replace("\\", "/")
        with open(list_file, "w") as f:
            f.write(f"file '{abs1}'\n")
            f.write(f"file '{abs2}'\n")

        cmd = ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", list_file,
               "-c", "copy", "-movflags", "+faststart", output_file]
        subprocess.run(cmd, check=True, capture_output=True, timeout=300)
        return output_file
    finally:
        for p in (temp1, temp2, list_file):
            if p is not None and os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass

class VideoTrimApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Video Trimmer")
        self.root.geometry("720x580")
        self.root.minsize(600, 450)

        self.video_path = tk.StringVar()
        self.duration_sec = 0.0
        self.start_sec = tk.DoubleVar(value=0.0)
        self.end_sec = tk.DoubleVar(value=0.0)
        self.mode_var = tk.StringVar(value="trim")  # "trim" or "cut_middle"
        self.status_var = tk.StringVar(value="Select a video to begin")
        self.queue = []  # List of (input_path, start_sec, end_sec, mode)
        self.processing = False

        self._build_ui()

    def _build_ui(self):
        # File selection
        file_frame = tk.LabelFrame(self.root, text="Video file", padx=10, pady=8)
        file_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        tk.Entry(file_frame, textvariable=self.video_path, width=70, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(file_frame, text="Browse...", command=self._browse).pack(side=tk.LEFT)

        # Duration display
        dur_frame = tk.Frame(self.root)
        dur_frame.pack(fill=tk.X, padx=10, pady=2)
        self.duration_label = tk.Label(dur_frame, text="Duration: —", fg="gray")
        self.duration_label.pack(anchor="w")

        # Mode selector
        mode_frame = tk.LabelFrame(self.root, text="Operation", padx=10, pady=6)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Radiobutton(mode_frame, text="1) Trim — keep start to end", variable=self.mode_var,
                       value="trim", command=self._on_mode_change).pack(anchor="w")
        tk.Radiobutton(mode_frame, text="2) Cut middle part — remove a section and join", variable=self.mode_var,
                       value="cut_middle", command=self._on_mode_change).pack(anchor="w")

        # Time controls
        self.time_frame = tk.LabelFrame(self.root, text="", padx=10, pady=8)
        self.time_frame.pack(fill=tk.X, padx=10, pady=5)
        self.time_frame.columnconfigure(2, weight=1)

        self.start_label_widget = tk.Label(self.time_frame, text="Start:")
        self.start_label_widget.grid(row=0, column=0, sticky="w", pady=2)
        self.start_entry = tk.Entry(self.time_frame, width=12)
        self.start_entry.insert(0, "0:00")
        self.start_entry.grid(row=0, column=1, padx=(8, 8), pady=2)
        self.start_entry.bind("<FocusOut>", self._sync_start_from_entry)
        tk.Button(self.time_frame, text="+0.1", width=5, command=self._start_add).grid(row=0, column=2, padx=2, pady=2)
        tk.Button(self.time_frame, text="-0.1", width=5, command=self._start_sub).grid(row=0, column=3, padx=2, pady=2)

        self.end_label_widget = tk.Label(self.time_frame, text="End:")
        self.end_label_widget.grid(row=1, column=0, sticky="w", pady=2)
        self.end_entry = tk.Entry(self.time_frame, width=12)
        self.end_entry.insert(0, "0:00")
        self.end_entry.grid(row=1, column=1, padx=(8, 8), pady=2)
        self.end_entry.bind("<FocusOut>", self._sync_end_from_entry)
        tk.Button(self.time_frame, text="+0.1", width=5, command=self._end_add).grid(row=1, column=2, padx=2, pady=2)
        tk.Button(self.time_frame, text="-0.1", width=5, command=self._end_sub).grid(row=1, column=3, padx=2, pady=2)

        # Add to queue & Process
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)
        self.add_btn = tk.Button(btn_frame, text="Add to queue", command=self._add_to_queue,
                                 bg="#2196F3", fg="white", cursor="hand2", padx=12, pady=4)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.process_btn = tk.Button(btn_frame, text="Process queue", command=self._process_queue,
                                    bg="#4CAF50", fg="white", font=("", 10, "bold"), cursor="hand2", padx=16, pady=4)
        self.process_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.clear_queue_btn = tk.Button(btn_frame, text="Clear queue", command=self._clear_queue, padx=12, pady=4)
        self.clear_queue_btn.pack(side=tk.LEFT)

        # Queue list
        queue_frame = tk.LabelFrame(self.root, text="Queue", padx=8, pady=6)
        queue_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.queue_listbox = tk.Listbox(queue_frame, height=6, font=("Consolas", 9))
        self.queue_listbox.pack(fill=tk.BOTH, expand=True)

        # Progress (like video_download)
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X)
        self.progress_label = tk.Label(progress_frame, text="0%", fg="gray")
        self.progress_label.pack(anchor="e")

        self.status_label = tk.Label(self.root, textvariable=self.status_var, fg="gray", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

        self._on_mode_change()
        self._sync_entries_and_labels()

    def _on_mode_change(self):
        """Update labels and frame title based on trim vs cut_middle mode."""
        if self.mode_var.get() == "cut_middle":
            self.time_frame.config(text="Section to remove (mm:ss.s or +/- 0.1s)")
            self.start_label_widget.config(text="Cut from:")
            self.end_label_widget.config(text="Cut to:")
        else:
            self.time_frame.config(text="Trim range (mm:ss.s or +/- 0.1s)")
            self.start_label_widget.config(text="Start:")
            self.end_label_widget.config(text="End:")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select video file",
            filetypes=[("Video files", "*.mp4 *.mov *.mkv *.webm *.avi")]
        )
        if path:
            self.video_path.set(path)
            self.duration_sec = get_video_duration(path)
            if self.duration_sec is not None:
                self.duration_label.config(text=f"Duration: {seconds_to_time_str(self.duration_sec)} ({self.duration_sec:.1f} sec)")
                if self.mode_var.get() == "cut_middle":
                    # Default: suggest cutting 2 seconds from the middle
                    mid = self.duration_sec / 2
                    self.start_sec.set(max(0, mid - 1))
                    self.end_sec.set(min(self.duration_sec, mid + 1))
                else:
                    self.start_sec.set(0.0)
                    self.end_sec.set(self.duration_sec)
                self._sync_entries_and_labels()
            else:
                self.duration_label.config(text="Duration: (could not detect)")
                self.duration_sec = 0.0
            self.status_var.set("Ready")

    def _sync_start_from_entry(self, event=None):
        try:
            s = time_to_seconds(self.start_entry.get())
            s = max(0, min(s, self.duration_sec or s))
            self.start_sec.set(s)
            self.start_entry.delete(0, tk.END)
            self.start_entry.insert(0, seconds_to_time_str(s))
        except ValueError:
            pass

    def _sync_end_from_entry(self, event=None):
        try:
            s = time_to_seconds(self.end_entry.get())
            s = max(0, min(s, self.duration_sec or s))
            self.end_sec.set(s)
            self.end_entry.delete(0, tk.END)
            self.end_entry.insert(0, seconds_to_time_str(s))
        except ValueError:
            pass

    def _start_add(self):
        s = round(self.start_sec.get() + 0.1, 1)
        max_start = (self.end_sec.get() - 0.1) if self.duration_sec else s
        s = min(s, self.duration_sec or s, max_start)
        self.start_sec.set(s)
        self.start_entry.delete(0, tk.END)
        self.start_entry.insert(0, seconds_to_time_str(s))

    def _start_sub(self):
        s = round(self.start_sec.get() - 0.1, 1)
        s = max(0, s)
        self.start_sec.set(s)
        self.start_entry.delete(0, tk.END)
        self.start_entry.insert(0, seconds_to_time_str(s))

    def _end_add(self):
        s = round(self.end_sec.get() + 0.1, 1)
        s = min(s, self.duration_sec or s)
        self.end_sec.set(s)
        self.end_entry.delete(0, tk.END)
        self.end_entry.insert(0, seconds_to_time_str(s))

    def _end_sub(self):
        s = round(self.end_sec.get() - 0.1, 1)
        min_end = self.start_sec.get() + 0.1
        s = max(min_end, s)
        self.end_sec.set(s)
        self.end_entry.delete(0, tk.END)
        self.end_entry.insert(0, seconds_to_time_str(s))

    def _sync_entries_and_labels(self):
        self.start_entry.delete(0, tk.END)
        self.start_entry.insert(0, seconds_to_time_str(self.start_sec.get()))
        self.end_entry.delete(0, tk.END)
        self.end_entry.insert(0, seconds_to_time_str(self.end_sec.get()))

    def _add_to_queue(self):
        path = self.video_path.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("No video", "Please select a video file first.")
            return
        try:
            start = time_to_seconds(self.start_entry.get())
            end = time_to_seconds(self.end_entry.get())
        except ValueError as e:
            messagebox.showerror("Invalid time", str(e))
            return
        if end <= start:
            messagebox.showerror("Invalid range", "End must be after start.")
            return
        if self.duration_sec and (start < 0 or end > self.duration_sec):
            messagebox.showerror("Out of range", f"Start/end must be within 0 and {seconds_to_time_str(self.duration_sec)}.")
            return
        mode = self.mode_var.get()
        self.queue.append((path, start, end, mode))
        name = os.path.basename(path)
        op = "cut" if mode == "cut_middle" else "trim"
        self.queue_listbox.insert(tk.END, f"[{op}] {name}  {seconds_to_time_str(start)} → {seconds_to_time_str(end)}")
        self.status_var.set(f"Added. Queue: {len(self.queue)} job(s)")

    def _clear_queue(self):
        self.queue.clear()
        self.queue_listbox.delete(0, tk.END)
        self.status_var.set("Queue cleared")

    def _update_progress(self, percent):
        self.progress_bar["value"] = percent
        self.progress_label.config(text=f"{percent:.1f}%")
        self.root.update_idletasks()

    def _process_queue(self):
        if not self.queue:
            messagebox.showwarning("Empty queue", "Add trim jobs to the queue first.")
            return
        if self.processing:
            return
        self.processing = True
        self.add_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)

        def worker():
            total = len(self.queue)
            for i, item in enumerate(self.queue, 1):
                path, start, end = item[0], item[1], item[2]
                mode = item[3] if len(item) > 3 else "trim"
                def status_cb(msg):
                    self.root.after(0, lambda m=msg: self.status_var.set(m))
                def progress_cb(pct):
                    overall = (i - 1) / total * 100 + (pct / 100) / total * 100
                    self.root.after(0, lambda o=overall: self._update_progress(o))
                op = "Cut" if mode == "cut_middle" else "Trim"
                self.root.after(0, lambda: self.status_var.set(f"Processing {i}/{total} ({op}): {os.path.basename(path)}"))
                try:
                    if mode == "cut_middle":
                        cut_middle_video(path, start, end, progress_callback=progress_cb, status_callback=status_cb)
                    else:
                        trim_video(path, start, end, progress_callback=progress_cb, status_callback=status_cb)
                except Exception as e:
                    self.root.after(0, lambda err=str(e): messagebox.showerror("Error", err))
            self.root.after(0, self._process_complete)

        threading.Thread(target=worker, daemon=True).start()

    def _process_complete(self):
        self.processing = False
        self.add_btn.config(state=tk.NORMAL)
        self.process_btn.config(state=tk.NORMAL)
        self._update_progress(100)
        self.status_var.set("All done!")
        messagebox.showinfo("Done", "All trims complete.")
        self._clear_queue()

    def run(self):
        self.root.mainloop()


# -------------------------------------------------------
# CLI entry
# -------------------------------------------------------

def cli_trim(file_path, start_str, end_str):
    """Run trim from command line (no GUI)."""
    try:
        start_sec = time_to_seconds(start_str)
        end_sec = time_to_seconds(end_str)
        log.info("Trimming %s from %s to %s...", file_path, start_str, end_str)
        trim_video(file_path, start_sec, end_sec)
        log.info("Done.")
    except Exception as e:
        log.error("Error: %s", e)
        exit(1)


def cli_cut_middle(file_path, cut_start_str, cut_end_str):
    """Run cut middle from command line (no GUI)."""
    try:
        cut_start_sec = time_to_seconds(cut_start_str)
        cut_end_sec = time_to_seconds(cut_end_str)
        log.info("Cutting %s to %s from %s...", cut_start_str, cut_end_str, file_path)
        cut_middle_video(file_path, cut_start_sec, cut_end_sec)
        log.info("Done.")
    except Exception as e:
        log.error("Error: %s", e)
        exit(1)


def choose_file():
    """CLI helper: open file dialog when no file specified."""
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select video file",
        filetypes=[("Video files", "*.mp4 *.mov *.mkv *.webm *.avi")]
    )
    root.destroy()
    if not path:
        log.error("No file selected.")
        exit(1)
    return path


def ask_time(prompt):
    """CLI helper: ask for time via simpledialog."""
    root = tk.Tk()
    root.withdraw()
    from tkinter import simpledialog
    while True:
        val = simpledialog.askstring("Input", prompt + " (mm:ss or mm:ss.s or hh:mm:ss)")
        if val is None:
            exit(1)
        try:
            time_to_seconds(val)
            root.destroy()
            return val
        except ValueError:
            messagebox.showerror("Error", "Invalid format. Use mm:ss or hh:mm:ss (fractions like 1:30.5 ok).")
    root.destroy()


# -------------------------------------------------------
# Main
# -------------------------------------------------------

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Video Trimmer - trim or cut middle section with FFmpeg")
    parser.add_argument("-f", "--file", help="Video file path (CLI mode)")
    parser.add_argument("-s", "--start", help="Start time (mm:ss or mm:ss.s)")
    parser.add_argument("-e", "--end", help="End time (mm:ss or mm:ss.s)")
    parser.add_argument("--cut", action="store_true", help="Cut middle mode: remove start→end section and join")
    parser.add_argument("--gui", action="store_true", help="Force GUI mode")
    args = parser.parse_args()

    if args.gui or (not args.file and not args.start and not args.end):
        app = VideoTrimApp()
        app.run()
    else:
        if not args.file:
            args.file = choose_file()
        if not args.start:
            args.start = ask_time("Enter cut START time (start of section to remove)" if args.cut else "Enter START time")
        if not args.end:
            args.end = ask_time("Enter cut END time (end of section to remove)" if args.cut else "Enter END time")
        if args.cut:
            cli_cut_middle(args.file, args.start, args.end)
        else:
            cli_trim(args.file, args.start, args.end)
