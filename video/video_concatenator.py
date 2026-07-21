"""
Video Concatenator - Merge multiple videos into one using FFmpeg.
Supports both GUI and command-line interfaces. Uses stream copy (no re-encode) by default.
"""

import argparse
import logging
import os
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from _ffmpeg_utils import get_video_duration

log = logging.getLogger(__name__)


class VideoConcatenator:
    """Class to handle video concatenation operations using ffmpeg."""

    def concatenate_videos(
        self,
        video_paths: List[str],
        output_path: str,
        progress_callback=None,
        status_callback=None,
        sort_alphabetically: bool = False,
    ) -> bool:
        """
        Concatenate multiple video files into a single output video using ffmpeg.

        Args:
            video_paths: List of paths to video files to concatenate (order preserved unless sort_alphabetically)
            output_path: Path where the output video will be saved
            progress_callback: Optional callback(percent) for progress updates
            status_callback: Optional callback(message) for status updates
            sort_alphabetically: If True, sort paths alphabetically; else use list order

        Returns:
            bool: True if successful, False otherwise
        """
        for path in video_paths:
            if not os.path.exists(path):
                if status_callback:
                    status_callback(f"Video file not found: {path}")
                return False

        paths = sorted(video_paths) if sort_alphabetically else list(video_paths)
        fd, list_file = tempfile.mkstemp(suffix=".txt", prefix="video_concat_")
        os.close(fd)

        try:
            abs_paths = [os.path.abspath(p).replace("\\", "/") for p in paths]
            with open(list_file, "w") as f:
                for p in abs_paths:
                    f.write(f"file '{p}'\n")

            if status_callback:
                status_callback(f"Concatenating {len(paths)} videos...")

            # Probe each input's duration up front so progress can be computed
            # as (elapsed output time / total input duration) instead of the
            # old `elapsed * 2` guess, which hit 99% almost immediately on
            # long concatenations (audit issue #67). Stream-copy concat's
            # output duration is the sum of the input durations.
            total_duration: Optional[float] = 0.0
            for p in abs_paths:
                duration = get_video_duration(p)
                if duration is None:
                    total_duration = None
                    break
                total_duration += duration

            cmd = [
                "ffmpeg",
                "-f", "concat",
                "-safe", "0",
                "-i", list_file,
                "-c", "copy",
                "-map", "0",
                "-y",
                output_path,
            ]

            # Run with progress parsing for concat (ffmpeg may output time=)
            process = subprocess.Popen(
                cmd,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
            time_pat = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")
            for line in process.stderr:
                if progress_callback and "time=" in line:
                    m = time_pat.search(line)
                    if m:
                        h, mn, s, ms = map(int, m.groups())
                        elapsed = h * 3600 + mn * 60 + s + ms / 100.0
                        if total_duration:
                            percent = (elapsed / total_duration) * 100.0
                        else:
                            # Could not probe durations upfront; fall back to
                            # the previous rough estimate.
                            percent = elapsed * 2
                        progress_callback(min(99.0, percent))
            process.wait()

            if os.path.exists(list_file):
                os.remove(list_file)

            if process.returncode == 0:
                if progress_callback:
                    progress_callback(100.0)
                if status_callback:
                    status_callback(f"Done: {output_path}")
                return True
            else:
                # Try filter_complex fallback
                return self._concatenate_with_filter_complex(
                    paths, output_path, progress_callback, status_callback
                )

        except FileNotFoundError:
            if status_callback:
                status_callback("FFmpeg not found. Add it to your PATH.")
            return False
        except Exception as e:
            if os.path.exists(list_file):
                try:
                    os.remove(list_file)
                except OSError:
                    pass
            if status_callback:
                status_callback(str(e))
            return False

    def _concatenate_with_filter_complex(
        self,
        video_paths: List[str],
        output_path: str,
        progress_callback=None,
        status_callback=None,
    ) -> bool:
        """Alternative concatenation using filter_complex (re-encodes)."""
        try:
            if status_callback:
                status_callback("Trying alternative method (re-encoding)...")

            inputs = []
            filter_parts = []
            for i in range(len(video_paths)):
                inputs.extend(["-i", video_paths[i]])
                filter_parts.append(f"[{i}:v][{i}:a]")
            filter_complex = "".join(filter_parts) + f"concat=n={len(video_paths)}:v=1:a=1[outv][outa]"

            cmd = [
                "ffmpeg",
                *inputs,
                "-filter_complex", filter_complex,
                "-map", "[outv]",
                "-map", "[outa]",
                "-c:v", "libx264",
                "-c:a", "aac",
                "-y",
                output_path,
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
            )
            if result.returncode == 0:
                if progress_callback:
                    progress_callback(100.0)
                if status_callback:
                    status_callback(f"Done: {output_path}")
                return True
            if status_callback:
                status_callback(result.stderr[:200] if result.stderr else "FFmpeg failed")
            return False
        except Exception as e:
            if status_callback:
                status_callback(str(e))
            return False


# -----------------------------------------------------------------------------
# GUI Application
# -----------------------------------------------------------------------------

class VideoConcatenatorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Video Concatenator")
        self.root.geometry("640x480")
        self.root.minsize(500, 400)

        self.video_paths: List[str] = []
        self.output_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Add videos and choose output")
        self.processing = False

        self._build_ui()

    def _build_ui(self):
        # Video list
        list_frame = tk.LabelFrame(self.root, text="Videos to concatenate (order matters)", padx=8, pady=6)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        btn_row = tk.Frame(list_frame)
        btn_row.pack(fill=tk.X)
        tk.Button(btn_row, text="Add videos...", command=self._add_videos).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Remove selected", command=self._remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Move up", command=self._move_up).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Move down", command=self._move_down).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_row, text="Clear all", command=self._clear_list).pack(side=tk.LEFT)

        self.listbox = tk.Listbox(list_frame, height=10, font=("Consolas", 9))
        self.listbox.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        # Output
        out_frame = tk.Frame(self.root)
        out_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(out_frame, text="Output file:").pack(side=tk.LEFT, padx=(0, 8))
        tk.Entry(out_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Button(out_frame, text="Browse...", command=self._browse_output).pack(side=tk.LEFT)

        # Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=8)
        self.process_btn = tk.Button(
            btn_frame, text="Concatenate", command=self._start_process,
            bg="#4CAF50", fg="white", font=("", 10, "bold"), cursor="hand2", padx=20, pady=5
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 8))

        # Progress
        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(progress_frame, maximum=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X)
        self.progress_label = tk.Label(progress_frame, text="0%", fg="gray")
        self.progress_label.pack(anchor="e")

        self.status_label = tk.Label(self.root, textvariable=self.status_var, fg="gray", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

    def _add_videos(self):
        paths = filedialog.askopenfilenames(
            title="Select video files",
            filetypes=[
                ("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm"),
                ("All files", "*.*"),
            ],
        )
        for p in paths:
            if p not in self.video_paths:
                self.video_paths.append(p)
                self.listbox.insert(tk.END, os.path.basename(p))
        self.status_var.set(f"{len(self.video_paths)} video(s) loaded")

    def _remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.video_paths[idx]
        self.listbox.delete(idx)
        self.status_var.set(f"{len(self.video_paths)} video(s)")

    def _move_up(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self.video_paths[idx], self.video_paths[idx - 1] = self.video_paths[idx - 1], self.video_paths[idx]
        self.listbox.delete(idx)
        self.listbox.insert(idx - 1, os.path.basename(self.video_paths[idx - 1]))
        self.listbox.selection_set(idx - 1)

    def _move_down(self):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.video_paths) - 1:
            return
        idx = sel[0]
        self.video_paths[idx], self.video_paths[idx + 1] = self.video_paths[idx + 1], self.video_paths[idx]
        self.listbox.delete(idx)
        self.listbox.insert(idx + 1, os.path.basename(self.video_paths[idx + 1]))
        self.listbox.selection_set(idx + 1)

    def _clear_list(self):
        self.video_paths.clear()
        self.listbox.delete(0, tk.END)
        self.status_var.set("Add videos")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save concatenated video as",
            defaultextension=".mp4",
            filetypes=[("MP4 files", "*.mp4"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _update_progress(self, percent):
        self.progress_bar["value"] = percent
        self.progress_label.config(text=f"{percent:.1f}%")
        self.root.update_idletasks()

    def _start_process(self):
        if len(self.video_paths) < 2:
            messagebox.showwarning("Not enough videos", "Add at least 2 videos to concatenate.")
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showwarning("No output", "Choose an output file path.")
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

            ok = VideoConcatenator().concatenate_videos(
                self.video_paths, out,
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
            messagebox.showinfo("Success", f"Video saved to:\n{self.output_path.get()}")
        else:
            messagebox.showerror("Error", "Concatenation failed. Check status for details.")

    def run(self):
        self.root.mainloop()


# -----------------------------------------------------------------------------
# CLI helpers and main
# -----------------------------------------------------------------------------

def select_video_files() -> List[str]:
    root = tk.Tk()
    root.withdraw()
    paths = filedialog.askopenfilenames(
        title="Select video files to concatenate",
        filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm"), ("All files", "*.*")],
    )
    root.destroy()
    return list(paths)


def select_output_file() -> str:
    root = tk.Tk()
    root.withdraw()
    path = filedialog.asksaveasfilename(
        title="Save concatenated video as",
        defaultextension=".mp4",
        filetypes=[("MP4 files", "*.mp4"), ("All files", "*.*")],
    )
    root.destroy()
    return path


def main_cli(video_paths: List[str], output_path: str, sort: bool = False) -> bool:
    return VideoConcatenator().concatenate_videos(video_paths, output_path, sort_alphabetically=sort)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Video Concatenator - merge videos with FFmpeg")
    parser.add_argument("-o", "--output", help="Output file path (CLI mode)")
    parser.add_argument("files", nargs="*", help="Input video files (CLI mode)")
    parser.add_argument("--sort", action="store_true", help="Sort input files alphabetically")
    parser.add_argument("--gui", action="store_true", help="Force GUI mode")
    args = parser.parse_args()

    if args.gui or (not args.output and not args.files):
        app = VideoConcatenatorApp()
        app.run()
    else:
        if not args.files:
            args.files = select_video_files()
        if not args.output:
            args.output = select_output_file()
        if not args.files or len(args.files) < 2:
            log.error("Need at least 2 video files.")
            exit(1)
        if not args.output:
            log.error("No output path specified.")
            exit(1)
        ok = main_cli(args.files, args.output, sort=args.sort)
        exit(0 if ok else 1)
