"""
Screen Recorder GUI – adjustable FPS and mouse overlay
======================================================
Captures screen with mouse pointer overlay using `mss` + `pyautogui.position()`.
• Multi-monitor support (choose from dropdown)
• Always captures mouse (white circle with black outline)
• Adjustable FPS from the GUI (default 5 FPS)

Dependencies:
    pip install opencv-python mss pyautogui screeninfo numpy

Notes:
• Lower FPS → smaller files but choppier video.
• Cursor overlay ensures visibility on light/dark backgrounds.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import cv2
import numpy as np
import pyautogui
from datetime import datetime
import threading
import time
from pathlib import Path
from screeninfo import get_monitors
import mss

# --- CONFIGURATION -------------------------------------------------------
OUTPUT_DIR = Path("C:/Users/u0150867/Downloads/videos")
DEFAULT_FPS = 5
CODEC = cv2.VideoWriter_fourcc(*"mp4v")
MAX_FILE_MB: int | None = None
# -------------------------------------------------------------------------

class ScreenRecorderGUI:
    def __init__(self, master: tk.Tk):
        self.master = master
        master.title("Screen Recorder")

        self.monitors = get_monitors()
        if not self.monitors:
            messagebox.showerror("Error", "No monitors detected.")
            master.destroy()
            return

        monitor_options = [
            f"{idx}: {m.width}x{m.height} @ ({m.x},{m.y})" for idx, m in enumerate(self.monitors)
        ]

        # --- Monitor selector ---
        mon_frame = tk.Frame(master)
        mon_frame.pack(pady=(10, 4))

        tk.Label(mon_frame, text="Screen:").pack(side="left", padx=(0, 6))
        self.monitor_var = tk.StringVar(value=monitor_options[0])
        self.monitor_cb = ttk.Combobox(
            mon_frame,
            textvariable=self.monitor_var,
            values=monitor_options,
            width=45,
            state="readonly",
        )
        self.monitor_cb.pack(side="left")

        # --- FPS selector ---
        fps_frame = tk.Frame(master)
        fps_frame.pack(pady=(2, 2))

        tk.Label(fps_frame, text="Frames per second:").pack(side="left", padx=(0, 6))
        self.fps_var = tk.IntVar(value=DEFAULT_FPS)
        self.fps_entry = tk.Entry(fps_frame, textvariable=self.fps_var, width=5)
        self.fps_entry.pack(side="left")

        # --- Controls ---
        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=8)

        self.start_stop_btn = tk.Button(btn_frame, text="Start", width=12, command=self.toggle_recording)
        self.start_stop_btn.pack(side="left", padx=12)

        self.pause_resume_btn = tk.Button(btn_frame, text="Pause", width=12, command=self.toggle_pause, state=tk.DISABLED)
        self.pause_resume_btn.pack(side="left", padx=12)

        self.quit_btn = tk.Button(btn_frame, text="Quit", width=12, command=self.quit_app)
        self.quit_btn.pack(side="left", padx=12)

        # --- Status ---
        self.status_var = tk.StringVar(value="Idle")
        self.status_label = tk.Label(master, textvariable=self.status_var, anchor="w", justify="left", wraplength=720)
        self.status_label.pack(fill="x", padx=10, pady=(2, 10))

        # --- Internal state ---
        self._recording_event = threading.Event()
        self._paused = False
        self._thread = None
        self._writer = None
        self._output_path = None
        self._segment_index = 1
        self._bbox = None
        self._frame_dim = None
        self._fps = DEFAULT_FPS
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    def toggle_recording(self):
        if not self._recording_event.is_set():
            self.start_recording()
        else:
            self.stop_recording()

    def toggle_pause(self):
        if not self._recording_event.is_set():
            return
        self._paused = not self._paused
        self.pause_resume_btn.configure(text="Resume" if self._paused else "Pause")
        state_text = "Paused" if self._paused else "Recording..."
        self.status_var.set(f"{state_text} → {self._output_path}")

    def quit_app(self):
        if self._recording_event.is_set():
            self.stop_recording()
        self.master.destroy()

    def _generate_output_path(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        sfx = f"_{self._segment_index:02d}" if self._segment_index > 1 else ""
        return OUTPUT_DIR / f"screen_record_{ts}{sfx}.mp4"

    def start_recording(self):
        idx = int(self.monitor_var.get().split(":")[0])
        mon = self.monitors[idx]
        self._bbox = {"left": mon.x, "top": mon.y, "width": mon.width, "height": mon.height}
        self._frame_dim = (mon.width, mon.height)

        try:
            fps_value = int(self.fps_var.get())
            if fps_value <= 0:
                raise ValueError
            self._fps = fps_value
        except ValueError:
            messagebox.showerror("Invalid FPS", "Please enter a positive integer for FPS.")
            return

        self._segment_index = 1
        self._output_path = self._generate_output_path()
        self._writer = cv2.VideoWriter(str(self._output_path), CODEC, self._fps, self._frame_dim)
        if not self._writer.isOpened():
            messagebox.showerror("Error", "Failed to initialize video writer.")
            return

        self._recording_event.set()
        self._paused = False
        self._thread = threading.Thread(target=self._record, daemon=True)
        self._thread.start()

        self.start_stop_btn.configure(text="Stop")
        self.pause_resume_btn.configure(state=tk.NORMAL, text="Pause")
        self.monitor_cb.configure(state="disabled")
        self.fps_entry.configure(state="disabled")
        self.status_var.set(f"Recording... → {self._output_path}")

    def stop_recording(self):
        if self._recording_event.is_set():
            self._recording_event.clear()
            if self._thread:
                self._thread.join()
                self._thread = None

        if self._writer:
            self._writer.release()
            self._writer = None

        self.start_stop_btn.configure(text="Start")
        self.pause_resume_btn.configure(state=tk.DISABLED, text="Pause")
        self.monitor_cb.configure(state="readonly")
        self.fps_entry.configure(state="normal")
        self.status_var.set(f"Saved: {self._output_path}")
        messagebox.showinfo("Recording finished", f"Saved to:\n{self._output_path}")

    def _record(self):
        try:
            with mss.mss() as sct:
                frame_interval = 1 / self._fps
                while self._recording_event.is_set():
                    start_time = time.time()

                    if self._paused:
                        time.sleep(0.1)
                        continue

                    img = np.array(sct.grab(self._bbox))
                    frame = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

                    mouse_x, mouse_y = pyautogui.position()
                    rel_x = mouse_x - self._bbox["left"]
                    rel_y = mouse_y - self._bbox["top"]
                    if 0 <= rel_x < self._frame_dim[0] and 0 <= rel_y < self._frame_dim[1]:
                        cv2.circle(frame, (rel_x, rel_y), 10, (0, 0, 0), 3)
                        cv2.circle(frame, (rel_x, rel_y), 6, (255, 255, 255), -1)

                    self._writer.write(frame)
                    self._check_split()

                    elapsed = time.time() - start_time
                    sleep_time = frame_interval - elapsed
                    if sleep_time > 0:
                        time.sleep(sleep_time)
        except Exception as e:
            self.master.after(0, lambda: messagebox.showerror("Recording Error", str(e)))
            self.master.after(0, self.stop_recording)

    def _check_split(self):
        if MAX_FILE_MB is not None and self._output_path.stat().st_size / (1024 * 1024) >= MAX_FILE_MB:
            self._writer.release()
            self._segment_index += 1
            self._output_path = self._generate_output_path()
            self._writer = cv2.VideoWriter(str(self._output_path), CODEC, self._fps, self._frame_dim)
            self.status_var.set(f"Recording... → {self._output_path}")


# --- Entry point ---
def main():
    root = tk.Tk()
    root.geometry("520x140")
    root.resizable(True, True)
    ScreenRecorderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()