"""Main-window GUI.

Shows server status with a traffic light, start/stop buttons, language
toggle, and a big "Record" button. Recording uses the shared popup. Window
close can optionally hide to tray (caller-controlled flag).
"""

from __future__ import annotations

# Standard library imports
import logging
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

# Third-party imports
import pyperclip

from ..core import (
    AppConfig,
    AudioRecorder,
    TranscriptionClient,
    TranscriptionError,
)
from ..whisper_server import OWNERSHIP_OURS, WhisperServerManager
from .recording_popup import RecordingPopup

logger = logging.getLogger(__name__)

POLL_MS = 2000


class TranscriberApp:
    def __init__(self, config: AppConfig, tray_on_close: bool = False) -> None:
        self.config = config
        self.tray_on_close = tray_on_close
        self.server = WhisperServerManager()

        self.root = tk.Tk()
        self.root.title("Voice Transcription")
        self.root.geometry("420x340")
        self.root.resizable(False, False)
        self.root.configure(background="#2E2E2E")

        self.status_var = tk.StringVar(value="checking…")
        self.language_var = tk.StringVar(value=config.language)
        self.translate_var = tk.BooleanVar(value=config.translate)

        self._build_widgets()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._poll_status()

    # ------------------------------------------------------------ layout

    def _build_widgets(self) -> None:
        pad = {"padx": 16, "pady": 6}

        title = ttk.Label(self.root, text="Voice Transcription", font=("Segoe UI", 14, "bold"))
        title.pack(pady=(14, 4))

        # Server status row
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, **pad)
        ttk.Label(status_frame, text="Server:").pack(side=tk.LEFT)
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, font=("Segoe UI", 10, "bold"))
        self.status_label.pack(side=tk.LEFT, padx=8)

        server_btn_frame = ttk.Frame(self.root)
        server_btn_frame.pack(fill=tk.X, **pad)
        self.start_btn = ttk.Button(server_btn_frame, text="▶ Start server", command=self._start_server)
        self.start_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        self.stop_btn = ttk.Button(server_btn_frame, text="■ Stop server", command=self._stop_server)
        self.stop_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(4, 0))

        # Language + translate
        lang_frame = ttk.Frame(self.root)
        lang_frame.pack(fill=tk.X, **pad)
        ttk.Label(lang_frame, text="Language:").pack(side=tk.LEFT)
        lang_combo = ttk.Combobox(
            lang_frame, textvariable=self.language_var, state="readonly", width=12,
            values=("Spanish", "English", "Italian", "French", "German", "Portuguese", "auto"),
        )
        lang_combo.pack(side=tk.LEFT, padx=8)
        ttk.Checkbutton(lang_frame, text="Translate to English", variable=self.translate_var).pack(side=tk.LEFT, padx=8)

        # Primary actions
        record_btn = ttk.Button(self.root, text="🎤 Record", command=self._record)
        record_btn.pack(fill=tk.X, **pad)
        record_btn.configure(padding=(0, 10))

        file_btn = ttk.Button(self.root, text="📁 Transcribe file…", command=self._transcribe_file_dialog)
        file_btn.pack(fill=tk.X, **pad)

        quit_btn = ttk.Button(self.root, text="Quit", command=self._quit)
        quit_btn.pack(fill=tk.X, **pad)

    # --------------------------------------------------- server status polling

    def _poll_status(self) -> None:
        status = self.server.status()
        if status.running and status.ownership == OWNERSHIP_OURS:
            self.status_var.set(f"🟢 running (ours) :{status.port}")
        elif status.running:
            self.status_var.set(f"🟢 running (external) :{status.port}")
        else:
            self.status_var.set(f"🔴 not running :{status.port}")

        self.start_btn.state(["disabled"] if status.running else ["!disabled"])
        self.stop_btn.state(["!disabled"] if status.running and status.ownership == OWNERSHIP_OURS else ["disabled"])

        self.root.after(POLL_MS, self._poll_status)

    def _start_server(self) -> None:
        self.status_var.set("⏳ starting…")
        threading.Thread(target=self._start_server_worker, daemon=True).start()

    def _start_server_worker(self) -> None:
        try:
            self.server.start()
        except RuntimeError as e:
            logger.error(str(e))
            self.root.after(0, lambda: messagebox.showerror("Server failed to start", str(e)))

    def _stop_server(self) -> None:
        threading.Thread(target=self.server.stop, daemon=True).start()

    # ---------------------------------------------------------- record flow

    def _record(self) -> None:
        status = self.server.status()
        if not status.running:
            messagebox.showwarning(
                "Server not running",
                "Start the whisper-server first (▶ Start server).",
            )
            return

        recorder = AudioRecorder(
            sample_rate=self.config.sample_rate,
            preferred_mics=self.config.resolve_preferred_mics(),
        )
        RecordingPopup(
            parent=self.root,
            recorder=recorder,
            max_seconds=self.config.max_record_seconds,
            on_done=self._on_record_done,
        )

    def _on_record_done(self, recording, error) -> None:
        if error is not None:
            messagebox.showerror("Recording error", error)
            return
        if recording is None:
            return
        threading.Thread(
            target=self._transcribe_and_show,
            args=(recording,),
            daemon=True,
        ).start()

    def _transcribe_and_show(self, recording) -> None:
        status = self.server.status()
        client = TranscriptionClient(status.base_url)
        try:
            text = client.transcribe_array(
                recording.samples, recording.sample_rate,
                language=self.language_var.get(),
                translate=self.translate_var.get(),
            )
        except TranscriptionError as e:
            logger.error(f"❌ {e}")
            self.root.after(0, lambda: messagebox.showerror("Transcription failed", str(e)))
            return

        if self.config.auto_copy:
            try:
                pyperclip.copy(text)
            except Exception as exc:
                logger.warning(f"⚠️  Clipboard copy failed: {exc}")

        self.root.after(0, lambda: self._show_result(text))

    def _show_result(self, text: str) -> None:
        win = tk.Toplevel(self.root)
        win.title("Transcription")
        win.geometry("640x360")
        win.transient(self.root)

        text_widget = tk.Text(win, wrap=tk.WORD)
        text_widget.insert(tk.END, text)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        btns = ttk.Frame(win)
        btns.pack(fill=tk.X, padx=10, pady=(0, 10))

        def copy() -> None:
            pyperclip.copy(text)
            copy_btn.config(text="✓ Copied")

        copy_btn = ttk.Button(btns, text="Copy", command=copy)
        copy_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(btns, text="Close", command=win.destroy).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

    # ---------------------------------------------------------- file flow

    def _transcribe_file_dialog(self) -> None:
        status = self.server.status()
        if not status.running:
            messagebox.showwarning("Server not running", "Start the whisper-server first.")
            return
        path = filedialog.askopenfilename(
            title="Select audio file",
            filetypes=(
                ("Audio", "*.mp3 *.wav *.m4a *.flac *.ogg *.mp4 *.webm"),
                ("All files", "*.*"),
            ),
        )
        if not path:
            return
        threading.Thread(target=self._transcribe_file_worker, args=(path,), daemon=True).start()

    def _transcribe_file_worker(self, path: str) -> None:
        status = self.server.status()
        client = TranscriptionClient(status.base_url)
        try:
            text = client.transcribe_file(
                path,
                language=self.language_var.get(),
                translate=self.translate_var.get(),
            )
        except TranscriptionError as e:
            self.root.after(0, lambda: messagebox.showerror("Transcription failed", str(e)))
            return
        if self.config.auto_copy:
            try:
                pyperclip.copy(text)
            except Exception:
                pass
        self.root.after(0, lambda: self._show_result(text))

    # ------------------------------------------------------------ lifecycle

    def _on_close(self) -> None:
        if self.tray_on_close:
            self.root.withdraw()
            return
        self._quit()

    def _quit(self) -> None:
        status = self.server.status()
        if status.running and status.ownership == OWNERSHIP_OURS:
            self.server.stop()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def run(self) -> int:
        self.root.mainloop()
        return 0


def run_gui(config: AppConfig, tray_on_close: bool = False) -> int:
    app = TranscriberApp(config, tray_on_close=tray_on_close)
    return app.run()
