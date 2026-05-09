"""System-tray wrapper for the Flask launcher server.

Run:
    python launcher/tray.py

The Flask server starts automatically in a background thread.
Double-click the tray icon (or choose Open) to see status and logs.
Quit via the tray menu to shut down cleanly.
"""

from __future__ import annotations

import logging
import queue
import socket
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import ttk
from typing import Optional

import pystray
from PIL import Image, ImageDraw
from werkzeug.serving import make_server

# launcher.py is in the same directory; Python adds the script's directory
# to sys.path automatically, so this import works when run as:
#   python launcher/tray.py   (from the repo root)
sys.path.insert(0, str(Path(__file__).parent))
from launcher import HOST, PORT, SSL_CONTEXT, USE_HTTPS, app, log  # noqa: E402


# ---------------------------------------------------------------------------
# Log capture
# ---------------------------------------------------------------------------

class _QueueHandler(logging.Handler):
    """Routes log records into a thread-safe queue for the UI to drain."""

    def __init__(self) -> None:
        super().__init__()
        self._q: queue.Queue[str] = queue.Queue()

    def emit(self, record: logging.LogRecord) -> None:
        self._q.put(self.format(record))

    def drain(self) -> list[str]:
        lines: list[str] = []
        try:
            while True:
                lines.append(self._q.get_nowait())
        except queue.Empty:
            pass
        return lines


_LOG_FMT = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")


# ---------------------------------------------------------------------------
# Tray application
# ---------------------------------------------------------------------------

class TrayApp:
    _ICON_SIZE = 64

    def __init__(self) -> None:
        self._log_handler = _QueueHandler()
        self._log_handler.setFormatter(_LOG_FMT)
        logging.getLogger().addHandler(self._log_handler)

        self._log_lines: list[str] = []
        self._server = None
        self._root: Optional[tk.Tk] = None
        self._window: Optional[tk.Toplevel] = None
        self._log_text: Optional[tk.Text] = None

        self._icon = pystray.Icon(
            "launcher",
            self._make_icon(running=False),
            "Launcher",
            pystray.Menu(
                pystray.MenuItem("Open", self._on_open_browser, default=True),
                pystray.MenuItem("Show logs", self._on_show_logs),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._on_quit),
            ),
        )

    # --- tray icon ---

    def _make_icon(self, *, running: bool) -> Image.Image:
        img = Image.new("RGBA", (self._ICON_SIZE, self._ICON_SIZE), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        fill = (34, 197, 94) if running else (107, 114, 128)  # green / gray
        draw.ellipse([4, 4, 60, 60], fill=fill)
        draw.ellipse([18, 18, 46, 46], fill=(255, 255, 255, 160))
        return img

    # --- menu callbacks (called from pystray / main thread) ---

    def _on_open_browser(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        scheme = "https" if USE_HTTPS else "http"
        webbrowser.open(f"{scheme}://localhost:{PORT}")

    def _on_show_logs(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self._root:
            self._root.after(0, self._show_window)

    def _on_quit(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self._root:
            self._root.after(0, self._shutdown)
        else:
            icon.stop()

    # --- Flask server ---

    def _run_server(self) -> None:
        scheme = "https" if USE_HTTPS else "http"
        log.info("ℹ️ Launcher serving on %s://%s:%s", scheme, HOST, PORT)
        try:
            self._server = make_server(HOST, PORT, app, ssl_context=SSL_CONTEXT, threaded=True)
            self._icon.icon = self._make_icon(running=True)
            self._server.serve_forever()
        except Exception:
            log.exception("❌ Launcher server failed")

    # --- log drain (runs every 200 ms on root.after) ---

    def _drain_logs(self) -> None:
        new_lines = self._log_handler.drain()
        if new_lines:
            self._log_lines.extend(new_lines)
            if self._log_text and self._window and self._window.winfo_exists():
                txt = self._log_text
                txt.configure(state=tk.NORMAL)
                for line in new_lines:
                    txt.insert(tk.END, line + "\n")
                txt.configure(state=tk.DISABLED)
                txt.see(tk.END)
        if self._root:
            self._root.after(200, self._drain_logs)

    # --- main window ---

    def _show_window(self) -> None:
        if self._window and self._window.winfo_exists():
            self._window.deiconify()
            self._window.lift()
            return
        self._build_window()

    def _build_window(self) -> None:
        win = tk.Toplevel(self._root)
        win.title("Launcher")
        win.geometry("700x440")
        win.minsize(500, 300)
        self._window = win

        scheme = "https" if USE_HTTPS else "http"
        url = f"{scheme}://localhost:{PORT}"

        # --- status row ---
        status_bar = ttk.Frame(win, padding=(10, 8, 10, 6))
        status_bar.pack(fill=tk.X)
        tk.Label(status_bar, text="●", fg="#22c55e", font=("", 15)).pack(side=tk.LEFT)
        ttk.Label(status_bar, text=f"  Running  —  {url}", font=("", 11)).pack(side=tk.LEFT)

        ttk.Separator(win, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=6)

        # --- log area ---
        log_frame = ttk.Frame(win, padding=(6, 4, 6, 6))
        log_frame.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(
            log_frame,
            state=tk.DISABLED,
            wrap=tk.NONE,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            selectbackground="#264f78",
            relief=tk.FLAT,
            padx=4,
            pady=4,
        )
        sy = ttk.Scrollbar(log_frame, command=txt.yview)
        sx = ttk.Scrollbar(log_frame, orient=tk.HORIZONTAL, command=txt.xview)
        txt.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        txt.pack(fill=tk.BOTH, expand=True)
        self._log_text = txt

        # Populate buffered lines accumulated before the window opened
        txt.configure(state=tk.NORMAL)
        for line in self._log_lines:
            txt.insert(tk.END, line + "\n")
        txt.configure(state=tk.DISABLED)
        txt.see(tk.END)

        # Hide instead of destroy on close so the tray icon stays alive
        win.protocol("WM_DELETE_WINDOW", win.withdraw)

    # --- shutdown ---

    def _shutdown(self) -> None:
        log.info("ℹ️ Shutting down launcher…")
        if self._server:
            threading.Thread(target=self._server.shutdown, daemon=True).start()
        self._icon.stop()
        if self._root:
            self._root.quit()

    # --- entry point ---

    def _setup(self, icon: pystray.Icon) -> None:
        """Called by pystray in a background thread once the icon is ready."""
        icon.visible = True
        if self._port_in_use():
            icon.notify("Launcher server is already running.", "Already Running")
            icon.stop()
            return
        self._root = tk.Tk()
        self._root.withdraw()
        self._root.after(200, self._drain_logs)
        threading.Thread(target=self._run_server, daemon=True, name="flask").start()
        self._root.mainloop()
        # mainloop exits → stop the icon (causes icon.run() to return)
        icon.stop()

    def _port_in_use(self) -> bool:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            return s.connect_ex(("127.0.0.1", PORT)) == 0

    def run(self) -> None:
        """Block until quit. pystray message pump runs in the calling thread."""
        self._icon.run(setup=self._setup)


if __name__ == "__main__":
    TrayApp().run()
