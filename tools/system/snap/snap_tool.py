"""Hotkey region screenshot: press the hotkey, drag a rectangle, get a timestamped JPEG.

The process stays resident so the hotkey path does no imports or start-up work: a
Win32 ``RegisterHotKey`` thread wakes the Tk thread, which grabs the whole virtual
screen with ``mss``, paints it onto a pre-built borderless overlay and waits for the
drag. Release saves the crop into the configured folder (and the clipboard); Esc or
a right-click cancels.

Run ``snap.bat`` (no console window), or ``--install-startup`` once to add a Startup
shortcut. See ``README.md`` for the config keys.
"""

from __future__ import annotations

import argparse
import ctypes
import io
import logging
import os
import sys
import threading
import time
import tkinter as tk
from ctypes import wintypes
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

import mss
from PIL import Image, ImageDraw, ImageTk

TOOL_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOL_DIR.parent))

from _lib.runtime_files import file_handler, make_dirs, runtime_data_dir, save_bytes  # noqa: E402
from _lib.single_instance import acquire_named_mutex  # noqa: E402
from snap_core import (  # noqa: E402
    MOD_NOREPEAT,
    ConfigError,
    SnapConfig,
    load_config,
    parse_hotkey,
    screen_to_image,
    selection_rect,
    unique_path,
    window_geometry,
)

logger = logging.getLogger("snap")

CONFIG_PATH = TOOL_DIR / "config.json"
MUTEX_NAME = "automation.snap.single_instance"
STARTUP_SHORTCUT_NAME = "Snap region screenshot.lnk"

HOTKEY_ID = 1
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
CF_DIB = 8
ERROR_HOTKEY_ALREADY_REGISTERED = 1409
HOTKEY_REGISTER_TIMEOUT_S = 3.0
CLIPBOARD_ATTEMPTS = 5
CLIPBOARD_RETRY_S = 0.02

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
kernel32.GetCurrentThreadId.restype = wintypes.DWORD
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]


def make_dpi_aware() -> None:
    """Report physical pixels to Tk and mss; without this a scaled display crops the wrong region."""
    try:
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))  # per-monitor v2
    except (AttributeError, OSError):
        ctypes.windll.shcore.SetProcessDpiAwareness(2)


def show_error_box(message: str) -> None:
    """The tool has no console under pythonw, so a fatal start-up error also gets a visible dialog."""
    user32.MessageBoxW(None, message, "Snap", 0x10)


def copy_image_to_clipboard(image: Image.Image) -> bool:
    """Put ``image`` on the clipboard as a DIB; the clipboard can be held by another app, so retry briefly."""
    import win32clipboard

    buffer = io.BytesIO()
    image.save(buffer, "BMP")
    dib = buffer.getvalue()[14:]  # drop the 14-byte BMP file header
    for _ in range(CLIPBOARD_ATTEMPTS):
        try:
            win32clipboard.OpenClipboard()
        except Exception:  # pywintypes.error: another process owns the clipboard right now
            time.sleep(CLIPBOARD_RETRY_S)
            continue
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(CF_DIB, dib)
            return True
        finally:
            win32clipboard.CloseClipboard()
    return False


class SnapApp:
    """Owns the Tk overlay, the hotkey thread and the tray icon."""

    def __init__(self, config: SnapConfig) -> None:
        self.config = config
        self.root = tk.Tk()
        self.root.withdraw()
        self.overlay = tk.Toplevel(self.root)
        self.overlay.withdraw()
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True)
        self.canvas = tk.Canvas(self.overlay, highlightthickness=0, bd=0, cursor="crosshair", bg="black")
        self.canvas.pack(fill="both", expand=True)
        self._image_item = self.canvas.create_image(0, 0, anchor="nw")
        self._photo: Optional[ImageTk.PhotoImage] = None
        self._frozen: Optional[Image.Image] = None
        self._origin = (0, 0)
        self._drag_start: Optional[Tuple[int, int]] = None
        self._rect_items: Tuple[int, int] = (
            self.canvas.create_rectangle(0, 0, 0, 0, outline="white", width=1, state="hidden"),
            self.canvas.create_rectangle(0, 0, 0, 0, outline="black", dash=(4, 4), width=1, state="hidden"),
        )
        self._hotkey_t0 = 0.0
        self._hotkey_thread_id = 0
        self._hotkey_ready = threading.Event()
        self._hotkey_error: Optional[str] = None
        self._tray = None
        self._busy = False

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Button-3>", lambda _event: self._cancel("right-click"))
        self.overlay.bind("<Escape>", lambda _event: self._cancel("Esc"))
        self.root.bind("<<SnapHotkey>>", self._on_hotkey)
        self.root.bind("<<SnapQuit>>", lambda _event: self.quit())

    # --- hotkey thread -------------------------------------------------

    def start_hotkey(self) -> bool:
        """Register the hotkey on its own message-pump thread; False (already logged) when it cannot be taken."""
        threading.Thread(target=self._hotkey_loop, name="snap-hotkey", daemon=True).start()
        if not self._hotkey_ready.wait(HOTKEY_REGISTER_TIMEOUT_S):
            self._hotkey_error = "hotkey registration did not answer in time"
        if self._hotkey_error:
            logger.error("❌ %s", self._hotkey_error)
            return False
        return True

    def _hotkey_loop(self) -> None:
        modifiers, key = parse_hotkey(self.config.hotkey)
        self._hotkey_thread_id = kernel32.GetCurrentThreadId()
        if not user32.RegisterHotKey(None, HOTKEY_ID, modifiers | MOD_NOREPEAT, key):
            code = ctypes.get_last_error() or kernel32.GetLastError()
            reason = (
                "is already registered by another application"
                if code == ERROR_HOTKEY_ALREADY_REGISTERED
                else f"could not be registered (Win32 error {code})"
            )
            self._hotkey_error = f"hotkey {self.config.hotkey} {reason}; pick another in config.json"
            self._hotkey_ready.set()
            return
        self._hotkey_ready.set()
        message = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
            if message.message == WM_HOTKEY:
                self._hotkey_t0 = time.perf_counter()
                self.root.event_generate("<<SnapHotkey>>", when="tail")
        user32.UnregisterHotKey(None, HOTKEY_ID)

    # --- capture and overlay -------------------------------------------

    def _on_hotkey(self, _event: tk.Event) -> None:
        if self._busy:
            return
        self._busy = True
        t0 = self._hotkey_t0 or time.perf_counter()
        try:
            with mss.MSS() as grabber:
                virtual = grabber.monitors[0]
                shot = grabber.grab(virtual)
            t_grab = time.perf_counter()
            frozen = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            self._frozen = frozen
            self._origin = (virtual["left"], virtual["top"])
            self._photo = ImageTk.PhotoImage(frozen)
            self.canvas.itemconfigure(self._image_item, image=self._photo)
            for item in self._rect_items:
                self.canvas.itemconfigure(item, state="hidden")
            self.overlay.geometry(window_geometry(virtual["left"], virtual["top"], shot.width, shot.height))
            self.overlay.deiconify()
            self.overlay.update()
            self.overlay.focus_force()
            t_overlay = time.perf_counter()
        except Exception:
            self._busy = False
            logger.exception("❌ capture failed")
            return
        logger.info(
            "ℹ️ overlay up: %dx%d at %+d%+d, grab %.0f ms, hotkey-to-overlay %.0f ms",
            shot.width, shot.height, virtual["left"], virtual["top"],
            (t_grab - t0) * 1000, (t_overlay - t0) * 1000,
        )
        if (self.overlay.winfo_rootx(), self.overlay.winfo_rooty()) != self._origin:
            logger.warning(
                "⚠️ overlay sits at %d,%d instead of %s; the crop follows screen coordinates, so it stays correct",
                self.overlay.winfo_rootx(), self.overlay.winfo_rooty(), self._origin,
            )

    def _image_point(self, event: tk.Event) -> Tuple[int, int]:
        return screen_to_image(event.x_root, event.y_root, self._origin)

    def _on_press(self, event: tk.Event) -> None:
        self._drag_start = self._image_point(event)

    def _on_motion(self, event: tk.Event) -> None:
        if self._drag_start is None:
            return
        x0, y0 = self._drag_start
        x1, y1 = self._image_point(event)
        for item in self._rect_items:
            self.canvas.coords(item, x0, y0, x1, y1)
            self.canvas.itemconfigure(item, state="normal")

    def _on_release(self, event: tk.Event) -> None:
        if self._drag_start is None or self._frozen is None:
            return
        rect = selection_rect(self._drag_start, self._image_point(event), self._frozen.size)
        if rect is None:
            self._cancel("selection too small")
            return
        t_release = time.perf_counter()
        self._close_overlay()
        crop = self._frozen.crop(rect)
        try:
            data = io.BytesIO()
            crop.save(data, "JPEG", quality=self.config.jpeg_quality)
            make_dirs(self.config.folder)
            path = unique_path(self.config.folder, datetime.now())
            save_bytes(path, data.getvalue())
        except Exception:
            logger.exception("❌ saving the capture failed")
            self._finish()
            return
        copied = copy_image_to_clipboard(crop) if self.config.copy_to_clipboard else None
        if copied is False:
            logger.warning("⚠️ clipboard stayed locked by another app; the file was saved anyway")
        logger.info(
            "✅ saved %s (%dx%d, %d bytes), clipboard=%s, release-to-file %.0f ms, hotkey-to-file %.0f ms",
            path, crop.width, crop.height, path.stat().st_size, copied,
            (time.perf_counter() - t_release) * 1000, (time.perf_counter() - self._hotkey_t0) * 1000,
        )
        self._finish()

    def _cancel(self, reason: str) -> None:
        if not self._busy:
            return
        logger.info("ℹ️ cancelled (%s); nothing saved", reason)
        self._close_overlay()
        self._finish()

    def _close_overlay(self) -> None:
        self.overlay.withdraw()
        self.overlay.update_idletasks()

    def _finish(self) -> None:
        self._drag_start = None
        self._frozen = None
        self._busy = False

    # --- tray and shutdown ---------------------------------------------

    def start_tray(self) -> None:
        import pystray

        icon_image = Image.new("RGB", (64, 64), "#1f2937")
        ImageDraw.Draw(icon_image).rectangle((14, 14, 50, 50), outline="white", width=3)
        menu = pystray.Menu(
            pystray.MenuItem("Open folder", lambda: self._open_folder()),
            pystray.MenuItem("Quit", lambda: self.root.event_generate("<<SnapQuit>>", when="tail")),
        )
        self._tray = pystray.Icon("snap", icon_image, f"Snap ({self.config.hotkey})", menu)
        self._tray.run_detached()

    def _open_folder(self) -> None:
        make_dirs(self.config.folder)
        os.startfile(self.config.folder)  # noqa: S606 -- opens the user's own capture folder in Explorer

    def quit(self) -> None:
        logger.info("ℹ️ quitting")
        if self._tray is not None:
            self._tray.stop()
        if self._hotkey_thread_id:
            user32.PostThreadMessageW(self._hotkey_thread_id, WM_QUIT, 0, 0)
        self.root.quit()

    def run(self) -> None:
        self.root.mainloop()


def startup_shortcut_path() -> Path:
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup" / STARTUP_SHORTCUT_NAME


def install_startup() -> None:
    """Create the Startup shortcut that launches this tool without a console; only runs on explicit request."""
    import win32com.client

    interpreter = Path(sys.executable)
    windowless = interpreter.with_name("pythonw.exe")
    shortcut = win32com.client.Dispatch("WScript.Shell").CreateShortCut(str(startup_shortcut_path()))
    shortcut.TargetPath = str(windowless if windowless.exists() else interpreter)
    shortcut.Arguments = f'"{Path(__file__).resolve()}"'
    shortcut.WorkingDirectory = str(TOOL_DIR)
    shortcut.Save()
    print(f"Startup shortcut created: {startup_shortcut_path()}")


def remove_startup() -> None:
    target = startup_shortcut_path()
    if target.exists():
        target.unlink()
        print(f"Startup shortcut removed: {target}")
    else:
        print("No startup shortcut to remove.")


def configure_logging() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s", handlers=[])
    handlers = [file_handler(runtime_data_dir() / "snap" / "snap.log")]
    if sys.stderr is not None:  # pythonw has no console streams
        handlers.append(logging.StreamHandler())
    formatter = logging.Formatter("%(asctime)s %(message)s")
    for handler in handlers:
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    logger.setLevel(logging.INFO)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--config", type=Path, default=CONFIG_PATH, help="config file (default: config.json beside this script)")
    parser.add_argument("--install-startup", action="store_true", help="add a Startup shortcut, then exit")
    parser.add_argument("--remove-startup", action="store_true", help="remove the Startup shortcut, then exit")
    args = parser.parse_args()
    if args.install_startup:
        install_startup()
        return 0
    if args.remove_startup:
        remove_startup()
        return 0

    configure_logging()
    try:
        config = load_config(args.config)
    except ConfigError as exc:
        logger.error("❌ config: %s", exc)
        show_error_box(f"Snap config error:\n{exc}")
        return 1

    mutex, acquired = acquire_named_mutex(MUTEX_NAME)  # noqa: F841 -- the local keeps the handle alive for the process lifetime
    if not acquired:
        logger.info("ℹ️ already running; this instance exits")
        return 0

    make_dpi_aware()
    app = SnapApp(config)
    if not app.start_hotkey():
        show_error_box(app._hotkey_error or "hotkey registration failed")
        return 1
    app.start_tray()
    logger.info("✅ ready: %s -> %s (quality %d, clipboard %s)", config.hotkey, config.folder, config.jpeg_quality, config.copy_to_clipboard)
    app.run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
