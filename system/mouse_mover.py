#!/usr/bin/env python3
"""
mouse_mover.py

Continuously moves the mouse pointer in a rectangle and prevents the screen from locking
by using both mouse movement and system-level activity signals.

Runs from the system tray. Click the tray icon to open the window;
closing the window minimizes it back to the tray. Quit from the tray menu.

UI:
  * STOP/RESUME the motion
  * QUIT via the tray menu

Parameters:
    --up     Distance in pixels to move upward.
    --right  Distance in pixels to move rightward.
    --down   Distance in pixels to move downward.
    --left   Distance in pixels to move leftward.
    --speed  Delay (in seconds) between each pixel movement.
    --click  Left click at each direction change.

Requirements:
    pip install pyautogui pystray Pillow

Usage:
    pythonw mouse_mover.py --up 100 --right 200 --down 100 --left 200 --speed 0.005
"""

import argparse
import logging
import socket
import threading
import time
import tkinter as tk
import ctypes
from typing import Optional

import pyautogui
import pystray
from PIL import Image, ImageDraw

# ------------------------------------------------------------------------------
# Configuration
# ------------------------------------------------------------------------------
pyautogui.FAILSAFE = False

# Constants for system-level screen-lock prevention
ES_CONTINUOUS       = 0x80000000
ES_SYSTEM_REQUIRED  = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


def prevent_sleep() -> None:
    """Tell Windows not to turn off screen or go to sleep."""
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )


class MouseMover(threading.Thread):
    def __init__(self, click_on_turn: bool = False):
        super().__init__()
        self.daemon = True  # Thread will exit when main program does
        self.running = True
        self.paused = False
        # Instance variables instead of dict
        self.up = 5
        self.right = 5
        self.down = 5
        self.left = 5
        self.speed = 0.005
        self.click_on_turn = click_on_turn

    def create_controls(self, parent):
        """Create parameter control panel"""
        control_frame = tk.LabelFrame(parent, text="Movement Parameters")
        control_frame.pack(padx=5, pady=5, fill="x")

        params = {
            'up': (self.up, 1, 500),
            'right': (self.right, 1, 500),
            'down': (self.down, 1, 500),
            'left': (self.left, 1, 500),
            'speed': (self.speed, 0.005, 1)
        }

        for row, (param, (default, min_val, max_val)) in enumerate(params.items()):
            tk.Label(control_frame, text=f"{param.capitalize()}:").grid(row=row, column=0, padx=5, pady=2)

            slider = tk.Scale(control_frame,
                            from_=min_val,
                            to=max_val,
                            resolution=0.001 if param == 'speed' else 1,
                            orient="horizontal",
                            length=200)
            slider.set(default)
            slider.grid(row=row, column=1, padx=5, pady=2)

            def update_value(value, param_name=param):
                setattr(self, param_name, float(value))

            slider.config(command=lambda v, p=param: update_value(v, p))

        # Add checkbox for click_on_turn
        self.click_var = tk.BooleanVar(value=self.click_on_turn)
        def on_click_toggle():
            self.click_on_turn = self.click_var.get()
        click_checkbox = tk.Checkbutton(
            control_frame, text="Left click at direction change",
            variable=self.click_var, command=on_click_toggle
        )
        click_checkbox.grid(row=len(params), column=0, columnspan=2, sticky="w", padx=5, pady=5)

    def run(self):
        """Main loop that moves the mouse"""
        while self.running:
            if not self.paused:
                try:
                    prevent_sleep()
                    # Move UP pixel by pixel
                    for _ in range(int(self.up)):
                        if not self.running or self.paused: break
                        pyautogui.moveRel(0, -1, duration=self.speed)
                    if self.click_on_turn and not self.paused and self.running:
                        pyautogui.click()

                    # Move RIGHT pixel by pixel
                    for _ in range(int(self.right)):
                        if not self.running or self.paused: break
                        pyautogui.moveRel(1, 0, duration=self.speed)
                    if self.click_on_turn and not self.paused and self.running:
                        pyautogui.click()

                    # Move DOWN pixel by pixel
                    for _ in range(int(self.down)):
                        if not self.running or self.paused: break
                        pyautogui.moveRel(0, 1, duration=self.speed)
                    if self.click_on_turn and not self.paused and self.running:
                        pyautogui.click()

                    # Move LEFT pixel by pixel
                    for _ in range(int(self.left)):
                        if not self.running or self.paused: break
                        pyautogui.moveRel(-1, 0, duration=self.speed)
                    if self.click_on_turn and not self.paused and self.running:
                        pyautogui.click()
                except Exception:
                    pass
            time.sleep(0.1)


class MouseMoverApp:
    """Tray-resident wrapper around MouseMover.

    Mirrors foldersearcher's pattern: the pystray icon is created up front;
    the Tk root is built inside the pystray setup callback so tkinter shares
    the icon's message-loop thread. The control window is a Toplevel built
    lazily on first Open and reused thereafter (close = withdraw).
    """

    _ICON_SIZE = 64
    _LOCK_PORT = 50218  # single-instance lock via 127.0.0.1 socket

    def __init__(self, args: argparse.Namespace):
        self.args = args

        self.mover = MouseMover(click_on_turn=args.click)
        self.mover.up = args.up
        self.mover.right = args.right
        self.mover.down = args.down
        self.mover.left = args.left
        self.mover.speed = args.speed

        self.root: Optional[tk.Tk] = None
        self.window: Optional[tk.Toplevel] = None
        self._widgets_built = False
        self.btn_stop: Optional[tk.Button] = None

        self._lock_socket: Optional[socket.socket] = None

        self._icon = pystray.Icon(
            "mouse_mover",
            self._make_icon(),
            "Mouse Mover",
            pystray.Menu(
                pystray.MenuItem("Open", self._on_open, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._on_quit),
            ),
        )

        logger.info("Mouse Mover application initialized")

    # ------------------------------------------------------------------
    # Tray icon
    # ------------------------------------------------------------------

    def _make_icon(self) -> Image.Image:
        """Draw a simple computer-mouse tray icon."""
        size = self._ICON_SIZE
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        body = (55, 65, 81)        # dark grey
        highlight = (96, 165, 250) # blue scroll wheel
        # mouse body — rounded ellipse
        draw.ellipse([14, 8, 50, 58], fill=body, outline=(17, 24, 39), width=2)
        # split between left/right buttons
        draw.line([(32, 8), (32, 28)], fill=(17, 24, 39), width=2)
        # horizontal split below buttons
        draw.line([(15, 28), (49, 28)], fill=(17, 24, 39), width=2)
        # scroll wheel
        draw.rectangle([29, 14, 35, 24], fill=highlight, outline=(17, 24, 39), width=1)
        # cable
        draw.line([(32, 8), (32, 2)], fill=body, width=3)
        return img

    def _on_open(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.root:
            self.root.after(0, self._show_window)

    def _on_quit(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.root:
            self.root.after(0, self._shutdown)
        else:
            icon.stop()

    def _shutdown(self) -> None:
        logger.info("Shutting down Mouse Mover")
        self.mover.running = False
        self.mover.paused = True
        try:
            self._icon.stop()
        except Exception:
            pass
        if self.root:
            self.root.quit()

    # ------------------------------------------------------------------
    # Window lifecycle
    # ------------------------------------------------------------------

    def _show_window(self) -> None:
        if not self._widgets_built:
            self._build_window()
            return
        if self.window and self.window.winfo_exists():
            self.window.deiconify()
            self.window.attributes("-topmost", True)
            self.window.lift()
            self.window.focus_force()

    def _build_window(self) -> None:
        assert self.root is not None
        win = tk.Toplevel(self.root)
        win.title("Mouse Mover")
        win.resizable(True, True)
        win.attributes("-topmost", True)
        self.window = win

        self.mover.create_controls(win)

        def toggle():
            self.mover.paused = not self.mover.paused
            if self.btn_stop is not None:
                self.btn_stop.config(text="RESUME" if self.mover.paused else "STOP")

        self.btn_stop = tk.Button(win, text="STOP", width=12, command=toggle)
        self.btn_stop.pack(side=tk.LEFT, padx=10, pady=20)

        win.protocol("WM_DELETE_WINDOW", win.withdraw)
        self._widgets_built = True
        win.lift()
        win.focus_force()

    # ------------------------------------------------------------------
    # Single-instance lock and pystray setup
    # ------------------------------------------------------------------

    def _acquire_lock(self) -> bool:
        """Bind to a fixed local port as a mutex. False if another instance holds it."""
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("127.0.0.1", self._LOCK_PORT))
            s.listen(1)
            self._lock_socket = s
            return True
        except OSError:
            s.close()
            return False

    def _setup_pystray(self, icon: pystray.Icon) -> None:
        """Run on pystray's worker thread once the icon is visible."""
        icon.visible = True

        if not self._acquire_lock():
            icon.notify("Mouse Mover is already running.", "Already Running")
            icon.stop()
            return

        self.root = tk.Tk()
        self.root.withdraw()

        # Start the mover thread now that the Tk root exists
        self.mover.start()

        self.root.mainloop()
        try:
            icon.stop()
        except Exception:
            pass

    def run(self) -> None:
        """Start the tray application. Blocks until Quit."""
        logger.info("Starting Mouse Mover (tray)")
        self._icon.run(setup=self._setup_pystray)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Move mouse in a rectangle and prevent screen lock.")
    p.add_argument("--up",    type=int, default=100, help="Pixels to move up")
    p.add_argument("--right", type=int, default=100, help="Pixels to move right")
    p.add_argument("--down",  type=int, default=100, help="Pixels to move down")
    p.add_argument("--left",  type=int, default=100, help="Pixels to move left")
    p.add_argument("--speed", type=float, default=0.005, help="Delay (sec) between pixel moves")
    p.add_argument("--click", action="store_true", help="Left click at each direction change")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    app = MouseMoverApp(args)
    app.run()


if __name__ == "__main__":
    main()
