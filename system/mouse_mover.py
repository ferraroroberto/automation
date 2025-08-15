#!/usr/bin/env python3
"""
mouse_mover.py

Continuously moves the mouse pointer in a rectangle and prevents the screen from locking
by using both mouse movement and system-level activity signals.

UI:
  • STOP/RESUME the motion
  • QUIT the application

Parameters:
    --up     Distance in pixels to move upward.
    --right  Distance in pixels to move rightward.
    --down   Distance in pixels to move downward.
    --left   Distance in pixels to move leftward.
    --speed  Delay (in seconds) between each pixel movement.

Requirements:
    pip install pyautogui

Usage:
    python mouse_mover.py --up 100 --right 200 --down 100 --left 200 --speed 0.005
"""

import argparse
import threading
import time
import sys
import tkinter as tk
import ctypes
import pyautogui

# ------------------------------------------------------------------------------
# Configuration
# ------------------------------------------------------------------------------
pyautogui.FAILSAFE = False

# Constants for system-level screen-lock prevention
ES_CONTINUOUS       = 0x80000000
ES_SYSTEM_REQUIRED  = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

def prevent_sleep():
    """Tell Windows not to turn off screen or go to sleep."""
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

class MouseMover(threading.Thread):
    def __init__(self, click_on_turn=False):
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
                except:
                    pass
            time.sleep(0.1)

def parse_args():
    p = argparse.ArgumentParser(description="Move mouse in a rectangle and prevent screen lock.")
    p.add_argument("--up",    type=int, default=100, help="Pixels to move up")
    p.add_argument("--right", type=int, default=100, help="Pixels to move right")
    p.add_argument("--down",  type=int, default=100, help="Pixels to move down")
    p.add_argument("--left",  type=int, default=100, help="Pixels to move left")
    p.add_argument("--speed", type=float, default=0.005, help="Delay (sec) between pixel moves")
    p.add_argument("--click", action="store_true", help="Left click at each direction change")
    return p.parse_args()

def main():
    args = parse_args()
    root = tk.Tk()
    root.title("Mouse Mover")
    root.resizable(True, True)

    mover = MouseMover(click_on_turn=args.click)
    # Set initial values from args
    mover.up = args.up
    mover.right = args.right
    mover.down = args.down
    mover.left = args.left
    mover.speed = args.speed

    mover.create_controls(root)
    mover.start()

    def toggle():
        mover.paused = not mover.paused
        btn_stop.config(text="RESUME" if mover.paused else "STOP")

    def on_close():
        mover.running = False
        mover.paused = True
        root.after(100, root.destroy)

    btn_stop = tk.Button(root, text="STOP", width=12, command=toggle)
    btn_stop.pack(side=tk.LEFT, padx=10, pady=20)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()
