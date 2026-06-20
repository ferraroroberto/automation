import PySimpleGUI as sg # type: ignore
import pygetwindow as gw
import pyautogui as pag
import threading
import time
import sys
import keyboard

# ---------- helper functions ----------
def find_chrome_window():
    """Return the last active Chrome window or None."""
    for w in gw.getAllWindows():
        if "Google Chrome" in w.title and not w.isMinimized:
            return w
    return None

def type_text(text, delay, stop_event, status_callback):
    for i, char in enumerate(text):
        if stop_event.is_set():
            status_callback(f"Stopped at {i}/{len(text)} chars.")
            return
        if char == "\n":
            keyboard.send("shift+enter")  # inserts newline inside message
        else:
            keyboard.write(char, delay=delay)
    status_callback("Done!")

# ---------- GUI layout ----------
layout = [
    [sg.Text("Text to “type-paste”:")],
    [sg.Multiline(key="-TEXT-", size=(80, 20))],
    [
        sg.Text("Typing delay (ms):"),
        sg.Input("1", size=(6,1), key="-DELAY-"),
        sg.Button("▶ Start typing", key="-START-", button_color=("white", "green")),
        sg.Button("■ Stop", key="-STOP-", disabled=True, button_color=("white","firebrick3")),
    ],
    [sg.Text("Status: Ready", key="-STATUS-")]
]

window = sg.Window("KeyCaster", layout, font=("Segoe UI", 10))

# ---------- event loop ----------
typing_thread = None
stop_event = threading.Event()

def update_status(msg):
    window["-STATUS-"].update(f"Status: {msg}")

while True:
    event, values = window.read(timeout=100)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        stop_event.set()
        break

    if event == "-START-":
        chrome = find_chrome_window()
        if not chrome:
            update_status("Chrome window not found!")
            continue
        chrome.activate()      # bring to foreground
        time.sleep(0.2)        # let the focus settle

        text = values["-TEXT-"]
        try:
            delay = max(0, float(values["-DELAY-"])/1000)
        except ValueError:
            update_status("Delay must be a number.")
            continue

        stop_event.clear()
        typing_thread = threading.Thread(
            target=type_text,
            args=(text, delay, stop_event, update_status),
            daemon=True
        )
        typing_thread.start()
        window["-START-"].update(disabled=True)
        window["-STOP-"].update(disabled=False)
        update_status("Typing…")

    if event == "-STOP-":
        stop_event.set()
        window["-START-"].update(disabled=False)
        window["-STOP-"].update(disabled=True)
        update_status("Aborting…")
