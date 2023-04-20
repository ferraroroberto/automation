import subprocess
import keyboard
import win32gui
import time

# define the two commands to run
cmd1 = 'python email-automation-save.py'
cmd2 = 'python email-automation-archive.py'

# define function to launch a command in a new window
def launch_in_new_window(command):
    subprocess.Popen(['start', 'cmd', '/k', command], shell=True)
    print(f"Launched command '{command}' in a new window.")

# define function to check if a window is active
def is_window_active(window_title):
    return win32gui.GetWindowText(win32gui.GetForegroundWindow()).startswith(window_title)

# define function to check if a window is in the foreground
def is_window_in_foreground(window_title):
    return win32gui.GetWindowText(win32gui.GetForegroundWindow()) == window_title

# set up keyboard hotkeys
keyboard.add_hotkey('ctrl+shift+1', lambda: launch_in_new_window(cmd1))
keyboard.add_hotkey('ctrl+shift+2', lambda: launch_in_new_window(cmd2))

# wait for hotkeys to be pressed
keyboard.wait()