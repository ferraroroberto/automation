import subprocess
import keyboard
import win32gui

# Define the two commands to run
SAVE_COMMAND = 'python email-automation-save.py'
ARCHIVE_COMMAND = 'python email-automation-archive.py'

# Define function to launch a command in a new window
def launch_command_prompt(command):
    subprocess.Popen(['start', 'cmd', '/k', command], shell=True)
    print(f"Launched command '{command}' in a new window.")

# Define function to check if a window is active
def is_window_active(window_title):
    return win32gui.GetWindowText(win32gui.GetForegroundWindow()).startswith(window_title)

# Define function to check if a window is in the foreground
def is_window_in_foreground(window_title):
    return win32gui.GetWindowText(win32gui.GetForegroundWindow()) == window_title

# Set up keyboard hotkeys
keyboard.add_hotkey('ctrl+shift+1', lambda: launch_command_prompt(SAVE_COMMAND))
keyboard.add_hotkey('ctrl+shift+2', lambda: launch_command_prompt(ARCHIVE_COMMAND))

# Wait for hotkeys to be pressed
keyboard.wait()