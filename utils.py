# requirements: public
import win32gui

# Function to read the parameters from the txt file
def read_params_from_txt_file(file_path):
    params = {}
    with open(file_path, 'r') as f:
        for line in f:
            if line.strip():
                key, value = line.strip().split(" = ", 1)
                params[key.strip()] = value.strip()
    return params

# Function to get the path of the foreground windowx explorer: it onluy works if the explorer window has the full path showing
def get_active_explorer_path():
    # Get a list of all open windows
    windows = []
    win32gui.EnumWindows(lambda hwnd, windows: windows.append(hwnd), windows)

    # Find the first window that has a path to a directory in its title, taking into account network drives too
    explorer_handle = None
    for hwnd in windows:
        window_text = win32gui.GetWindowText(hwnd)
        if (":\\" in window_text or window_text.startswith("\\\\S555")) and not window_text.endswith(".exe") and not window_text.endswith(".py"):
            explorer_handle = hwnd
            break

    # If no Explorer window with a path in its title is found, print an error message and return None
    if explorer_handle is None:
        print("No Windows Explorer instance found with a path in its title.")
        return None

    # Get the window text of the selected Explorer window, which is the full path to the directory being displayed
    folder_path = win32gui.GetWindowText(explorer_handle)

    # Debugging output
    print(f"Window handle: {explorer_handle}, Folder path: {folder_path}")

    # Return the full path to the directory being displayed in the Explorer window
    return folder_path