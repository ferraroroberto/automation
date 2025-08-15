# Mouse Mover

A utility that continuously moves the mouse pointer in a rectangular pattern to prevent screen locking. The application uses both mouse movement and system-level activity signals for maximum effectiveness.

## Features

- Prevents screen locking by simulating mouse activity
- Configurable movement pattern (distance in all four directions)
- Adjustable movement speed
- Optional mouse clicking at direction changes
- Simple GUI with stop/resume functionality
- System-level screen lock prevention on Windows

## Requirements

- Python 3
- PyAutoGUI library

## Installation

1. Ensure Python 3 is installed on your system
2. Install the required dependency:
   ```
   pip install pyautogui
   ```

## Usage

### Command Line Parameters

The program accepts the following command-line arguments:

- `--up`: Distance in pixels to move upward (default: 100)
- `--right`: Distance in pixels to move rightward (default: 100)
- `--down`: Distance in pixels to move downward (default: 100) 
- `--left`: Distance in pixels to move leftward (default: 100)
- `--speed`: Delay (in seconds) between each pixel movement (default: 0.005)
- `--click`: Enable mouse clicking at each direction change

### Example

```
python mouse_mover.py --up 100 --right 200 --down 100 --left 200 --speed 0.005 --click
```

## User Interface

The application provides a simple GUI with:

- Sliders to adjust movement parameters in real-time
- Checkbox to toggle clicks at direction changes
- STOP/RESUME button to control the mouse movement
- Standard window close button to exit the application

## Implementation Details

### Main Components

1. **MouseMover Class**: A threaded class that handles the mouse movement logic.
2. **System-Level Sleep Prevention**: Uses Windows API to prevent the system from sleeping.
3. **GUI Controls**: Tkinter-based interface for runtime configuration.

### Code Structure

```python
# Key imports
import argparse, threading, time, tkinter, ctypes, pyautogui

# MouseMover class
class MouseMover(threading.Thread):
    # Initialization with movement parameters
    # UI controls creation
    # Main thread loop for mouse movement

# Helper functions
def prevent_sleep()  # Prevents Windows from sleeping
def parse_args()     # Parses command line arguments
def main()           # Sets up the UI and starts the mouse mover thread
```

### Implementation Notes

- The mouse movement occurs pixel by pixel with configurable delays
- The application runs in a separate thread from the UI
- Uses PyAutoGUI's moveRel() for relative mouse movement
- Can optionally click at each direction change
- Parameters can be adjusted in real-time via the UI

## How It Works

1. The application creates a thread that moves the mouse in a rectangular pattern
2. It also uses the Windows API to prevent system sleep
3. The UI allows for adjusting all movement parameters
4. The movement can be paused/resumed with the STOP/RESUME button
5. The application can be closed via the window's close button

## Recreating From Scratch

To recreate this application from scratch:

1. Create a new Python file named `mouse_mover.py`
2. Add the shebang line and docstring with description
3. Import required libraries
4. Implement the prevent_sleep function for Windows
5. Create the MouseMover class as a thread
6. Implement UI controls creation method
7. Implement the main mouse movement loop
8. Add command-line argument parsing
9. Create the main function with UI setup
10. Add the standard Python entry point

## License

This tool is provided as-is without any warranty. Use at your own risk.
