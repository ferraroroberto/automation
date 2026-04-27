# AI source > https://chatgpt.com/c/67963e0c-6fd8-8009-88b4-df158e8732c2

import logging
import os
import sys
from datetime import datetime
from mss import mss
from PIL import Image

log = logging.getLogger(__name__)

def list_monitors():
    """List all available monitors and their dimensions."""
    with mss() as sct:
        monitors = sct.monitors  # List of all monitors
    return monitors

def capture_specific_monitor(monitor_index, output_folder="D:\\MwSnap_temporal"):
    """Capture a screenshot of a specific monitor by its index."""
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Generate a filename with the current timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d, %H_%M_%S")
    file_name = f"capture-{timestamp}.jpg"
    file_path = os.path.join(output_folder, file_name)

    # Capture the screenshot for the selected monitor
    with mss() as sct:
        monitors = sct.monitors
        if monitor_index < 1 or monitor_index > len(monitors):
            raise ValueError("Invalid monitor index. Please choose a valid one.")

        monitor = monitors[monitor_index]
        screenshot = sct.grab(monitor)

        # Save the screenshot as a high-quality JPEG
        img = Image.frombytes("RGB", screenshot.size, screenshot.rgb)
        img.save(file_path, "JPEG", quality=95)

    log.info("Screenshot saved to %s", file_path)

def get_bottom_left_monitor(monitors):
    """Find the monitor at the bottom-left position."""
    return min(monitors[1:], key=lambda m: (m["top"], m["left"]))

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    output_folder = "E:\\downloads\\snaps"
    monitors = list_monitors()
    monitor_index = None

    # Check if parameters are provided
    if len(sys.argv) > 1:
        param = sys.argv[1].lower()
        if param.isdigit():
            monitor_index = int(param)
        elif param == "bottom_left":
            bottom_left_monitor = get_bottom_left_monitor(monitors)
            monitor_index = monitors.index(bottom_left_monitor)
        elif param == "primary":
            monitor_index = 3  # The primary monitor is always at index 1
        else:
            log.error("Invalid parameter. Use a monitor number, 'primary', or 'bottom_left'.")
            sys.exit(1)
    else:
        log.info("Available monitors:")
        for idx, monitor in enumerate(monitors[1:], start=1):
            log.info("%d: %s", idx, monitor)
        try:
            monitor_index = int(input("Enter the monitor number to capture (1, 2, ...): "))
        except ValueError:
            log.error("Invalid input. Please enter a number.")
            sys.exit(1)

    # Validate and capture the monitor
    try:
        capture_specific_monitor(monitor_index, output_folder)
    except ValueError as e:
        log.error("%s", e)
        log.info("Available monitors:")
        for idx, monitor in enumerate(monitors[1:], start=1):
            log.info("%d: %s", idx, monitor)
        monitor_index = int(input("Enter the monitor number to capture (1, 2, ...): "))
        capture_specific_monitor(monitor_index, output_folder)
