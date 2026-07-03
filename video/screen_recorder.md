# Screen Recorder

A Tkinter GUI for recording the screen to MP4, with adjustable FPS, multi-monitor selection, and an always-on mouse cursor overlay. Frames are captured with `mss` and written directly via OpenCV's `VideoWriter` (`mp4v` codec) — there is no FFmpeg step and no audio track.

## Features

- **Multi-monitor support** – A dropdown lists every monitor detected by `screeninfo.get_monitors()` (`index: WIDTHxHEIGHT @ (x,y)`); recording captures only the selected monitor's region
- **Adjustable FPS** – Enter any positive integer frames-per-second in the GUI (default 5); lower FPS gives smaller files but choppier video
- **Mouse cursor overlay** – Every frame draws the cursor as a white filled circle (radius 6) with a black outline (radius 10, thickness 3) at the pointer's position via `pyautogui.position()`, so the cursor stays visible on light and dark backgrounds; it's only drawn while the pointer is inside the selected monitor's bounds
- **Pause / Resume** – Pausing stops writing new frames without closing the output file or losing the recording session
- **Segment splitting (present but disabled by default)** – The code can start a new numbered output file once the current one exceeds `MAX_FILE_MB`, but that constant is `None` in the shipped script, so splitting never triggers unless the constant is edited

## Requirements

- Python 3.x (tkinter for the GUI, included with standard Python on Windows)
- `opencv-python`, `mss`, `pyautogui`, `screeninfo`, `numpy` (`pip install opencv-python mss pyautogui screeninfo numpy`)
- No FFmpeg dependency — encoding is done entirely through OpenCV's `VideoWriter`

## Running the Application

**GUI (only mode — no CLI arguments):**
```bash
python video/screen_recorder.py
```

**Double-click:** `screen_recorder.bat`

The window opens at 520x140 (resizable) titled "Screen Recorder". Note: as shipped, `screen_recorder.bat` launches `pythonw.exe` from a legacy environment path (`c:\Mis Datos en Local\temporal\python\projects\work\automation\...`) that does not exist on this machine, unlike the sibling `.bat` launchers in this folder which resolve their venv from `E:\automation\automation\.env`'s `VENV_FOLDER`. Until the `.bat` is updated to match that pattern, run the script directly with `python video/screen_recorder.py` from the automation venv.

## GUI Usage

1. **Screen** – Choose the monitor to record from the dropdown.
2. **Frames per second** – Enter the capture rate (default 5).
3. **Start** – Begins recording; the monitor dropdown and FPS field lock while recording is active, and the button becomes **Stop**.
4. **Pause / Resume** – Toggles frame capture on and off without ending the recording; the status line shows "Paused" or "Recording..." with the current output path.
5. **Stop** – Finalizes the video file, unlocks the controls, and shows a "Recording finished" dialog with the saved path.
6. **Quit** – Stops any in-progress recording, then closes the window.

The status label at the bottom always reflects the current state: `Idle`, `Recording... → <path>`, `Paused → <path>`, or `Saved: <path>`.

## Output Files

Recordings are saved to `~/Downloads/videos` (created automatically if missing), named:

```
screen_record_{YYYYMMDD_HHMMSS}.mp4
```

If segment splitting were enabled (it isn't by default — see Features), later segments in the same recording would get a zero-padded suffix, e.g. `screen_record_20260703_143000_02.mp4`.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No monitors detected" | `screeninfo` found no monitors; the app closes immediately — check display drivers/configuration |
| "Please enter a positive integer for FPS" | The FPS field must be a whole number greater than 0 |
| "Failed to initialize video writer" | OpenCV's `VideoWriter` couldn't open the output file/codec; check that `opencv-python` is installed correctly and the output folder is writable |
| "Recording Error" dialog mid-recording | An exception occurred in the capture loop (e.g. a monitor was disconnected); recording stops automatically and the partial file remains at the last saved path |
| Double-clicking `screen_recorder.bat` does nothing | The `.bat` points at a legacy Python path that no longer exists on this machine; run `python video/screen_recorder.py` directly, or update the `.bat` to read `VENV_FOLDER` from `.env` like `video_trim.bat` does |
