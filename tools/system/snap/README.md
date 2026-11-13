# Snap

Press a global hotkey anywhere, drag a rectangle over the frozen screen, release: the selection is saved as a timestamped JPEG and copied to the clipboard. No editor, no dialog. Esc or a right-click cancels and saves nothing.

## Run

- `snap.bat` starts it in the system tray with no console window (it copies `config.sample.json` to `config.json` on first run). The tray menu has **Open folder** and **Quit**.
- `.\.venv\Scripts\python.exe tools\system\snap\snap_tool.py --install-startup` adds a Startup shortcut so it starts at login; `--remove-startup` removes it. Nothing is registered unless you run that command.
- A second launch while one is running exits at once.

## Config (`config.json`, gitignored; `config.sample.json` is the template)

| Key | Default | Meaning |
|---|---|---|
| `folder` | `~/Pictures/snaps` (the sample points at the folder the `/screen` skill reads) | Where captures go; `~` and `%VARS%` expand |
| `hotkey` | `ctrl+shift+a` | One or more of `ctrl`, `shift`, `alt`, `win`, then a letter, digit, `f1`-`f24`, `printscreen` or `space` |
| `jpeg_quality` | `90` | 1-100 |
| `copy_to_clipboard` | `true` | Also put the capture on the clipboard |

An invalid file stops the tool with a dialog and a log line naming the setting. A hotkey another app already owns does the same, so it is never a silent no-op; pick another combination in `config.json`.

## Files

Captures are named `snap_YYYY-MM-DD_HH-MM-SS-mmm.jpg`; a name already taken gets `_2`, `_3`, ... so nothing is overwritten.

The log is `snap.log` under `runtime_data_dir()/snap/` (`%LOCALAPPDATA%\automation\snap\`). Each capture logs one line for the overlay (grab time and hotkey-to-overlay ms) and one for the save (path, size, release-to-file ms).

## How it stays fast

The process is resident and does its start-up work once. `RegisterHotKey` runs on its own message-pump thread and wakes the Tk thread; the overlay is a borderless top-most window built in advance, so a keypress only grabs the virtual screen with `mss` and paints it. The process is per-monitor DPI aware, so screen, capture and crop all use physical pixels on scaled displays, and the capture spans every monitor (a monitor left of or above the primary has a negative origin, handled in `snap_core.py`).

## Tests

`tests/system/snap/` covers naming, crop maths with monitor offsets, hotkey parsing and config. The overlay and hotkey path needs a real desktop and is checked by hand.
