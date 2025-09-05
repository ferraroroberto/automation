# Build a minimal always-on-top command launcher (“QuickDeck”) in Python

**Role:** You are an expert Python developer.
**Goal:** Implement a tiny, single-file desktop app (no third-party deps) that behaves like a mini Stream Deck: a compact window that stays on top and shows a grid of configurable buttons loaded from a JSON file. Each button launches an action (shell command, open URL, copy text, run a small Python snippet, etc.).

## Hard constraints

* **Single file:** `quickdeck.py`.
* **No external libraries.** Use only Python’s standard library (e.g., `tkinter`, `subprocess`, `json`, `webbrowser`, `platform`, `shutil`, `os`, `sys`, `pathlib`, `traceback`, `tkinter.messagebox`, `tkinter.filedialog`, `tkinter.ttk`).
* **Cross-platform:** Windows, macOS, Linux.
* **No admin rights required** and no device drivers.
* **Config via JSON file** next to the script: `quickdeck.json`. If missing, auto-create a default config.
* **Always on top** window with an option to toggle borderless mode and drag the window when borderless.
* **Graceful error handling** and minimal logging to a rotating log file `quickdeck.log` in the same directory.

## Features

1. **Window**

   * Always on top (configurable).
   * Optional **borderless** mode; if borderless, allow dragging by click-and-drag on background/title area.
   * Resizable; grid adapts.
   * Compact, clean UI (Tkinter), dark theme by default.
2. **Buttons grid**

   * Buttons loaded from JSON; support label + optional emoji.
   * Tooltips (label or a separate `hint` field).
   * Keyboard: `Ctrl+R` to reload config; `Esc` to close; optional global hotkey is **not** required.
3. **Actions (by `type`)**

   * `shell`: run a command (support `cmd` string + optional `args` list). Expand `~` and env vars like `%USERPROFILE%` or `$HOME`.
   * `url`: open with `webbrowser`.
   * `copy`: copy given text to clipboard.
   * `python`: execute a short Python snippet **safely in-process**; capture `stdout/stderr` to a pop-up or status bar.
   * `reload`: reload config from disk.
4. **JSON schema**

   * Top-level keys:

     * `title` (str), `geometry` (e.g., `"420x300+20+20"`), `topmost` (bool), `borderless` (bool)
     * `theme`: `{ "bg": "#111827", "fg": "#E5E7EB", "button_bg": "#1F2937" }`
     * `grid`: `{ "cols": 4, "pad": 8, "button_height": 80 }`
     * `buttons`: list of button objects (see below)
   * Button object fields by type:

     * Common: `label` (str), `emoji` (str, optional), `hint` (str, optional)
     * `shell`: `{"type":"shell","cmd":"cmd_or_path","args":["..."], "cwd":"~" (optional)}`
     * `url`: `{"type":"url","url":"https://..."}`
     * `copy`: `{"type":"copy","text":"..."}`
     * `python`: `{"type":"python","code":"print('hello')"}`
     * `reload`: `{"type":"reload"}`
5. **Status & errors**

   * A small status bar (bottom) to show last action or error summary.
   * On action failure, show a message box with a succinct traceback.
6. **Config reload**

   * “Reload” button and `Ctrl+R` to re-read `quickdeck.json` and rebuild the grid without restarting the app.
7. **Defaults**

   * If `quickdeck.json` doesn’t exist, create it with a useful sample.

## Sample `quickdeck.json`

```json
{
  "title": "QuickDeck",
  "geometry": "420x300+40+40",
  "topmost": true,
  "borderless": false,
  "theme": { "bg": "#111827", "fg": "#E5E7EB", "button_bg": "#1F2937" },
  "grid": { "cols": 4, "pad": 8, "button_height": 80 },
  "buttons": [
    { "label": "Browser", "emoji": "🌐", "type": "url", "url": "https://intranet.example" },
    { "label": "Explorer", "emoji": "📁", "type": "shell", "cmd": "explorer", "args": ["%USERPROFILE%"] },
    { "label": "Copy email", "emoji": "📋", "type": "copy", "text": "name.surname@company.com" },
    { "label": "Ping DNS", "emoji": "📡", "type": "shell", "cmd": "cmd", "args": ["/c", "ping -n 4 8.8.8.8"] },
    { "label": "Snippet", "emoji": "📝", "type": "python", "code": "print('Hello from QuickDeck')" },
    { "label": "Reload", "emoji": "🔄", "type": "reload" }
  ]
}
```

## Implementation notes

* Use `tkinter` for UI; set fonts/padding for a neat look. Buttons should expand to fill their grid cell.
* Provide a tiny `Tooltip` helper (pure Tkinter).
* Implement env/tilde expansion for paths/args and `cwd`.
* Use `subprocess.Popen` for `shell` actions; don’t block the UI. On Windows, respect `shell=True` only when necessary (prefer list args).
* Clipboard: `root.clipboard_clear()` / `root.clipboard_append(text)`.
* Python snippet execution: run in a constrained namespace `{}` with `print` captured via `io.StringIO`; show output/errors in a modal dialog or status area.
* Maintain a simple `set_status(msg)` helper.
* Logging: `logging.handlers.RotatingFileHandler('quickdeck.log', maxBytes=200_000, backupCount=2)`; log action start/finish/errors.

## Quality & UX

* Clean dark theme; clear focus ring; hover states if simple to implement.
* Keyboard navigation between buttons via Tab.
* Window remembers previous geometry if you update `geometry` at runtime (optional but nice).

## Acceptance checks (please self-verify)

* Launch with no JSON → creates default config and shows UI.
* `Ctrl+R` or “Reload” button rebuilds buttons after editing JSON.
* Each action type works on Windows/macOS/Linux (within reason for platform-specific commands).
* Window stays topmost; borderless mode allows drag-to-move.
* Errors are surfaced without crashing; log file is written.

## Deliverables

* A single `quickdeck.py` file ready to run with Python 3.8+.
* Inline comments for key functions.
* Brief “How to run” section in a block comment at the top (e.g., `python quickdeck.py [optional path to config]`).

**Now generate the complete `quickdeck.py` implementation per the spec above.**
