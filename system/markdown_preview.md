# Markdown Preview

A self-contained, tray-resident Markdown previewer. It renders a `.md` file as GitHub-style HTML inside an Edge WebView2 window (via [pywebview](https://pywebview.flowrl.com/)), with a per-window light/dark toggle. Built for locked-down machines where installing a Markdown viewer (Typora, MarkText, Obsidian, …) is blocked — it runs as plain Python through the project's `.venv`, so there is no extra executable to install.

## Features

- **System Tray App**: runs in the system tray; click the icon to open the window, close the window to hide it back to the tray, Quit from the tray menu. Single-instance — launching it twice just shows a message and exits.
- **GitHub-style rendering**: headings, tables, fenced code with syntax highlighting (Pygments), blockquotes, images. Styling mirrors GitHub's light/dark color tokens.
- **Light/dark toggle**: a button in the top toolbar flips the theme instantly. New windows default to the current **Windows app theme**.
- **Live reload**: the open file is watched on disk; saving it from your editor refreshes the preview within ~1 second.
- **Open any file**: pass a path on launch, or pick one from the tray's **Open file…** dialog. Opening a different file reloads the same window.

## Requirements

- Python 3
- `pywebview`, `markdown`, `Pygments`, `pystray`, `Pillow`
- Windows with the **Edge WebView2 runtime** (preinstalled on Windows 11). If absent, install the Microsoft Evergreen WebView2 runtime (per-user install, no admin needed).

Install from the project requirements:

```
pip install -r ../../requirements.txt
```

## Usage

Launch silently into the tray (no console window) via the `.bat`:

```
markdown_preview.bat
markdown_preview.bat path\to\file.md
```

Or run directly:

```
pythonw markdown_preview.py [path\to\file.md]
```

Without a file argument the window shows a short placeholder until you pick a file from **Open file…**.

### Wiring to "Open with…"

Because the `.bat` forwards its first argument to the script, you can set it as a Windows "Open with…" handler for `.md` files (point the association at `markdown_preview.bat`).

## Tray menu

- **Open** — show / raise the window (default action; also the icon click).
- **Open file…** — pick a `.md` and load it into the window.
- **Quit** — exit the app and release the single-instance lock.

## How it works

1. A single-instance lock is taken via a Windows named mutex; a second launch detects the existing mutex, shows a message box, and exits.
2. The Markdown is converted to HTML in Python (`markdown` + `Pygments`) and wrapped in an embedded, offline GitHub-style stylesheet (both light and dark variants).
3. The HTML is displayed in a pywebview window backed by Edge WebView2. The light/dark toggle flips a `data-theme` attribute via JavaScript; the choice is mirrored back to Python so it survives a live-reload re-render.
4. pywebview's GUI loop owns the main thread; the pystray tray icon and the file watcher each run on a background thread.
5. The window's close event is intercepted to **hide** the window instead of destroying it, so the process stays in the tray. **Quit** destroys the window, which ends the loop and releases the lock.
