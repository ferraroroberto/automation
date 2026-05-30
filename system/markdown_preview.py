#!/usr/bin/env python3
"""
markdown_preview.py

A self-contained, tray-resident markdown previewer. Renders a Markdown file as
GitHub-style HTML inside an Edge WebView2 window (via pywebview), with a
per-window light/dark toggle that defaults to the current Windows app theme.

Runs from the system tray. Closing the window hides it back to the tray
(the process stays alive); reopen it from the tray menu. Opening a different
file reloads the same window. The open file is watched for changes on disk and
the preview re-renders within ~1s of a save. Quit only from the tray menu.

UI (tray menu):
  * Open        - show / raise the window (default action)
  * Open file...- pick a .md and load it into the window
  * Quit        - exit the app

Requirements:
    pip install pywebview markdown Pygments pystray Pillow

Usage:
    pythonw markdown_preview.py [path\\to\\file.md]
"""

import argparse
import ctypes
import logging
import threading
import time
from ctypes import wintypes
from pathlib import Path
from typing import Optional

import markdown as md
import pystray
import webview
from PIL import Image, ImageDraw
from pygments.formatters import HtmlFormatter

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

MARKDOWN_EXTENSIONS = ["fenced_code", "tables", "toc", "codehilite", "sane_lists"]


def detect_windows_theme() -> str:
    """Return 'dark' or 'light' from the Windows 'apps' theme setting."""
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        apps_use_light, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return "light" if apps_use_light else "dark"
    except OSError:
        return "light"


def _pygments_css() -> str:
    """Syntax-highlighting CSS for both themes, scoped by data-theme."""
    light = HtmlFormatter(style="default").get_style_defs(
        '[data-theme="light"] .codehilite'
    )
    dark = HtmlFormatter(style="monokai").get_style_defs(
        '[data-theme="dark"] .codehilite'
    )
    return light + "\n" + dark


# GitHub-flavored CSS. Colors mirror GitHub's light/dark tokens; selectors are
# driven by the data-theme attribute on <html> so the toggle is pure CSS.
_BASE_CSS = """
:root {
  --bg: #ffffff; --fg: #1f2328; --muted: #59636e; --border: #d1d9e0;
  --code-bg: #f6f8fa; --quote-fg: #59636e; --link: #0969da;
  --table-alt: #f6f8fa; --hr: #d1d9e0;
}
[data-theme="dark"] {
  --bg: #0d1117; --fg: #e6edf3; --muted: #9198a1; --border: #3d444d;
  --code-bg: #151b23; --quote-fg: #9198a1; --link: #4493f8;
  --table-alt: #151b23; --hr: #3d444d;
}
html, body { margin: 0; padding: 0; background: var(--bg); color: var(--fg); }
* { box-sizing: border-box; }
.toolbar {
  position: sticky; top: 0; z-index: 10;
  display: flex; align-items: center; gap: 10px;
  padding: 6px 16px; background: var(--bg);
  border-bottom: 1px solid var(--border);
  font: 12px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
.toolbar .fname { color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.toolbar button {
  cursor: pointer; border: 1px solid var(--border); background: var(--code-bg);
  color: var(--fg); border-radius: 6px; padding: 3px 10px; font-size: 13px; line-height: 1.4;
}
.toolbar button:hover { border-color: var(--muted); }
.markdown-body {
  max-width: 980px; margin: 0 auto; padding: 24px 32px 64px;
  font: 16px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans", Helvetica, Arial, sans-serif;
  word-wrap: break-word;
}
.markdown-body h1, .markdown-body h2 { border-bottom: 1px solid var(--border); padding-bottom: .3em; }
.markdown-body h1 { font-size: 2em; margin: .67em 0; }
.markdown-body h2 { font-size: 1.5em; margin-top: 24px; }
.markdown-body h3 { font-size: 1.25em; margin-top: 24px; }
.markdown-body h4, .markdown-body h5, .markdown-body h6 { margin-top: 24px; }
.markdown-body p, .markdown-body ul, .markdown-body ol, .markdown-body blockquote, .markdown-body table { margin-top: 0; margin-bottom: 16px; }
.markdown-body a { color: var(--link); text-decoration: none; }
.markdown-body a:hover { text-decoration: underline; }
.markdown-body code {
  background: var(--code-bg); border-radius: 6px; padding: .2em .4em;
  font: 85% ui-monospace, SFMono-Regular, "SF Mono", Consolas, "Liberation Mono", monospace;
}
.markdown-body pre {
  background: var(--code-bg); border-radius: 6px; padding: 16px;
  overflow: auto; line-height: 1.45;
}
.markdown-body pre code { background: transparent; padding: 0; font-size: 85%; }
.markdown-body .codehilite { background: var(--code-bg); border-radius: 6px; margin-bottom: 16px; }
.markdown-body .codehilite pre { margin: 0; background: transparent; }
.markdown-body blockquote {
  color: var(--quote-fg); border-left: .25em solid var(--border);
  padding: 0 1em; margin-left: 0;
}
.markdown-body table { border-collapse: collapse; display: block; width: max-content; max-width: 100%; overflow: auto; }
.markdown-body table th, .markdown-body table td { border: 1px solid var(--border); padding: 6px 13px; }
.markdown-body table tr:nth-child(2n) { background: var(--table-alt); }
.markdown-body img { max-width: 100%; }
.markdown-body hr { height: .25em; background: var(--hr); border: 0; margin: 24px 0; }
.markdown-body ul, .markdown-body ol { padding-left: 2em; }
"""

# Built once: the stylesheet is constant across renders.
_FULL_CSS = _BASE_CSS + _pygments_css()

_PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en" data-theme="{theme}">
<head>
<meta charset="utf-8">
<style>{css}</style>
</head>
<body>
<div class="toolbar">
  <button id="theme-btn" onclick="toggleTheme()" title="Toggle light/dark">{glyph}</button>
  <span class="fname">{fname}</span>
</div>
<article class="markdown-body">{body}</article>
<script>
function applyGlyph() {{
  var t = document.documentElement.getAttribute('data-theme');
  document.getElementById('theme-btn').textContent = (t === 'dark') ? '\\u2600' : '\\u263d';
}}
function toggleTheme() {{
  var el = document.documentElement;
  var next = (el.getAttribute('data-theme') === 'dark') ? 'light' : 'dark';
  el.setAttribute('data-theme', next);
  applyGlyph();
  if (window.pywebview && window.pywebview.api && window.pywebview.api.set_theme) {{
    window.pywebview.api.set_theme(next);
  }}
}}
applyGlyph();
</script>
</body>
</html>"""


class Api:
    """Exposed to JS so the window's theme survives a live-reload re-render."""

    def __init__(self, theme: str):
        self.theme = theme

    def set_theme(self, theme: str) -> None:
        self.theme = theme


class MarkdownPreviewApp:
    _ICON_SIZE = 64
    _MUTEX_NAME = "markdown_preview_singleton_v1"
    # Named auto-reset event: a second instance sets it; the running instance
    # wakes and shows its window. Lives alongside the mutex so "already running"
    # becomes "raise window" rather than a dead-end error box.
    _SHOW_EVENT_NAME = "markdown_preview_show_v1"
    _ERROR_ALREADY_EXISTS = 183

    def __init__(self, initial_file: Optional[Path]):
        self.current_file: Optional[Path] = initial_file
        self._last_mtime: Optional[float] = None
        self.api = Api(detect_windows_theme())

        self.window: Optional[webview.Window] = None
        self._quitting = False
        self._mutex: Optional[int] = None

        self._icon = pystray.Icon(
            "markdown_preview",
            self._make_icon(),
            "Markdown Preview",
            pystray.Menu(
                pystray.MenuItem("Open", self._on_open, default=True),
                pystray.MenuItem("Open file…", self._on_open_file),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._on_quit),
            ),
        )

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------

    def _render(self) -> str:
        theme = self.api.theme
        glyph = "☀" if theme == "dark" else "☽"
        if self.current_file and self.current_file.is_file():
            try:
                text = self.current_file.read_text(encoding="utf-8")
            except OSError as exc:
                body = f"<h1>Cannot read file</h1><p>{exc}</p>"
                fname = self.current_file.name
            else:
                body = md.markdown(text, extensions=MARKDOWN_EXTENSIONS)
                fname = str(self.current_file)
        else:
            body = (
                "<h1>Markdown Preview</h1>"
                "<p>Open a <code>.md</code> file from the tray menu "
                "(<strong>Open file…</strong>).</p>"
            )
            fname = "No file"
        return _PAGE_TEMPLATE.format(
            theme=theme,
            css=_FULL_CSS,
            glyph=glyph,
            fname=fname,
            body=body,
        )

    def _load_current(self) -> None:
        if self.window is None:
            return
        self._last_mtime = self._file_mtime()
        self.window.load_html(self._render())

    def _file_mtime(self) -> Optional[float]:
        if self.current_file and self.current_file.is_file():
            try:
                return self.current_file.stat().st_mtime
            except OSError:
                return None
        return None

    # ------------------------------------------------------------------
    # Tray icon
    # ------------------------------------------------------------------

    def _make_icon(self) -> Image.Image:
        """Draw a small document icon with a down-arrow (markdown glyph)."""
        size = self._ICON_SIZE
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        paper = (255, 255, 255)
        edge = (55, 65, 81)
        accent = (96, 165, 250)
        # page with a folded corner
        draw.polygon(
            [(14, 6), (42, 6), (52, 16), (52, 58), (14, 58)],
            fill=paper,
            outline=edge,
        )
        draw.line([(42, 6), (42, 16), (52, 16)], fill=edge, width=2)
        # down arrow (the "markdown" cue)
        draw.line([(33, 24), (33, 44)], fill=accent, width=4)
        draw.polygon([(25, 38), (41, 38), (33, 50)], fill=accent)
        return img

    def _on_open(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.window is not None:
            self.window.show()

    def _on_open_file(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.window is None:
            return
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Markdown (*.md;*.markdown;*.mdown)", "All files (*.*)"),
        )
        if not result:
            return
        self.current_file = Path(result[0])
        self._load_current()
        self.window.show()

    def _on_quit(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        logger.info("Shutting down Markdown Preview")
        self._quitting = True
        if self.window is not None:
            self.window.destroy()
        else:
            icon.stop()

    # ------------------------------------------------------------------
    # Window lifecycle
    # ------------------------------------------------------------------

    def _on_closing(self) -> bool:
        """Close button hides to tray; only Quit actually destroys."""
        if self._quitting:
            return True
        if self.window is not None:
            self.window.hide()
        return False

    def _watch_file(self) -> None:
        """Poll the open file's mtime and re-render on change."""
        while not self._quitting:
            time.sleep(0.7)
            if self.window is None or self.current_file is None:
                continue
            mtime = self._file_mtime()
            if mtime is not None and mtime != self._last_mtime:
                self._last_mtime = mtime
                self.window.load_html(self._render())

    # ------------------------------------------------------------------
    # Single-instance lock + show-event IPC
    # ------------------------------------------------------------------

    def _acquire_lock(self) -> bool:
        """Take the named mutex. If another instance already holds it, signal
        it to raise its window and return False so this process exits cleanly.
        """
        k = ctypes.windll.kernel32
        k.CreateMutexW.restype = wintypes.HANDLE
        k.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        self._mutex = k.CreateMutexW(None, True, self._MUTEX_NAME)
        if k.GetLastError() != self._ERROR_ALREADY_EXISTS:
            return True
        # Another instance is running — ask it to show its window.
        EVENT_MODIFY_STATE = 0x0002
        k.OpenEventW.restype = wintypes.HANDLE
        k.OpenEventW.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.LPCWSTR]
        h = k.OpenEventW(EVENT_MODIFY_STATE, False, self._SHOW_EVENT_NAME)
        if h:
            k.SetEvent(h)
            k.CloseHandle(h)
        else:
            # Running instance is in a bad state; fall back to an error box.
            ctypes.windll.user32.MessageBoxW(
                0, "Markdown Preview is already running.", "Markdown Preview", 0x40
            )
        return False

    def _monitor_show_event(self) -> None:
        """Daemon thread: wait for a second instance to signal us, then show the window."""
        k = ctypes.windll.kernel32
        k.CreateEventW.restype = wintypes.HANDLE
        k.CreateEventW.argtypes = [
            wintypes.LPVOID, wintypes.BOOL, wintypes.BOOL, wintypes.LPCWSTR
        ]
        # Auto-reset (bManualReset=False), initially non-signalled.
        h = k.CreateEventW(None, False, False, self._SHOW_EVENT_NAME)
        if not h:
            logger.warning("Could not create show-event; bring-to-front IPC disabled")
            return
        WAIT_OBJECT_0 = 0x00000000
        try:
            while not self._quitting:
                result = k.WaitForSingleObject(h, 500)
                if result == WAIT_OBJECT_0 and self.window is not None:
                    self.window.show()
        finally:
            k.CloseHandle(h)

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------

    def run(self) -> None:
        if not self._acquire_lock():
            return

        logger.info("Starting Markdown Preview (tray)")
        self.window = webview.create_window(
            "Markdown Preview",
            html=self._render(),
            js_api=self.api,
            width=1000,
            height=820,
            text_select=True,
        )
        self.window.events.closing += self._on_closing
        self._last_mtime = self._file_mtime()

        threading.Thread(target=self._icon.run, daemon=True).start()
        threading.Thread(target=self._watch_file, daemon=True).start()
        threading.Thread(target=self._monitor_show_event, daemon=True).start()

        webview.start(gui="edgechromium")

        # start() returns once the window is destroyed (Quit)
        self._quitting = True
        try:
            self._icon.stop()
        except Exception:
            pass


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Tray-resident GitHub-style markdown previewer.")
    p.add_argument("file", nargs="?", default=None, help="Path to a .md file to open on launch")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    initial = Path(args.file).expanduser() if args.file else None
    MarkdownPreviewApp(initial).run()


if __name__ == "__main__":
    main()
