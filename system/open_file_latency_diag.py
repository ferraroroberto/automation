"""
Compare raw file read latency vs. Windows Shell stack work for the same path.

The Shell path is designed to approximate what happens when Explorer / a common
file-open dialog touches a file: namespace item creation, icon + overlay
resolution (shell icon / overlay handlers), and property-store reads (property
handlers for that file type).

Run from the project venv (pywin32 is already in requirements.txt):

    .venv\\Scripts\\python.exe system\\open_file_latency_diag.py
        (opens a tkinter file picker if you omit --file)

    .venv\\Scripts\\python.exe system\\open_file_latency_diag.py --file C:\\path\\to\\file.pptx

Optional: average multiple iterations:

    .venv\\Scripts\\python.exe system\\open_file_latency_diag.py --iterations 5

Dependencies: ``pywin32`` (``pip install -r requirements.txt``); ``tkinter`` for the
file picker (usually bundled with Python on Windows).
"""

from __future__ import annotations

import argparse
import statistics
import sys
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

# --- Windows / COM (pywin32) -------------------------------------------------

try:
    import pythoncom
    import win32gui
    from win32com.propsys import propsys
    from win32com.shell import shell, shellcon
except ImportError as e:
    print(
        "This script requires pywin32. Use the project virtualenv, e.g.:\n"
        "  .venv\\Scripts\\python.exe system\\open_file_latency_diag.py\n"
        "Or: pip install pywin32",
        file=sys.stderr,
    )
    raise SystemExit(1) from e

IID_IShellItem = "{43826D1E-E718-42EE-BC55-A1E261C37BFE}"
SHGFI_ADDOVERLAYS = 0x00000800


def _shell_touch_file(path: Path) -> None:
    """
    Exercise Shell for ``path`` (must exist). Does not show a dialog.
    """
    s = str(path.resolve())

    item = shell.SHCreateItemFromParsingName(s, None, IID_IShellItem)

    item.GetDisplayName(shellcon.SIGDN_FILESYSPATH)
    item.GetDisplayName(shellcon.SIGDN_NORMALDISPLAY)

    item.GetAttributes(0xFFFFFFFF)

    flags = (
        shellcon.SHGFI_ICON
        | shellcon.SHGFI_LARGEICON
        | shellcon.SHGFI_TYPENAME
        | shellcon.SHGFI_DISPLAYNAME
        | SHGFI_ADDOVERLAYS
    )
    info = shell.SHGetFileInfo(s, 0, flags)
    if info and isinstance(info[1], tuple):
        hicon = info[1][0]
        if hicon:
            win32gui.DestroyIcon(hicon)

    store = propsys.SHGetPropertyStoreFromParsingName(
        s, None, 0, propsys.IID_IPropertyStore
    )
    n = store.GetCount()
    for i in range(n):
        key = store.GetAt(i)
        store.GetValue(key)


def bench_raw_read(path: Path, iterations: int) -> list[float]:
    times: list[float] = []
    for _ in range(iterations):
        t0 = time.perf_counter()
        _ = path.read_bytes()
        times.append(time.perf_counter() - t0)
    return times


def bench_shell(path: Path, iterations: int) -> list[float]:
    times: list[float] = []
    pythoncom.CoInitialize()
    try:
        for _ in range(iterations):
            t0 = time.perf_counter()
            _shell_touch_file(path)
            times.append(time.perf_counter() - t0)
    finally:
        pythoncom.CoUninitialize()
    return times


def _fmt_ms(seconds: float) -> str:
    return f"{seconds * 1000:.3f} ms"


def _ask_file_path() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    try:
        chosen = filedialog.askopenfilename(
            title="Select file for Shell vs raw read benchmark",
            filetypes=[
                ("Presentations", "*.pptx *.ppt"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()
    return Path(chosen) if chosen else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Raw read vs Shell API latency for one file (pywin32)."
    )
    parser.add_argument(
        "--file",
        "-f",
        type=Path,
        default=None,
        help="Test file path (if omitted, a tkinter Open dialog is shown)",
    )
    parser.add_argument(
        "--iterations",
        "-n",
        type=int,
        default=3,
        metavar="N",
        help="Repeat each test N times (default: 3)",
    )
    args = parser.parse_args()
    n: int = max(1, args.iterations)

    path: Path | None = args.file
    if path is None:
        path = _ask_file_path()
        if path is None:
            print("No file selected.", file=sys.stderr)
            return 1

    if not path.is_file():
        print(f"File not found: {path.resolve()}", file=sys.stderr)
        return 2

    print(f"Target: {path.resolve()}")
    print(f"Size:   {path.stat().st_size} bytes")
    print(f"Runs:   {n} per test (reported: mean, stdev)")
    print()

    raw_times = bench_raw_read(path, n)
    shell_times = bench_shell(path, n)

    raw_mean = statistics.mean(raw_times)
    sh_mean = statistics.mean(shell_times)
    raw_stdev = statistics.stdev(raw_times) if n > 1 else 0.0
    sh_stdev = statistics.stdev(shell_times) if n > 1 else 0.0

    print("--- Baseline: raw binary read (Python / kernel read path) ---")
    for i, t in enumerate(raw_times, 1):
        print(f"  Run {i}: {_fmt_ms(t)}")
    print(f"  Mean: {_fmt_ms(raw_mean)}  (stdev {_fmt_ms(raw_stdev)})")
    print()

    print(
        "--- Shell API simulation (IShellItem + SHGetFileInfo w/ overlays "
        "+ full IPropertyStore via SHGetPropertyStoreFromParsingName) ---"
    )
    print("  (No dialog is shown; this still drives much of the same Shell code.)")
    for i, t in enumerate(shell_times, 1):
        print(f"  Run {i}: {_fmt_ms(t)}")
    print(f"  Mean: {_fmt_ms(sh_mean)}  (stdev {_fmt_ms(sh_stdev)})")
    print()

    ratio = sh_mean / raw_mean if raw_mean > 0 else float("inf")
    delta = sh_mean - raw_mean
    print("--- Comparison ---")
    print(f"  Shell mean / Raw mean: {ratio:.2f}x")
    print(f"  Shell mean - Raw mean: {_fmt_ms(delta)}")
    print()
    print(
        "Interpretation: a large gap suggests time spent outside simple read(2)/"
        "cache (e.g. Shell handlers, overlays, property handlers, AV hooks on "
        "Shell APIs). Use ShellExView / Autoruns / AV exclusions and Process "
        "Monitor to narrow down further."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
