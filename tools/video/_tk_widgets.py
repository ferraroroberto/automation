"""
Shared Tk widget builders for the video tools' GUIs.

The progress bar, its percent label and the status line sit at the bottom of
every video tool window with identical layout; building them here keeps that
block in one place.
"""

import tkinter as tk
from tkinter import ttk
from typing import Tuple


def build_progress_section(
    root: tk.Misc, status_var: tk.StringVar
) -> Tuple[ttk.Progressbar, tk.Label, tk.Label]:
    """Pack a percent progress bar, its label and a status line onto ``root``.

    Returns ``(progress_bar, progress_label, status_label)`` for the caller to
    keep as attributes.
    """
    progress_frame = tk.Frame(root)
    progress_frame.pack(fill=tk.X, padx=10, pady=5)
    progress_bar = ttk.Progressbar(progress_frame, maximum=100, mode="determinate")
    progress_bar.pack(fill=tk.X)
    progress_label = tk.Label(progress_frame, text="0%", fg="gray")
    progress_label.pack(anchor="e")

    status_label = tk.Label(root, textvariable=status_var, fg="gray", anchor="w")
    status_label.pack(fill=tk.X, padx=10, pady=(0, 10))
    return progress_bar, progress_label, status_label


def set_progress(root: tk.Misc, progress_bar: ttk.Progressbar, progress_label: tk.Label, percent: float) -> None:
    """Show ``percent`` on the bar and its label, then repaint."""
    progress_bar["value"] = percent
    progress_label.config(text=f"{percent:.1f}%")
    root.update_idletasks()
