#!/usr/bin/env python3
"""
illustrations_formatter_gui.py - Tkinter GUI for image formatting

Layout
------
A notebook with two tabs sits above a shared progress / log panel.

Batch Processing tab
    Convert an entire folder of images to Instagram aspect ratios or
    1920 × 1080.  An *Extend border* checkbox replaces the flat-colour
    padding fill with a pixel-replication technique — useful when the
    image has a coloured border that differs from the auto-detected
    corner colour.

Single Image → 1920 × 1080 tab
    Pick a single file (typically a 1:1 square illustration) and export
    a 1920 × 1080 version immediately.  The same *Extend border* option
    is available.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import logging
import json
import time
from pathlib import Path
from typing import Optional, Tuple, Dict, Any
import sys
import os

try:
    from illustrations_formatter import IllustrationsFormatter, ProcessingResult, parse_color
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from illustrations_formatter import IllustrationsFormatter, ProcessingResult, parse_color


class TextHandler(logging.Handler):
    """Logging handler that writes to a tkinter Text widget via a thread-safe queue."""

    def __init__(self, text_widget, queue_obj):
        super().__init__()
        self.text_widget = text_widget
        self.queue = queue_obj

    def emit(self, record):
        self.queue.put(('log', self.format(record)))


class IllustrationsFormatterGUI:
    """
    Tkinter GUI for :class:`~illustrations_formatter.IllustrationsFormatter`.

    The window is divided into:

    * A :class:`ttk.Notebook` with two tabs:

      - **Batch Processing** – folder-level workflow (Instagram or 1920×1080)
        with an *Extend border* checkbox.
      - **Single Image → 1920×1080** – convert one file immediately, also
        with an *Extend border* checkbox.

    * A shared **Progress / log** panel at the bottom showing a progress bar,
      a status line, and a scrollable log.
    """

    CONFIG_FILE = 'illustrations_formatter_config.json'
    DEFAULT_1920X1080_FOLDER = (
        r'C:\Users\rober\iCloudDrive'
        r'\6LVTQB9699~com~seriflabs~affinitydesigner'
        r'\Roberto\archived_1920x1080'
    )
    DEFAULT_INSTAGRAM_FOLDER = (
        r'C:\Users\rober\iCloudDrive'
        r'\6LVTQB9699~com~seriflabs~affinitydesigner'
        r'\Roberto\archived_IGformat'
    )

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Illustrations Formatter")
        self.root.geometry("840x780")

        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.config = self.load_config()

        # ── Batch tab variables ───────────────────────────────────────────
        self.format_type    = tk.StringVar(value=self.config.get('format_type', 'instagram'))
        self.source_folder  = tk.StringVar(value=self.config.get('source_folder', ''))
        self.dest_folder    = tk.StringVar(value=self.config.get('destination_folder', ''))
        self.aspect_ratio   = tk.StringVar(value=self.config.get('aspect_ratio', '3:4'))
        self.bg_color       = tk.StringVar(value=self.config.get('background_color', ''))
        self.extend_border  = tk.BooleanVar(value=self.config.get('extend_border', False))
        self.save_to_config = tk.BooleanVar(value=False)

        # ── Single-image tab variables ────────────────────────────────────
        self.single_input   = tk.StringVar()
        self.single_out_dir = tk.StringVar(
            value=self.config.get('destination_folder_1920x1080', self.DEFAULT_1920X1080_FOLDER)
        )
        self.single_bg      = tk.StringVar(value=self.config.get('background_color', ''))
        self.single_extend  = tk.BooleanVar(value=self.config.get('extend_border', False))

        self.processing = False
        self.queue: queue.Queue = queue.Queue()

        self._build_ui()
        self.setup_formatter()
        self.update_destination_folder()
        self.root.after(100, self._process_queue)

    # ------------------------------------------------------------------
    # Config persistence
    # ------------------------------------------------------------------

    def load_config(self) -> Dict[str, Any]:
        """Load settings from the JSON config file, merging with built-in defaults."""
        config_path = Path(__file__).parent / self.CONFIG_FILE
        defaults: Dict[str, Any] = {
            'source_folder': '',
            'destination_folder': '',
            'destination_folder_instagram': self.DEFAULT_INSTAGRAM_FOLDER,
            'destination_folder_1920x1080': self.DEFAULT_1920X1080_FOLDER,
            'aspect_ratio': '3:4',
            'background_color': '',
            'format_type': 'instagram',
            'extend_border': False,
        }
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    defaults.update(json.load(f))
        except (json.JSONDecodeError, IOError) as e:
            logging.getLogger(__name__).warning("Warning: Could not load config file: %s", e)
        return defaults

    def save_config(self):
        """Persist current batch-tab settings to the JSON config file."""
        fmt = self.format_type.get()
        config = {
            'source_folder': self.source_folder.get(),
            'destination_folder': self.dest_folder.get(),
            'destination_folder_instagram': self.config.get(
                'destination_folder_instagram', self.DEFAULT_INSTAGRAM_FOLDER
            ),
            'destination_folder_1920x1080': self.config.get(
                'destination_folder_1920x1080', self.DEFAULT_1920X1080_FOLDER
            ),
            'aspect_ratio': self.aspect_ratio.get(),
            'background_color': self.bg_color.get(),
            'format_type': fmt,
            'extend_border': self.extend_border.get(),
        }
        if fmt == '1920x1080':
            config['destination_folder_1920x1080'] = self.dest_folder.get()
        else:
            config['destination_folder_instagram'] = self.dest_folder.get()
            config['destination_folder'] = self.dest_folder.get()

        config_path = Path(__file__).parent / self.CONFIG_FILE
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            logging.getLogger(__name__).warning("Warning: Could not save config file: %s", e)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        """Construct all widgets."""
        outer = ttk.Frame(self.root, padding="10")
        outer.grid(row=0, column=0, sticky='nsew')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        ttk.Label(
            outer, text="Illustrations Formatter", font=('Arial', 16, 'bold')
        ).grid(row=0, column=0, pady=(0, 8))

        nb = ttk.Notebook(outer)
        nb.grid(row=0, column=0, sticky='ew', pady=(32, 0))

        batch_tab  = ttk.Frame(nb, padding="12")
        single_tab = ttk.Frame(nb, padding="12")
        nb.add(batch_tab,  text="  Batch Processing  ")
        nb.add(single_tab, text="  Single Image → 1920×1080  ")

        self._build_batch_tab(batch_tab)
        self._build_single_tab(single_tab)
        self._build_progress_panel(outer, row=1)

    # ── Batch tab ──────────────────────────────────────────────────────

    def _build_batch_tab(self, parent: ttk.Frame):
        """Widgets for the Batch Processing tab."""
        parent.columnconfigure(1, weight=1)

        # Format type
        fmt_frame = ttk.LabelFrame(parent, text="Format Type", padding="8")
        fmt_frame.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 8))
        ttk.Radiobutton(
            fmt_frame, text="Instagram",
            variable=self.format_type, value='instagram',
            command=self.on_format_change,
        ).grid(row=0, column=0, padx=(0, 20))
        ttk.Radiobutton(
            fmt_frame, text="1920 × 1080",
            variable=self.format_type, value='1920x1080',
            command=self.on_format_change,
        ).grid(row=0, column=1)

        # Source folder
        ttk.Label(parent, text="Source Folder:").grid(
            row=1, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=self.source_folder).grid(
            row=1, column=1, sticky='ew', pady=4, padx=(5, 5))
        ttk.Button(parent, text="Browse", command=self._browse_source).grid(
            row=1, column=2, pady=4)

        # Destination folder
        ttk.Label(parent, text="Destination Folder:").grid(
            row=2, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=self.dest_folder).grid(
            row=2, column=1, sticky='ew', pady=4, padx=(5, 5))
        ttk.Button(parent, text="Browse", command=self._browse_dest).grid(
            row=2, column=2, pady=4)

        # Settings
        sf = ttk.LabelFrame(parent, text="Settings", padding="8")
        sf.grid(row=3, column=0, columnspan=3, sticky='ew', pady=8)
        sf.columnconfigure(1, weight=1)

        # Aspect ratio row (shown for Instagram)
        self._ar_label = ttk.Label(sf, text="Aspect Ratio:")
        self._ar_label.grid(row=0, column=0, sticky='w', pady=4)
        self._ar_frame = ttk.Frame(sf)
        self._ar_frame.grid(row=0, column=1, sticky='ew', pady=4)
        ratio_cb = ttk.Combobox(self._ar_frame, textvariable=self.aspect_ratio, width=12)
        ratio_cb['values'] = ['3:4', '4:5', '9:16', '1:1', '16:9']
        ratio_cb.grid(row=0, column=0)
        ttk.Label(self._ar_frame, text="(Instagram: 4:5  |  Stories: 9:16)").grid(
            row=0, column=1, padx=(10, 0))

        # Fixed-size label (shown for 1920x1080)
        self._sz_label = ttk.Label(sf, text="Target Size:")
        self._sz_label.grid(row=0, column=0, sticky='w', pady=4)
        self._sz_label.grid_remove()
        self._sz_frame = ttk.Frame(sf)
        self._sz_frame.grid(row=0, column=1, sticky='ew', pady=4)
        ttk.Label(self._sz_frame, text="1920 × 1080 pixels").grid(row=0, column=0, sticky='w')
        self._sz_frame.grid_remove()

        # Background colour
        ttk.Label(sf, text="Background Color:").grid(row=1, column=0, sticky='w', pady=4)
        cf = ttk.Frame(sf)
        cf.grid(row=1, column=1, sticky='ew', pady=4)
        ttk.Entry(cf, textvariable=self.bg_color, width=18).grid(row=0, column=0)
        ttk.Label(cf, text="Optional: #RRGGBB or R,G,B").grid(
            row=0, column=1, padx=(10, 0))

        # Extend border
        ttk.Checkbutton(
            sf,
            text="Extend border  –  replicate edge pixels instead of flat fill",
            variable=self.extend_border,
        ).grid(row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))

        # Save settings
        ttk.Checkbutton(
            sf,
            text="Save current settings to config file on completion",
            variable=self.save_to_config,
        ).grid(row=3, column=0, columnspan=2, sticky='w', pady=(4, 0))

        # Process button
        self.process_btn = ttk.Button(
            parent, text="Process Images", command=self._start_batch)
        self.process_btn.grid(row=4, column=0, columnspan=3, pady=16)

        self.on_format_change()

    # ── Single image tab ───────────────────────────────────────────────

    def _build_single_tab(self, parent: ttk.Frame):
        """
        Widgets for the Single Image → 1920 × 1080 tab.

        The user selects one source image file and an output folder.  The
        image is scaled (preserving aspect ratio) to fit within 1920 × 1080
        and saved with the same filename.  The *Extend border* option
        replicates the outermost edge pixels into the padding strips —
        especially valuable for square illustrations whose border colour
        is not a plain background shade.
        """
        parent.columnconfigure(1, weight=1)

        ttk.Label(
            parent,
            text=(
                "Convert a single image to 1920 × 1080.  "
                "The image is scaled to fit and centred; "
                "padding strips fill the remaining space."
            ),
            foreground='gray',
            wraplength=640,
            justify='left',
        ).grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 10))

        # Input image
        ttk.Label(parent, text="Input Image:").grid(
            row=1, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=self.single_input).grid(
            row=1, column=1, sticky='ew', pady=4, padx=(5, 5))
        ttk.Button(parent, text="Browse…", command=self._browse_single_input).grid(
            row=1, column=2, pady=4)

        # Output folder
        ttk.Label(parent, text="Output Folder:").grid(
            row=2, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=self.single_out_dir).grid(
            row=2, column=1, sticky='ew', pady=4, padx=(5, 5))
        ttk.Button(parent, text="Browse…", command=self._browse_single_output).grid(
            row=2, column=2, pady=4)

        # Settings
        sf = ttk.LabelFrame(parent, text="Settings", padding="8")
        sf.grid(row=3, column=0, columnspan=3, sticky='ew', pady=8)
        sf.columnconfigure(1, weight=1)

        ttk.Label(sf, text="Background Color:").grid(row=0, column=0, sticky='w', pady=4)
        cf = ttk.Frame(sf)
        cf.grid(row=0, column=1, sticky='ew', pady=4)
        ttk.Entry(cf, textvariable=self.single_bg, width=18).grid(row=0, column=0)
        ttk.Label(cf, text="Optional: #RRGGBB or R,G,B").grid(
            row=0, column=1, padx=(10, 0))

        ttk.Checkbutton(
            sf,
            text=(
                "Extend border  –  replicate edge pixels instead of flat fill\n"
                "  Recommended when the image has a coloured border that differs\n"
                "  from the detected background (e.g. a coloured ocean edge)."
            ),
            variable=self.single_extend,
        ).grid(row=1, column=0, columnspan=2, sticky='w', pady=(8, 0))

        # Convert button
        self.convert_btn = ttk.Button(
            parent, text="Convert Image", command=self._start_single)
        self.convert_btn.grid(row=4, column=0, columnspan=3, pady=16)

    # ── Shared progress panel ──────────────────────────────────────────

    def _build_progress_panel(self, parent: ttk.Frame, row: int):
        """Progress bar, status label, and scrollable log — shared by both tabs."""
        pf = ttk.LabelFrame(parent, text="Progress", padding="10")
        pf.grid(row=row, column=0, sticky='nsew', pady=(8, 0))
        pf.columnconfigure(0, weight=1)
        pf.rowconfigure(2, weight=1)
        parent.rowconfigure(row, weight=1)

        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(pf, variable=self.progress_var, maximum=100).grid(
            row=0, column=0, sticky='ew', pady=(0, 4))

        self.status_label = ttk.Label(pf, text="Ready", anchor='w')
        self.status_label.grid(row=1, column=0, sticky='ew', pady=(0, 4))

        self.log_text = scrolledtext.ScrolledText(pf, height=10, wrap=tk.WORD)
        self.log_text.grid(row=2, column=0, sticky='nsew')
        self.log_text.tag_config('SUCCESS', foreground='green')
        self.log_text.tag_config('ERROR',   foreground='red')

    # ------------------------------------------------------------------
    # Formatter setup
    # ------------------------------------------------------------------

    def setup_formatter(self):
        """Create the :class:`IllustrationsFormatter` and attach a GUI log handler."""
        logger = logging.getLogger('IllustrationsFormatter')
        logger.setLevel(logging.INFO)
        handler = TextHandler(self.log_text, self.queue)
        handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S'
        ))
        logger.addHandler(handler)
        self.formatter = IllustrationsFormatter(logger)

    # ------------------------------------------------------------------
    # Batch tab helpers
    # ------------------------------------------------------------------

    def update_destination_folder(self):
        """Swap the destination entry to the saved folder for the current format."""
        if self.format_type.get() == '1920x1080':
            default = self.config.get('destination_folder_1920x1080', self.DEFAULT_1920X1080_FOLDER)
            if not self.dest_folder.get() or self.dest_folder.get() == self.config.get(
                    'destination_folder_instagram', ''):
                self.dest_folder.set(default)
        else:
            default = self.config.get('destination_folder_instagram', self.DEFAULT_INSTAGRAM_FOLDER)
            if not self.dest_folder.get() or self.dest_folder.get() == self.config.get(
                    'destination_folder_1920x1080', ''):
                self.dest_folder.set(default)

    def on_format_change(self):
        """Show / hide aspect-ratio or fixed-size info based on the selected format."""
        self.update_destination_folder()
        if self.format_type.get() == '1920x1080':
            self._ar_label.grid_remove()
            self._ar_frame.grid_remove()
            self._sz_label.grid()
            self._sz_frame.grid()
        else:
            self._sz_label.grid_remove()
            self._sz_frame.grid_remove()
            self._ar_label.grid()
            self._ar_frame.grid()

    def _browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)

    def _browse_dest(self):
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_folder.set(folder)

    def _validate_batch(self) -> Optional[str]:
        if not self.source_folder.get():
            return "Please select a source folder."
        if not self.dest_folder.get():
            return "Please select a destination folder."
        if not Path(self.source_folder.get()).exists():
            return "Source folder does not exist."
        if self.format_type.get() == 'instagram':
            if not self.aspect_ratio.get():
                return "Please specify an aspect ratio."
            try:
                self.formatter.parse_aspect_ratio(self.aspect_ratio.get())
            except ValueError:
                return "Invalid aspect ratio format."
        if self.bg_color.get():
            try:
                parse_color(self.bg_color.get())
            except ValueError:
                return "Invalid color format. Use #RRGGBB or R,G,B."
        return None

    def _start_batch(self):
        """Validate inputs and launch batch processing in a background thread."""
        if self.processing:
            messagebox.showwarning("Busy", "Already processing images!")
            return
        err = self._validate_batch()
        if err:
            messagebox.showerror("Validation Error", err)
            return
        self._reset_progress("Processing batch…")
        self.processing = True
        self.process_btn.config(state='disabled')
        threading.Thread(target=self._run_batch, daemon=True).start()

    def _run_batch(self):
        """Worker for batch processing (runs in a background thread)."""
        try:
            bg = parse_color(self.bg_color.get()) if self.bg_color.get() else None
            if self.format_type.get() == '1920x1080':
                result = self.formatter.process_folder_fixed_size(
                    self.source_folder.get(),
                    self.dest_folder.get(),
                    1920, 1080,
                    bg,
                    extend_border=self.extend_border.get(),
                    progress_callback=self._progress_cb,
                )
            else:
                result = self.formatter.process_folder(
                    self.source_folder.get(),
                    self.dest_folder.get(),
                    self.aspect_ratio.get(),
                    bg,
                    extend_border=self.extend_border.get(),
                    progress_callback=self._progress_cb,
                )
            self.queue.put(('complete_batch', result))
        except Exception as e:
            self.queue.put(('error', str(e)))

    # ------------------------------------------------------------------
    # Single-image tab helpers
    # ------------------------------------------------------------------

    def _browse_single_input(self):
        """Open a file picker for the single input image."""
        path = filedialog.askopenfilename(
            title="Select Input Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.single_input.set(path)

    def _browse_single_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.single_out_dir.set(folder)

    def _validate_single(self) -> Optional[str]:
        if not self.single_input.get():
            return "Please select an input image."
        if not Path(self.single_input.get()).exists():
            return "Input image file does not exist."
        if not self.single_out_dir.get():
            return "Please select an output folder."
        if self.single_bg.get():
            try:
                parse_color(self.single_bg.get())
            except ValueError:
                return "Invalid color format. Use #RRGGBB or R,G,B."
        return None

    def _start_single(self):
        """Validate inputs and launch single-image conversion in a background thread."""
        if self.processing:
            messagebox.showwarning("Busy", "Already processing an image!")
            return
        err = self._validate_single()
        if err:
            messagebox.showerror("Validation Error", err)
            return
        self._reset_progress("Converting single image…")
        self.processing = True
        self.convert_btn.config(state='disabled')
        threading.Thread(target=self._run_single, daemon=True).start()

    def _run_single(self):
        """Worker for single-image conversion (runs in a background thread)."""
        try:
            t0 = time.time()
            input_path  = Path(self.single_input.get())
            output_path = Path(self.single_out_dir.get()) / input_path.name
            bg = parse_color(self.single_bg.get()) if self.single_bg.get() else None

            self.formatter.convert_single_to_1920x1080(
                input_path,
                output_path,
                background_color=bg,
                extend_border=self.single_extend.get(),
            )
            elapsed = time.time() - t0
            self.queue.put(('complete_single', (str(output_path), elapsed)))
        except Exception as e:
            self.queue.put(('error', str(e)))

    # ------------------------------------------------------------------
    # Shared progress / queue helpers
    # ------------------------------------------------------------------

    def _reset_progress(self, status: str = "Ready"):
        self.log_text.delete('1.0', tk.END)
        self.progress_var.set(0)
        self.status_label.config(text=status)

    def _progress_cb(self, current: int, total: int, message: str):
        pct = (current / total * 100) if total > 0 else 0
        self.queue.put(('progress', (pct, f"{current}/{total} – {message}")))

    def _process_queue(self):
        """Drain the inter-thread queue and update all GUI elements."""
        try:
            while True:
                kind, data = self.queue.get_nowait()

                if kind == 'log':
                    self.log_text.insert(tk.END, data + '\n')
                    self.log_text.see(tk.END)

                elif kind == 'progress':
                    pct, msg = data
                    self.progress_var.set(pct)
                    self.status_label.config(text=msg)

                elif kind == 'complete_batch':
                    self._on_batch_complete(data)

                elif kind == 'complete_single':
                    out_path, elapsed = data
                    self._on_single_complete(out_path, elapsed)

                elif kind == 'error':
                    self.processing = False
                    self.process_btn.config(state='normal')
                    self.convert_btn.config(state='normal')
                    self.status_label.config(text="Error occurred!")
                    messagebox.showerror("Processing Error", data)

        except queue.Empty:
            pass

        self.root.after(100, self._process_queue)

    def _on_batch_complete(self, result: ProcessingResult):
        self.processing = False
        self.process_btn.config(state='normal')
        self.progress_var.set(100)

        if self.save_to_config.get():
            self.save_config()

        summary = (
            f"\nBatch complete!\n"
            f"Total: {result.total_images}  |  "
            f"Done: {result.successful}  |  "
            f"Skipped: {result.skipped}  |  "
            f"Failed: {result.failed}\n"
            f"Time: {result.elapsed_time:.2f} s\n"
        )
        self.log_text.insert(tk.END, summary, 'SUCCESS')
        self.log_text.see(tk.END)
        self.status_label.config(text="Done!")

        if result.failed == 0:
            messagebox.showinfo(
                "Success",
                f"Successfully processed {result.successful} image(s).",
            )
        else:
            messagebox.showwarning(
                "Done with errors",
                f"{result.successful} image(s) processed.\n"
                f"{result.failed} failed — see log for details.",
            )

    def _on_single_complete(self, out_path: str, elapsed: float):
        self.processing = False
        self.convert_btn.config(state='normal')
        self.progress_var.set(100)

        msg = f"\nConverted in {elapsed:.2f} s → {out_path}\n"
        self.log_text.insert(tk.END, msg, 'SUCCESS')
        self.log_text.see(tk.END)
        self.status_label.config(text="Done!")
        messagebox.showinfo("Success", f"Image saved to:\n{out_path}")


def main():
    """Launch the Illustrations Formatter GUI application."""
    root = tk.Tk()
    IllustrationsFormatterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
