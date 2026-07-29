#!/usr/bin/env python3
"""
Folder Searcher - a fast way to find and open folders by name.

This module is the Tkinter/tray shell only. All matching, scanning,
persistence and pruning logic lives in ``foldersearcher_core``, which has no
GUI dependencies and is covered by ``test_foldersearcher_core.py``.

The window is a two-tab notebook:

- **Search / Access** (selected on open) searches the saved index without
  rescanning, and opens a result in Explorer on double-click.
- **Folders** manages the configured roots and rebuilds the index.

The app runs from the system tray. Click the tray icon to open the window;
closing the window minimizes it back to the tray. Quit from the tray menu.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ctypes
import logging
import os
import subprocess
import sys
from ctypes import wintypes
from typing import List, Optional

import pystray
from PIL import Image, ImageDraw

from foldersearcher_core import (
    FolderIndex,
    FolderSearcherConfig,
    SearchResult,
    load_config,
    normalize_root,
    save_config,
    search,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


class FolderSearcher:
    """
    Main application class for folder searching functionality.

    Lifecycle:
    - The pystray icon is created up front but the Tk root is only built
      inside the pystray setup callback (so tkinter lives on the same
      thread that owns the icon's message loop).
    - The main window is a Toplevel built lazily on the first Open click
      and reused thereafter — close hides it (withdraw), Quit shuts down.
    """

    _ICON_SIZE = 64
    # Single-instance via a Windows named mutex. A fixed loopback TCP port is
    # unreliable here: Windows reserves large high-port ranges (Hyper-V/WSL),
    # so bind() can fail with WinError 10013 even when no instance is running.
    _MUTEX_NAME = "foldersearcher_singleton_v1"
    _ERROR_ALREADY_EXISTS = 183

    def __init__(self):
        """Initialize the FolderSearcher application."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(script_dir, "foldersearcher.json")
        self.structure_file = os.path.join(script_dir, "folder_structure.txt")

        # Configuration and index (both GUI-free)
        self.config: FolderSearcherConfig = load_config(self.config_file, self.structure_file)
        self.index = FolderIndex()

        # Current result set, parallel to the results listbox rows
        self.results: List[SearchResult] = []

        # Tk objects — created in the pystray setup callback
        self.root: Optional[tk.Tk] = None
        self.window: Optional[tk.Toplevel] = None
        self._widgets_built = False

        # Tk variables — created once the Tk root exists
        self.search_var: Optional[tk.StringVar] = None
        self.skip_depth_var: Optional[tk.IntVar] = None
        self.prune_var: Optional[tk.BooleanVar] = None
        self.status_var: Optional[tk.StringVar] = None

        # Widgets populated in setup_ui
        self.notebook: Optional[ttk.Notebook] = None
        self.search_entry: Optional[ttk.Entry] = None
        self.results_listbox: Optional[tk.Listbox] = None
        self.roots_listbox: Optional[tk.Listbox] = None

        # Single-instance mutex handle — kept for the process lifetime
        self._mutex: Optional[int] = None

        # Tray icon
        self._icon = pystray.Icon(
            "foldersearcher",
            self._make_icon(),
            "Folder Searcher",
            pystray.Menu(
                pystray.MenuItem("Open", self._on_open, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Quit", self._on_quit),
            ),
        )

        logger.info("Folder Searcher application initialized")

    # ------------------------------------------------------------------
    # Tray icon
    # ------------------------------------------------------------------

    def _make_icon(self) -> Image.Image:
        """Draw a small folder-with-magnifier tray icon."""
        size = self._ICON_SIZE
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        # folder tab
        draw.polygon(
            [(8, 18), (26, 18), (32, 24), (56, 24), (56, 26), (8, 26)],
            fill=(245, 158, 11),
        )
        # folder body
        draw.rectangle([8, 24, 56, 54], fill=(251, 191, 36))
        # magnifier glass
        draw.ellipse([26, 28, 44, 46], outline=(31, 41, 55), width=3)
        draw.line([(42, 44), (52, 54)], fill=(31, 41, 55), width=4)
        return img

    def _on_open(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.root:
            self.root.after(0, self._show_window)

    def _on_quit(self, icon: pystray.Icon, item: pystray.MenuItem) -> None:
        if self.root:
            self.root.after(0, self._shutdown)
        else:
            icon.stop()

    def _shutdown(self) -> None:
        logger.info("Shutting down Folder Searcher")
        try:
            self._icon.stop()
        except Exception:
            pass
        if self.root:
            self.root.quit()

    # ------------------------------------------------------------------
    # Window lifecycle
    # ------------------------------------------------------------------

    def _show_window(self) -> None:
        if not self._widgets_built:
            self._build_window()
            return
        if self.window and self.window.winfo_exists():
            self.window.deiconify()
            self.window.attributes("-topmost", True)
            self.window.lift()
            self.window.focus_force()

    def _build_window(self) -> None:
        assert self.root is not None
        win = tk.Toplevel(self.root)
        win.title("Folder Searcher")
        win.geometry("1200x500")
        win.resizable(True, True)
        win.attributes("-topmost", True)
        self.window = win

        self.setup_ui()

        # Close button → hide to tray, don't destroy
        win.protocol("WM_DELETE_WINDOW", win.withdraw)
        self._widgets_built = True
        win.lift()
        win.focus_force()
        if self.search_entry:
            self.search_entry.focus_set()

    # ------------------------------------------------------------------
    # Config
    # ------------------------------------------------------------------

    def save_config(self) -> None:
        """Persist the current roots and display settings."""
        if self.skip_depth_var is not None:
            try:
                self.config.skip_depth = max(0, int(self.skip_depth_var.get()))
            except (tk.TclError, ValueError):
                logger.warning("Invalid skip depth in the UI; keeping %d", self.config.skip_depth)
                self.skip_depth_var.set(self.config.skip_depth)
        if self.prune_var is not None:
            self.config.prune_email_branches = bool(self.prune_var.get())
        try:
            save_config(self.config_file, self.config)
            self.status_var.set("Configuration saved")
        except OSError as exc:
            logger.error("Error saving configuration: %s", exc)
            messagebox.showerror("Error", f"Error saving configuration: {exc}", parent=self.window)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def setup_ui(self) -> None:
        """Build the notebook and both tabs inside the Toplevel window."""
        parent = self.window
        assert parent is not None

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        container = ttk.Frame(parent, padding="10")
        container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(container)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        search_tab = ttk.Frame(self.notebook, padding="10")
        folders_tab = ttk.Frame(self.notebook, padding="10")
        # Search / Access first, so it is the tab selected on open.
        self.notebook.add(search_tab, text="Search / Access")
        self.notebook.add(folders_tab, text="Folders")
        self.notebook.select(search_tab)

        self._build_search_tab(search_tab)
        self._build_folders_tab(folders_tab)

        status_bar = ttk.Label(container, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

    def _build_search_tab(self, tab: ttk.Frame) -> None:
        """Search box, results list, and the double-click-to-open binding."""
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        input_frame = ttk.Frame(tab)
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(0, weight=1)

        ttk.Label(input_frame, text="Search word(s):").grid(row=0, column=0, sticky=tk.W)

        self.search_entry = ttk.Entry(input_frame, textvariable=self.search_var, width=40)
        self.search_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        self.search_entry.bind('<Return>', lambda event: self.search_folders())

        ttk.Button(input_frame, text="Search",
                   command=self.search_folders).grid(row=1, column=1, padx=(5, 0))

        ttk.Label(tab, text="Results (double-click to open):").grid(
            row=1, column=0, sticky=tk.W, pady=(10, 5))

        list_frame = ttk.Frame(tab)
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.results_listbox = tk.Listbox(list_frame, height=15)
        self.results_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.results_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_listbox.configure(yscrollcommand=scrollbar.set)

        self.results_listbox.bind('<Double-Button-1>', self.open_folder)

    def _build_folders_tab(self, tab: ttk.Frame) -> None:
        """Root management, display settings, and the scan trigger."""
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ttk.Label(tab, text="Root folders to scan:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        list_frame = ttk.Frame(tab)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.roots_listbox = tk.Listbox(list_frame, height=8)
        self.roots_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        roots_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.roots_listbox.yview)
        roots_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.roots_listbox.configure(yscrollcommand=roots_scrollbar.set)

        button_frame = ttk.Frame(tab)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)

        ttk.Button(button_frame, text="Add Root", command=self.add_root).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Remove Selected", command=self.remove_root).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Scan All Roots", command=self.scan_all_roots).grid(row=0, column=2, padx=5)

        options_frame = ttk.LabelFrame(tab, text="Display", padding="10")
        options_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(options_frame, text="Skip path depth:").grid(row=0, column=0, sticky=tk.W)
        ttk.Spinbox(options_frame, from_=0, to=20, width=5,
                    textvariable=self.skip_depth_var).grid(row=0, column=1, sticky=tk.W, padx=(5, 20))

        ttk.Checkbutton(options_frame, text="Prune email branches to the owning item",
                        variable=self.prune_var).grid(row=0, column=2, sticky=tk.W)

        ttk.Button(options_frame, text="Save Config",
                   command=self.save_config).grid(row=0, column=3, sticky=tk.W, padx=(20, 0))

        self.refresh_roots_listbox()

    def refresh_roots_listbox(self) -> None:
        """Redraw the roots list from the current config."""
        if self.roots_listbox is None:
            return
        self.roots_listbox.delete(0, tk.END)
        for root_path in self.config.root_paths:
            self.roots_listbox.insert(tk.END, root_path)

    # ------------------------------------------------------------------
    # Root management
    # ------------------------------------------------------------------

    def add_root(self) -> None:
        """Pick a folder and add it as a root."""
        folder = filedialog.askdirectory(title="Select Root Folder", parent=self.window)
        if not folder:
            return
        normalized = normalize_root(folder)
        if normalized in self.config.root_paths:
            self.status_var.set(f"Root already configured: {normalized}")
            return
        self.config.root_paths.append(normalized)
        self.refresh_roots_listbox()
        self.save_config()
        self.status_var.set(f"Added root: {normalized}. Run Scan All Roots to index it.")
        logger.info("Added root: %s", normalized)

    def remove_root(self) -> None:
        """Drop the selected root from the config."""
        if self.roots_listbox is None:
            return
        selection = self.roots_listbox.curselection()
        if not selection:
            self.status_var.set("Select a root to remove")
            return
        removed = self.config.root_paths.pop(selection[0])
        self.refresh_roots_listbox()
        self.save_config()
        self.status_var.set(f"Removed root: {removed}. Run Scan All Roots to rebuild the index.")
        logger.info("Removed root: %s", removed)

    # ------------------------------------------------------------------
    # Scan / load
    # ------------------------------------------------------------------

    def scan_all_roots(self) -> None:
        """Rebuild the index across every configured root."""
        if not self.config.root_paths:
            messagebox.showerror("Error", "Add at least one root folder first", parent=self.window)
            return

        try:
            self.status_var.set("Scanning folder structure...")
            self.root.update()

            count = self.index.scan(self.config.root_paths)
            self.index.save(self.config.structure_file)

            self.status_var.set(
                f"Scan complete. {count} folders across {len(self.index.roots)} root(s)")
            logger.info("Scan complete: %d folders", count)
        except OSError as exc:
            logger.error("Error scanning folder structure: %s", exc)
            messagebox.showerror("Error", f"Error scanning folder structure: {exc}", parent=self.window)
            self.status_var.set("Scan failed")

    def load_structure(self) -> None:
        """Load the saved index without rescanning.

        The first configured root doubles as the legacy root, so a
        pre-multi-root ``folder_structure.txt`` of relative paths still
        resolves to openable absolute paths.
        """
        legacy_root = self.config.root_paths[0] if self.config.root_paths else None
        try:
            count = self.index.load(self.config.structure_file, legacy_root=legacy_root)
        except OSError as exc:
            logger.error("Error loading structure: %s", exc)
            self.status_var.set("Error loading structure file")
            return

        if count:
            self.status_var.set(f"Index loaded. {count} folders available")
        else:
            self.status_var.set("No index found. Add roots on the Folders tab, then Scan All Roots.")

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def search_folders(self) -> None:
        """Search the loaded index and render the results."""
        search_input = self.search_var.get().strip()
        if not search_input:
            messagebox.showwarning("Warning", "Please enter a search term", parent=self.window)
            return

        if not len(self.index):
            messagebox.showwarning(
                "Warning",
                "No index loaded. Add roots on the Folders tab, then Scan All Roots.",
                parent=self.window,
            )
            return

        self.status_var.set("Searching...")
        self.root.update()

        self.results = search(
            self.index,
            search_input,
            skip_depth=self.config.skip_depth,
            prune_email_branches=self.config.prune_email_branches,
        )

        self.results_listbox.delete(0, tk.END)
        for result in self.results:
            self.results_listbox.insert(tk.END, result.display_path)

        scope = "pruning on" if self.config.prune_email_branches else "pruning off"
        self.status_var.set(f"Found {len(self.results)} matching folders ({scope})")
        logger.info("Search for '%s' returned %d results (%s)", search_input, len(self.results), scope)

    def open_folder(self, event) -> None:
        """Open the selected result in Windows Explorer."""
        selection = self.results_listbox.curselection()
        if not selection:
            return

        result = self.results[selection[0]]
        # Results are stored absolute, so display trimming never affects this.
        full_path = os.path.normpath(result.absolute_path)
        logger.info("Opening folder: %s", full_path)

        if not os.path.exists(full_path):
            messagebox.showerror("Error", f"Folder does not exist: {full_path}", parent=self.window)
            return

        try:
            os.startfile(full_path)
        except OSError as exc:
            logger.error("Error opening folder: %s", exc)
            try:
                subprocess.run(
                    ['explorer', full_path],
                    check=True,
                    shell=True,
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                )
                logger.info("Opened folder using subprocess: %s", full_path)
            except (OSError, subprocess.SubprocessError) as exc2:
                logger.error("Error opening folder with subprocess: %s", exc2)
                messagebox.showerror("Error", f"Error opening folder: {exc2}", parent=self.window)

    # ------------------------------------------------------------------
    # Single-instance lock and pystray setup
    # ------------------------------------------------------------------

    def _acquire_lock(self) -> bool:
        """Take a process-wide named mutex. False if another instance holds it.

        The mutex is released automatically when the process exits, so no
        explicit cleanup is needed.
        """
        k = ctypes.windll.kernel32
        k.CreateMutexW.restype = wintypes.HANDLE
        k.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        self._mutex = k.CreateMutexW(None, True, self._MUTEX_NAME)
        return k.GetLastError() != self._ERROR_ALREADY_EXISTS

    def _setup_pystray(self, icon: pystray.Icon) -> None:
        """Run on pystray's worker thread once the icon is visible.

        We create the Tk root here so tkinter and the icon menu callbacks
        share the same thread. Menu callbacks marshal work back via
        root.after() to keep things on this thread.
        """
        icon.visible = True

        if not self._acquire_lock():
            icon.notify("Folder Searcher is already running.", "Already Running")
            icon.stop()
            return

        self.root = tk.Tk()
        self.root.withdraw()

        # Now that a Tk root exists, build tk variables
        self.search_var = tk.StringVar()
        self.skip_depth_var = tk.IntVar(value=self.config.skip_depth)
        self.prune_var = tk.BooleanVar(value=self.config.prune_email_branches)
        self.status_var = tk.StringVar(value="Ready")

        # Load the persisted index up front so Search / Access works on open
        self.load_structure()

        self.root.mainloop()
        # mainloop exits → make sure the icon stops too
        try:
            icon.stop()
        except Exception:
            pass

    def run(self) -> None:
        """Start the tray application. Blocks until Quit."""
        logger.info("Starting Folder Searcher (tray)")
        self._icon.run(setup=self._setup_pystray)


def main() -> None:
    """Main entry point for the application."""
    app = FolderSearcher()
    app.run()


if __name__ == "__main__":
    main()
