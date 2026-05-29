#!/usr/bin/env python3
"""
Folder Searcher - A simple tool to search for folders containing specific words.

This application allows users to:
1. Scan a folder structure and store it in memory
2. Search for folders containing specific words
3. Open found folders directly in Windows Explorer
4. Persist folder structure data between sessions
5. Search in the current active Explorer window path

The app runs from the system tray. Click the tray icon to open the window;
closing the window minimizes it back to the tray. Quit from the tray menu.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import logging
import ctypes
from ctypes import wintypes
from pathlib import Path
import subprocess
from typing import Dict, List, Optional
import win32gui
import win32process
import win32api
import win32con
import psutil
import pystray
from PIL import Image, ImageDraw

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
        # Application state
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(script_dir, "foldersearcher.json")
        self.structure_file = os.path.join(script_dir, "folder_structure.txt")
        self.folder_structure: Dict[str, List[str]] = {}
        self.root_folder = ""
        self.search_in_explorer = True  # Default to search in Explorer window path

        # Tk objects — created in the pystray setup callback
        self.root: Optional[tk.Tk] = None
        self.window: Optional[tk.Toplevel] = None
        self._widgets_built = False

        # Tk variables — created once the Tk root exists
        self.folder_var: Optional[tk.StringVar] = None
        self.search_var: Optional[tk.StringVar] = None
        self.search_in_explorer_var: Optional[tk.BooleanVar] = None
        self.status_var: Optional[tk.StringVar] = None

        # Widgets populated in setup_ui
        self.folder_entry: Optional[ttk.Entry] = None
        self.search_entry: Optional[ttk.Entry] = None
        self.scope_checkbox: Optional[ttk.Checkbutton] = None
        self.results_listbox: Optional[tk.Listbox] = None

        # Single-instance mutex handle — kept for the process lifetime
        self._mutex: Optional[int] = None

        # Load configuration (no Tk required)
        self.load_config()

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

    # ------------------------------------------------------------------
    # Config
    # ------------------------------------------------------------------

    def load_config(self):
        """Load configuration from JSON file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.root_folder = config.get('root_folder', '')
                    # Keep the structure file in the same directory as the script
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    self.structure_file = os.path.join(script_dir, "folder_structure.txt")
                    logger.info(f"Configuration loaded from {self.config_file}")
            else:
                # Create default configuration
                self.create_default_config()
        except Exception as e:
            logger.error(f"Error loading configuration: {e}")
            self.create_default_config()

    def create_default_config(self):
        """Create default configuration file."""
        config = {
            'root_folder': '',
            'structure_file': 'folder_structure.txt',
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            logger.info("Default configuration file created")
        except Exception as e:
            logger.error(f"Error creating default config: {e}")

    def save_config(self):
        """Save current configuration to JSON file."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config = {
            'root_folder': self.root_folder,
            'structure_file': os.path.join(script_dir, "folder_structure.txt"),
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            logger.info("Configuration saved")
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")

    # ------------------------------------------------------------------
    # Explorer detection
    # ------------------------------------------------------------------

    def get_active_explorer_path(self) -> Optional[str]:
        """
        Get the path of the most recently active Windows Explorer window.

        Returns:
            Optional[str]: The path of the active Explorer window, or None if not found
        """
        try:
            # Get all explorer.exe processes
            explorer_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['name'] and proc.info['name'].lower() == 'explorer.exe':
                        explorer_processes.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if not explorer_processes:
                logger.warning("No Explorer processes found")
                return None

            # Collect all Explorer windows and their paths
            all_explorer_windows = []

            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        # Check if this window belongs to any Explorer process
                        if any(proc.pid == pid for proc in explorer_processes):
                            title = win32gui.GetWindowText(hwnd)
                            if title:
                                # Try different patterns for extracting path from title
                                potential_path = None

                                # Pattern 1: "Path - File Explorer" (most common)
                                if title.endswith(" - File Explorer"):
                                    potential_path = title[:-15].strip()  # Remove " - File Explorer" and trim whitespace
                                    if os.path.isdir(potential_path):
                                        pass
                                    else:
                                        parent_dir = os.path.dirname(potential_path)
                                        if os.path.isdir(parent_dir):
                                            potential_path = parent_dir

                                # Pattern 2: "Folder Name - C:\path\to\folder"
                                elif " - " in title:
                                    title_parts = title.split(" - ")
                                    if len(title_parts) > 1:
                                        potential_path = title_parts[-1]

                                # Pattern 3: Just the path itself
                                elif os.path.exists(title):
                                    potential_path = title

                                # Pattern 4: Check if title contains a drive letter
                                elif any(drive in title for drive in ['C:', 'D:', 'E:', 'F:', 'G:', 'H:', 'I:', 'J:', 'K:', 'L:', 'M:', 'N:', 'O:', 'P:', 'Q:', 'R:', 'S:', 'T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:']):
                                    words = title.split()
                                    for word in words:
                                        if any(drive in word for drive in ['C:', 'D:', 'E:', 'F:', 'G:', 'H:', 'I:', 'J:', 'K:', 'L:', 'M:', 'N:', 'O:', 'P:', 'Q:', 'R:', 'S:', 'T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:']):
                                            if os.path.exists(word):
                                                potential_path = word
                                                break

                                if potential_path and os.path.exists(potential_path):
                                    windows.append({
                                        'hwnd': hwnd,
                                        'title': title,
                                        'path': potential_path,
                                        'pid': pid
                                    })
                                    logger.debug(f"Found Explorer window: '{title}' -> '{potential_path}'")
                    except Exception as e:
                        logger.debug(f"Error processing window {hwnd}: {e}")
                return True

            win32gui.EnumWindows(enum_windows_callback, all_explorer_windows)

            if not all_explorer_windows:
                logger.warning("No Explorer windows with valid paths found")
                return None

            # Try to find the most recently active window
            active_window = win32gui.GetForegroundWindow()

            try:
                _, _active_pid = win32process.GetWindowThreadProcessId(active_window)
                for window in all_explorer_windows:
                    if window['hwnd'] == active_window:
                        logger.debug(f"Found active Explorer window: {window['path']}")
                        return window['path']
            except Exception as e:
                logger.warning(f"Error getting active window info: {e}")

            if all_explorer_windows:
                first_window = all_explorer_windows[0]
                logger.debug(f"Using first Explorer window: {first_window['path']}")
                return first_window['path']

            logger.warning("No active Explorer window with valid path found")
            return None

        except Exception as e:
            logger.error(f"Error getting active Explorer path: {e}")
            return None

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def setup_ui(self):
        """Build the widgets inside the Toplevel window."""
        parent = self.window
        assert parent is not None

        # Main frame
        main_frame = ttk.Frame(parent, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # Folder selection section
        ttk.Label(main_frame, text="Root Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)

        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        folder_frame.columnconfigure(0, weight=1)

        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        self.folder_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).grid(row=0, column=1)

        # Scan button
        ttk.Button(main_frame, text="Scan Folder Structure",
                   command=self.scan_folder_structure).grid(row=2, column=0, columnspan=2, pady=10)

        # Search section
        search_frame = ttk.LabelFrame(main_frame, text="Search Folders", padding="10")
        search_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        search_frame.columnconfigure(0, weight=1)
        search_frame.rowconfigure(1, weight=1)

        # Search input
        search_input_frame = ttk.Frame(search_frame)
        search_input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        search_input_frame.columnconfigure(0, weight=1)

        ttk.Label(search_input_frame, text="Search word:").grid(row=0, column=0, sticky=tk.W)

        self.search_entry = ttk.Entry(search_input_frame, textvariable=self.search_var, width=40)
        self.search_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)

        # Search scope option
        scope_frame = ttk.Frame(search_input_frame)
        scope_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.scope_checkbox = ttk.Checkbutton(
            scope_frame,
            text="Search in current Explorer window path",
            variable=self.search_in_explorer_var,
            command=self.update_search_scope
        )
        self.scope_checkbox.grid(row=0, column=0, sticky=tk.W)

        # Bind Enter key to search function
        self.search_entry.bind('<Return>', lambda event: self.search_folders())

        ttk.Button(search_input_frame, text="Search",
                   command=self.search_folders).grid(row=1, column=1, padx=(5, 0))

        # Results section
        ttk.Label(search_frame, text="Results:").grid(row=1, column=0, sticky=tk.W, pady=(10, 5))

        # Results listbox with scrollbar
        list_frame = ttk.Frame(search_frame)
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.results_listbox = tk.Listbox(list_frame, height=15)
        self.results_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.results_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_listbox.configure(yscrollcommand=scrollbar.set)

        # Bind double-click event
        self.results_listbox.bind('<Double-Button-1>', self.open_folder)

        # Status bar
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

    def browse_folder(self):
        """Open folder browser dialog."""
        folder = filedialog.askdirectory(title="Select Root Folder", parent=self.window)
        if folder:
            self.folder_var.set(folder)
            self.root_folder = folder
            self.save_config()
            logger.info(f"Selected folder: {folder}")

    def update_search_scope(self):
        """Update the search scope based on checkbox state."""
        self.search_in_explorer = self.search_in_explorer_var.get()
        if self.search_in_explorer:
            explorer_path = self.get_active_explorer_path()
            if explorer_path:
                logger.debug(f"Explorer path detected: {explorer_path}")
                try:
                    if os.path.exists(explorer_path):
                        contents = os.listdir(explorer_path)
                        logger.debug(f"Contents of {explorer_path}: {contents[:10]}...")
                    else:
                        logger.warning(f"Explorer path does not exist: {explorer_path}")
                except Exception as e:
                    logger.error(f"Error listing directory contents: {e}")

                self.status_var.set(f"Search scope: Current Explorer path ({explorer_path})")
            else:
                self.status_var.set("Search scope: Current Explorer path (not found)")
                self.search_in_explorer_var.set(False)
                self.search_in_explorer = False
                self.debug_explorer_windows()
        else:
            self.status_var.set("Search scope: Full scanned structure")
        logger.debug(f"Search scope updated: {'Explorer path' if self.search_in_explorer else 'Full structure'}")

    def debug_explorer_windows(self):
        """Debug method to show all Explorer windows and their titles."""
        try:
            logger.debug("=== DEBUG: Explorer Windows Detection ===")

            explorer_processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if proc.info['name'] and proc.info['name'].lower() == 'explorer.exe':
                        explorer_processes.append(proc)
                        logger.debug(f"Found Explorer process: PID {proc.pid}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if not explorer_processes:
                logger.warning("No Explorer processes found")
                return

            all_windows = []

            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        title = win32gui.GetWindowText(hwnd)
                        if title and any(proc.pid == pid for proc in explorer_processes):
                            windows.append({
                                'hwnd': hwnd,
                                'title': title,
                                'pid': pid
                            })
                            logger.debug(f"Explorer window: '{title}' (PID: {pid})")
                    except Exception as e:
                        logger.debug(f"Error processing window {hwnd}: {e}")
                return True

            win32gui.EnumWindows(enum_windows_callback, all_windows)

            logger.debug(f"Total Explorer windows found: {len(all_windows)}")
            logger.debug("=== END DEBUG ===")

        except Exception as e:
            logger.error(f"Error in debug_explorer_windows: {e}")

    # ------------------------------------------------------------------
    # Scan / load / save folder structure
    # ------------------------------------------------------------------

    def scan_folder_structure(self):
        """Scan the folder structure and save it to file."""
        root_folder = self.folder_var.get()
        if not root_folder or not os.path.exists(root_folder):
            messagebox.showerror("Error", "Please select a valid folder", parent=self.window)
            return

        try:
            self.status_var.set("Scanning folder structure...")
            self.root.update()

            self.folder_structure = {}

            for root, dirs, files in os.walk(root_folder):
                rel_path = os.path.relpath(root, root_folder)
                if rel_path == '.':
                    rel_path = ''

                folder_key = os.path.normpath(rel_path) if rel_path else 'root'

                if folder_key not in self.folder_structure:
                    self.folder_structure[folder_key] = []

                    for dir_name in dirs:
                        if folder_key == 'root':
                            full_subdir_path = dir_name
                        else:
                            full_subdir_path = os.path.join(folder_key, dir_name)
                        self.folder_structure[folder_key].append(full_subdir_path)

            self.save_structure()

            self.status_var.set(f"Scan complete. Found {len(self.folder_structure)} folders")
            logger.info(f"Folder structure scanned and saved. Total folders: {len(self.folder_structure)}")

        except Exception as e:
            logger.error(f"Error scanning folder structure: {e}")
            messagebox.showerror("Error", f"Error scanning folder structure: {e}", parent=self.window)
            self.status_var.set("Scan failed")

    def save_structure(self):
        """Save folder structure to text file."""
        try:
            with open(self.structure_file, 'w', encoding='utf-8') as f:
                for folder_path, subdirs in sorted(self.folder_structure.items()):
                    f.write(f"{folder_path}\n")
                    for subdir in sorted(subdirs):
                        f.write(f"  {subdir}\n")
                    f.write("\n")
            logger.info(f"Folder structure saved to {self.structure_file}")
        except Exception as e:
            logger.error(f"Error saving structure: {e}")

    def load_structure(self):
        """Load folder structure from file."""
        if not os.path.exists(self.structure_file):
            self.status_var.set("No structure file found. Please scan a folder.")
            return

        try:
            self.folder_structure = {}
            current_folder = ""

            with open(self.structure_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('  '):
                        current_folder = line
                        current_folder = os.path.normpath(current_folder)
                        if current_folder not in self.folder_structure:
                            self.folder_structure[current_folder] = []
                    elif line.startswith('  ') and current_folder:
                        subdir = line[2:]
                        subdir = os.path.normpath(subdir)
                        if subdir not in self.folder_structure[current_folder]:
                            self.folder_structure[current_folder].append(subdir)

            self.status_var.set(f"Structure loaded. {len(self.folder_structure)} folders available")
            logger.info(f"Folder structure loaded from {self.structure_file}")
            sample_entries = list(self.folder_structure.items())[:3]
            for folder_path, subdirs in sample_entries:
                logger.debug(f"Folder: '{folder_path}' -> Subdirs: {subdirs[:3]}")

        except Exception as e:
            logger.error(f"Error loading structure: {e}")
            self.status_var.set("Error loading structure file")

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def search_folders(self):
        """Search for folders containing the specified word(s)."""
        search_input = self.search_var.get().strip()
        if not search_input:
            messagebox.showwarning("Warning", "Please enter a search term", parent=self.window)
            return

        search_input_processed = search_input.replace(';', ' ')
        search_terms = [term.strip().lower() for term in search_input_processed.split() if term.strip()]
        if not search_terms:
            messagebox.showwarning("Warning", "Please enter valid search term(s)", parent=self.window)
            return

        try:
            self.status_var.set("Searching...")
            self.root.update()

            self.results_listbox.delete(0, tk.END)

            matching_folders: set = set()

            self.search_in_explorer = self.search_in_explorer_var.get()

            if self.search_in_explorer:
                explorer_path = self.get_active_explorer_path()
                if not explorer_path:
                    messagebox.showwarning(
                        "Warning",
                        "No active Explorer window found. Searching in full structure instead.",
                        parent=self.window,
                    )
                    self.search_in_explorer_var.set(False)
                    self.search_in_explorer = False
                    self._search_in_full_structure(search_terms, matching_folders)
                else:
                    self._search_in_explorer_path(search_terms, explorer_path, matching_folders)
            else:
                if not self.folder_structure:
                    messagebox.showwarning(
                        "Warning",
                        "No folder structure loaded. Please scan a folder first.",
                        parent=self.window,
                    )
                    return
                self._search_in_full_structure(search_terms, matching_folders)

            matching_folders_sorted = sorted(list(matching_folders))

            for folder in matching_folders_sorted:
                self.results_listbox.insert(tk.END, folder)

            scope_text = "Explorer path" if self.search_in_explorer else "full structure"
            self.status_var.set(f"Found {len(matching_folders_sorted)} matching folders in {scope_text}")
            logger.info(f"Search completed. Found {len(matching_folders_sorted)} folders containing '{search_input}' in {scope_text}")

        except Exception as e:
            logger.error(f"Error searching folders: {e}")
            messagebox.showerror("Error", f"Error searching folders: {e}", parent=self.window)
            self.status_var.set("Search failed")

    def _search_in_full_structure(self, search_terms: List[str], matching_folders: set):
        """Search in the full scanned folder structure."""
        all_relative_paths = set()
        for folder_path, subdirs in self.folder_structure.items():
            if folder_path != 'root':
                all_relative_paths.add(folder_path)
            for subdir in subdirs:
                all_relative_paths.add(subdir)

        for rel_path in all_relative_paths:
            full_path = os.path.join(self.root_folder, rel_path)
            if all(term in full_path.lower() for term in search_terms):
                matching_folders.add(rel_path)
                logger.debug(f"Found match: '{rel_path}'")

    def _search_in_explorer_path(self, search_terms: List[str], explorer_path: str, matching_folders: set):
        """Search only in the current Explorer window path."""
        try:
            logger.debug(f"Starting search in Explorer path: {explorer_path}")
            logger.debug(f"Search terms: {search_terms}")

            for root, dirs, files in os.walk(explorer_path):
                if all(term in root.lower() for term in search_terms):
                    matching_folders.add(root)
                    logger.debug(f"Found matching directory by full path: '{root}'")

                for dir_name in dirs:
                    full_path = os.path.join(root, dir_name)
                    if all(term in full_path.lower() for term in search_terms):
                        matching_folders.add(full_path)
                        logger.debug(f"Found matching subdirectory by full path: '{full_path}'")

            logger.debug(f"Search completed in Explorer path: {explorer_path}")

        except Exception as e:
            logger.error(f"Error searching in Explorer path: {e}")
            raise

    def open_folder(self, event):
        """Open the selected folder in Windows Explorer."""
        selection = self.results_listbox.curselection()
        if not selection:
            return

        folder_path = self.results_listbox.get(selection[0])

        logger.info(f"Selected folder path: '{folder_path}'")

        if os.path.isabs(folder_path):
            full_path = folder_path
        else:
            logger.info(f"Root folder: '{self.root_folder}'")
            if folder_path == 'root':
                full_path = self.root_folder
            else:
                full_path = os.path.join(self.root_folder, folder_path)

        full_path = os.path.normpath(full_path)
        logger.info(f"Constructed full path: '{full_path}'")

        if not os.path.exists(full_path):
            messagebox.showerror("Error", f"Folder does not exist: {full_path}", parent=self.window)
            return

        try:
            os.startfile(full_path)
            logger.info(f"Opened folder: {full_path}")
        except Exception as e:
            logger.error(f"Error opening folder: {e}")
            try:
                subprocess.run(['explorer', full_path], check=True, shell=True)
                logger.info(f"Opened folder using subprocess: {full_path}")
            except Exception as e2:
                logger.error(f"Error opening folder with subprocess: {e2}")
                messagebox.showerror("Error", f"Error opening folder: {e2}", parent=self.window)

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
        self.folder_var = tk.StringVar(value=self.root_folder)
        self.search_var = tk.StringVar()
        self.search_in_explorer_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready")

        # Load persisted structure (uses status_var)
        self.load_structure()

        self.root.mainloop()
        # mainloop exits → make sure the icon stops too
        try:
            icon.stop()
        except Exception:
            pass

    def run(self):
        """Start the tray application. Blocks until Quit."""
        logger.info("Starting Folder Searcher (tray)")
        self._icon.run(setup=self._setup_pystray)


def main():
    """Main entry point for the application."""
    app = FolderSearcher()
    app.run()


if __name__ == "__main__":
    main()
