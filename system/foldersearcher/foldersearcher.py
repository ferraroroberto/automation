#!/usr/bin/env python3
"""
Folder Searcher - A simple tool to search for folders containing specific words.

This application allows users to:
1. Scan a folder structure and store it in memory
2. Search for folders containing specific words
3. Open found folders directly in Windows Explorer
4. Persist folder structure data between sessions
5. Search in the current active Explorer window path

Author: Assistant
Date: 2024
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import logging
from pathlib import Path
import subprocess
from typing import Dict, List, Optional
import win32gui
import win32process
import win32api
import win32con
import psutil

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
    
    This class handles:
    - Loading configuration from JSON
    - Scanning folder structures
    - Searching for folders containing specific words
    - Opening folders in Windows Explorer
    - Persisting data between sessions
    """
    
    def __init__(self):
        """Initialize the FolderSearcher application."""
        self.root = tk.Tk()
        self.root.title("Folder Searcher")
        self.root.geometry("1200x500")
        self.root.resizable(True, True)
        
        # Application state
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(script_dir, "foldersearcher.json")
        self.structure_file = os.path.join(script_dir, "folder_structure.txt")
        self.folder_structure: Dict[str, List[str]] = {}
        self.root_folder = ""
        self.search_in_explorer = True  # Default to search in Explorer window path
        self.always_on_top = False  # Default to not always on top
        
        # Load configuration and setup UI
        self.load_config()
        self.setup_ui()
        self.load_structure()
        
        logger.info("Folder Searcher application initialized")
    
    def load_config(self):
        """Load configuration from JSON file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.root_folder = config.get('root_folder', '')
                    self.always_on_top = config.get('always_on_top', False)
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
            'always_on_top': False
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            logger.info("Default configuration file created")
        except Exception as e:
            logger.error(f"Error creating default config: {e}")
    
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
                                    # If the path ends with a folder name, we need to get the parent directory
                                    # because the title shows the current folder, but we want to search in its parent
                                    if os.path.isdir(potential_path):
                                        # This is already a directory, use it as is
                                        pass
                                    else:
                                        # Try to get the parent directory
                                        parent_dir = os.path.dirname(potential_path)
                                        if os.path.isdir(parent_dir):
                                            potential_path = parent_dir
                                
                                # Pattern 2: "Folder Name - C:\path\to\folder"
                                elif " - " in title:
                                    title_parts = title.split(" - ")
                                    if len(title_parts) > 1:
                                        potential_path = title_parts[-1]
                                
                                # Pattern 3: Just the path itself (for some Explorer windows)
                                elif os.path.exists(title):
                                    potential_path = title
                                
                                # Pattern 4: Check if title contains a drive letter
                                elif any(drive in title for drive in ['C:', 'D:', 'E:', 'F:', 'G:', 'H:', 'I:', 'J:', 'K:', 'L:', 'M:', 'N:', 'O:', 'P:', 'Q:', 'R:', 'S:', 'T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:']):
                                    # Try to extract path from title
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
            active_pid = None
            
            try:
                _, active_pid = win32process.GetWindowThreadProcessId(active_window)
                # Look for the active window in our list
                for window in all_explorer_windows:
                    if window['hwnd'] == active_window:
                        logger.debug(f"Found active Explorer window: {window['path']}")
                        return window['path']
            except Exception as e:
                logger.warning(f"Error getting active window info: {e}")
            
            # If no active window found, return the first valid Explorer window
            if all_explorer_windows:
                first_window = all_explorer_windows[0]
                logger.debug(f"Using first Explorer window: {first_window['path']}")
                return first_window['path']
            
            logger.warning("No active Explorer window with valid path found")
            return None
            
        except Exception as e:
            logger.error(f"Error getting active Explorer path: {e}")
            return None
    
    def setup_ui(self):
        """Setup the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Folder selection section
        ttk.Label(main_frame, text="Root Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        folder_frame.columnconfigure(0, weight=1)
        
        self.folder_var = tk.StringVar(value=self.root_folder)
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
        
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_input_frame, textvariable=self.search_var, width=40)
        self.search_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Search scope option
        scope_frame = ttk.Frame(search_input_frame)
        scope_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.search_in_explorer_var = tk.BooleanVar(value=False)  # Default to unchecked
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
        
        # Always on top toggle
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.always_on_top_var = tk.BooleanVar(value=self.always_on_top)
        self.always_on_top_checkbox = ttk.Checkbutton(
            top_frame,
            text="Always on top",
            variable=self.always_on_top_var,
            command=self.toggle_always_on_top
        )
        self.always_on_top_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        # Apply initial always on top state
        if self.always_on_top:
            self.root.attributes('-topmost', True)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def toggle_always_on_top(self):
        """Toggle the always on top state of the window."""
        if self.always_on_top_var.get():
            self.root.attributes('-topmost', True)
            logger.info("Window set to always on top")
        else:
            self.root.attributes('-topmost', False)
            logger.info("Window removed from always on top")
        
        # Save the configuration
        self.save_config()
    
    def browse_folder(self):
        """Open folder browser dialog."""
        folder = filedialog.askdirectory(title="Select Root Folder")
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
                # List the contents of the directory to verify
                try:
                    if os.path.exists(explorer_path):
                        contents = os.listdir(explorer_path)
                        logger.debug(f"Contents of {explorer_path}: {contents[:10]}...")  # Show first 10 items
                    else:
                        logger.warning(f"Explorer path does not exist: {explorer_path}")
                except Exception as e:
                    logger.error(f"Error listing directory contents: {e}")
                
                self.status_var.set(f"Search scope: Current Explorer path ({explorer_path})")
            else:
                self.status_var.set("Search scope: Current Explorer path (not found)")
                # Uncheck the box if no Explorer window found
                self.search_in_explorer_var.set(False)
                self.search_in_explorer = False
                # Show debug info
                self.debug_explorer_windows()
        else:
            self.status_var.set("Search scope: Full scanned structure")
        logger.debug(f"Search scope updated: {'Explorer path' if self.search_in_explorer else 'Full structure'}")
    
    def debug_explorer_windows(self):
        """Debug method to show all Explorer windows and their titles."""
        try:
            logger.debug("=== DEBUG: Explorer Windows Detection ===")
            
            # Get all explorer.exe processes
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
            
            # Collect all windows
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
    
    def save_config(self):
        """Save current configuration to JSON file."""
        # Get the script directory for consistent path handling
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config = {
            'root_folder': self.root_folder,
            'structure_file': os.path.join(script_dir, "folder_structure.txt"),
            'always_on_top': self.always_on_top_var.get()
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2)
            logger.info("Configuration saved")
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")
    
    def scan_folder_structure(self):
        """Scan the folder structure and save it to file."""
        root_folder = self.folder_var.get()
        if not root_folder or not os.path.exists(root_folder):
            messagebox.showerror("Error", "Please select a valid folder")
            return
        
        try:
            self.status_var.set("Scanning folder structure...")
            self.root.update()
            
            # Clear existing structure
            self.folder_structure = {}
            
            # Scan all folders recursively
            for root, dirs, files in os.walk(root_folder):
                # Get relative path from root folder
                rel_path = os.path.relpath(root, root_folder)
                if rel_path == '.':
                    rel_path = ''
                
                # Normalize the path to ensure consistency
                folder_key = os.path.normpath(rel_path) if rel_path else 'root'
                
                # Only add if not already present (avoid duplicates)
                if folder_key not in self.folder_structure:
                    self.folder_structure[folder_key] = []
                    
                    # Add subdirectories as full paths
                    for dir_name in dirs:
                        if folder_key == 'root':
                            full_subdir_path = dir_name
                        else:
                            full_subdir_path = os.path.join(folder_key, dir_name)
                        self.folder_structure[folder_key].append(full_subdir_path)
            
            # Save to file
            self.save_structure()
            
            self.status_var.set(f"Scan complete. Found {len(self.folder_structure)} folders")
            logger.info(f"Folder structure scanned and saved. Total folders: {len(self.folder_structure)}")
            
        except Exception as e:
            logger.error(f"Error scanning folder structure: {e}")
            messagebox.showerror("Error", f"Error scanning folder structure: {e}")
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
                        # Normalize the folder path to use proper separators
                        current_folder = os.path.normpath(current_folder)
                        # Only add if not already present (avoid duplicates)
                        if current_folder not in self.folder_structure:
                            self.folder_structure[current_folder] = []
                    elif line.startswith('  ') and current_folder:
                        subdir = line[2:]  # Remove leading spaces
                        # Normalize the subdir path as well
                        subdir = os.path.normpath(subdir)
                        # Only add subdir if not already present in this folder
                        if subdir not in self.folder_structure[current_folder]:
                            self.folder_structure[current_folder].append(subdir)
            
            self.status_var.set(f"Structure loaded. {len(self.folder_structure)} folders available")
            logger.info(f"Folder structure loaded from {self.structure_file}")
            # Debug: log some sample entries
            sample_entries = list(self.folder_structure.items())[:3]
            for folder_path, subdirs in sample_entries:
                logger.debug(f"Folder: '{folder_path}' -> Subdirs: {subdirs[:3]}")
            
        except Exception as e:
            logger.error(f"Error loading structure: {e}")
            self.status_var.set("Error loading structure file")
    
    def search_folders(self):
        """Search for folders containing the specified word(s)."""
        search_input = self.search_var.get().strip()
        if not search_input:
            messagebox.showwarning("Warning", "Please enter a search term")
            return
        
        # Parse multiple search terms separated by space or semicolon
        search_input_processed = search_input.replace(';', ' ')
        search_terms = [term.strip().lower() for term in search_input_processed.split() if term.strip()]
        if not search_terms:
            messagebox.showwarning("Warning", "Please enter valid search term(s)")
            return
        
        try:
            self.status_var.set("Searching...")
            self.root.update()
            
            # Clear previous results
            self.results_listbox.delete(0, tk.END)
            
            matching_folders = set()  # Use set to eliminate duplicates
            
            # Update search scope based on current checkbox state
            self.search_in_explorer = self.search_in_explorer_var.get()
            
            if self.search_in_explorer:
                # Search in current Explorer window path (reload path every time)
                explorer_path = self.get_active_explorer_path()
                if not explorer_path:
                    messagebox.showwarning("Warning", "No active Explorer window found. Searching in full structure instead.")
                    self.search_in_explorer_var.set(False)
                    self.search_in_explorer = False
                    # Fall back to full structure search
                    self._search_in_full_structure(search_terms, matching_folders)
                else:
                    # Search only in the Explorer path
                    self._search_in_explorer_path(search_terms, explorer_path, matching_folders)
            else:
                # Search in full scanned structure
                if not self.folder_structure:
                    messagebox.showwarning("Warning", "No folder structure loaded. Please scan a folder first.")
                    return
                self._search_in_full_structure(search_terms, matching_folders)
            
            # Convert set to sorted list
            matching_folders = sorted(list(matching_folders))
            
            # Display results
            for folder in matching_folders:
                self.results_listbox.insert(tk.END, folder)
            
            scope_text = "Explorer path" if self.search_in_explorer else "full structure"
            self.status_var.set(f"Found {len(matching_folders)} matching folders in {scope_text}")
            logger.info(f"Search completed. Found {len(matching_folders)} folders containing '{search_input}' in {scope_text}")
            
        except Exception as e:
            logger.error(f"Error searching folders: {e}")
            messagebox.showerror("Error", f"Error searching folders: {e}")
            self.status_var.set("Search failed")
    
    def _search_in_full_structure(self, search_terms: List[str], matching_folders: set):
        """Search in the full scanned folder structure."""
        # Create a single list of all relative paths to check
        all_relative_paths = set()
        for folder_path, subdirs in self.folder_structure.items():
            if folder_path != 'root':
                all_relative_paths.add(folder_path)
            for subdir in subdirs:
                all_relative_paths.add(subdir)

        # Iterate through the unique relative paths
        for rel_path in all_relative_paths:
            # Construct the full, absolute path for checking
            full_path = os.path.join(self.root_folder, rel_path)
            
            # Check if the full path contains all search terms (AND condition)
            if all(term in full_path.lower() for term in search_terms):
                matching_folders.add(rel_path) # Add the relative path to results
                logger.debug(f"Found match: '{rel_path}'")
    
    def _search_in_explorer_path(self, search_terms: List[str], explorer_path: str, matching_folders: set):
        """Search only in the current Explorer window path."""
        try:
            logger.debug(f"Starting search in Explorer path: {explorer_path}")
            logger.debug(f"Search terms: {search_terms}")
            
            # Walk through the Explorer path directory
            for root, dirs, files in os.walk(explorer_path):
                # Check if the full root path contains all search terms
                if all(term in root.lower() for term in search_terms):
                    matching_folders.add(root)
                    logger.debug(f"Found matching directory by full path: '{root}'")
                
                # Check subdirectories against their full path
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
        
        # Determine if this is a full path (from Explorer search) or relative path (from full structure search)
        if os.path.isabs(folder_path):
            # This is already a full path (from Explorer search)
            full_path = folder_path
        else:
            # This is a relative path (from full structure search)
            logger.info(f"Root folder: '{self.root_folder}'")
            if folder_path == 'root':
                full_path = self.root_folder
            else:
                full_path = os.path.join(self.root_folder, folder_path)
        
        # Normalize the path to use proper Windows path separators
        full_path = os.path.normpath(full_path)
        logger.info(f"Constructed full path: '{full_path}'")
        
        # Check if the folder exists
        if not os.path.exists(full_path):
            messagebox.showerror("Error", f"Folder does not exist: {full_path}")
            return
        
        try:
            # Use os.startfile for better Windows compatibility
            os.startfile(full_path)
            logger.info(f"Opened folder: {full_path}")
        except Exception as e:
            logger.error(f"Error opening folder: {e}")
            # Fallback to subprocess if os.startfile fails
            try:
                subprocess.run(['explorer', full_path], check=True, shell=True)
                logger.info(f"Opened folder using subprocess: {full_path}")
            except Exception as e2:
                logger.error(f"Error opening folder with subprocess: {e2}")
                messagebox.showerror("Error", f"Error opening folder: {e2}")
    
    def run(self):
        """Start the application."""
        logger.info("Starting Folder Searcher application")
        self.root.mainloop()


def main():
    """Main entry point for the application."""
    app = FolderSearcher()
    app.run()


if __name__ == "__main__":
    main()