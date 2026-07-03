import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import logging
from pathlib import Path
import threading
from datetime import datetime
import io
import ctypes
import sys
import platform
import subprocess                 # NEW
from typing import Any, List

# --- NEW: Analyzer class for logic ---
class TreeSizeAnalyzer:
    def __init__(self) -> None:
        self.folder_sizes = {}      # Cache for folder sizes
        self.folder_disk_sizes = {} # Cache for folder disk sizes
        self.file_cache = {}        # Cache for file listings

    def get_folder_size_deep(self, path: str) -> int:
        if path in self.folder_sizes:
            return self.folder_sizes[path]
        total_size = 0
        file_count = 0
        error_count = 0
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                dirnames[:] = [d for d in dirnames if os.access(os.path.join(dirpath, d), os.R_OK)]
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        size = os.path.getsize(filepath)
                        total_size += size
                        file_count += 1
                    except OSError:
                        error_count += 1
        except OSError as e:
            if e.errno != 5:
                logging.warning(f"Error calculating deep size for {path}: {e}")
        self.folder_sizes[path] = total_size
        return total_size

    def get_folder_disk_size_deep(self, path: str) -> int:
        if path in self.folder_disk_sizes:
            return self.folder_disk_sizes[path]
        total_size = 0
        file_count = 0
        error_count = 0
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                dirnames[:] = [d for d in dirnames if os.access(os.path.join(dirpath, d), os.R_OK)]
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        size = self.get_space_on_disk(filepath)
                        total_size += size
                        file_count += 1
                    except OSError:
                        error_count += 1
        except OSError as e:
            if e.errno != 5:
                logging.warning(f"Error calculating disk size for {path}: {e}")
        self.folder_disk_sizes[path] = total_size
        return total_size

    def get_space_on_disk(self, path: str) -> int:
        if not os.path.exists(path):
            return 0
        if platform.system() == 'Windows':
            allocated = self._get_allocated_size_windows(path)
            if allocated is not None:
                return allocated
            try:
                if os.path.isfile(path):
                    stat = os.stat(path)
                    file_size = stat.st_size
                    if file_size == 0:
                        return 0
                    cluster_size = self.get_cluster_size(os.path.dirname(path))
                    return ((file_size + cluster_size - 1) // cluster_size) * cluster_size
            except Exception as e:
                logging.debug(f"Error getting disk space for {path}: {e}")
        return os.path.getsize(path) if os.path.exists(path) else 0

    def _get_allocated_size_windows(self, path: str) -> int | None:
        try:
            high = ctypes.c_ulong(0)
            low = ctypes.windll.kernel32.GetCompressedFileSizeW(ctypes.c_wchar_p(path),
                                                                ctypes.byref(high))
            if low == 0xFFFFFFFF:
                err = ctypes.GetLastError()
                if err:
                    raise ctypes.WinError(err)
            return (high.value << 32) + low
        except Exception as e:
            logging.debug(f"GetCompressedFileSizeW failed for {path}: {e}")
            return None

    def get_cluster_size(self, path: str) -> int:
        try:
            if platform.system() == 'Windows':
                return 4096
        except Exception:
            pass
        return 4096

    def get_folder_size_shallow(self, path: str) -> int:
        total_size = 0
        try:
            for item in os.listdir(path):
                item_path = os.path.join(path, item)
                if os.path.isfile(item_path):
                    try:
                        total_size += os.path.getsize(item_path)
                    except OSError:
                        pass
        except OSError as e:
            if e.errno != 5:
                logging.warning(f"Error calculating shallow size for {path}: {e}")
        return total_size

    def format_size(self, size: int) -> str:
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0
        return f"{size:.2f} PB"

# Configure logging
class TextHandler(logging.Handler):
    def __init__(self, text_widget: Any) -> None:
        logging.Handler.__init__(self)
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        def append() -> None:
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class TreeSizeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("TreeSize - Folder Size Analyzer")
        self.root.geometry("900x700")  # Made taller to accommodate log area
        
        self.current_path = None
        self.size_threads = {}      # Track running threads
        self._calc_start: dict[str, datetime] = {}      # NEW – track batch timers
        
        # Metric selection - default to space on disk
        self.size_metric = tk.StringVar(value="disk")
        self.analyzer = TreeSizeAnalyzer()  # NEW: use analyzer
        
        self.setup_ui()
        
    def setup_ui(self) -> None:
        # --- style / theme -------------------------------------------------
        style = ttk.Style()
        try:
            style.theme_use('clam')                        # modern look
        except tk.TclError:
            pass                                          # keep default if clam not available
        style.configure("Treeview",
                        font=("Segoe UI", 9),
                        rowheight=20)
        style.configure("Treeview.Heading",
                        font=("Segoe UI", 9, "bold"))
        # ------------------------------------------------------------------

        # Top frame for controls
        control_frame = ttk.Frame(self.root, padding="5")
        control_frame.pack(fill=tk.X)
        
        ttk.Button(control_frame, text="Select Folder", command=self.select_folder).pack(side=tk.LEFT, padx=5)
        
        self.path_label = ttk.Label(control_frame, text="No folder selected")
        self.path_label.pack(side=tk.LEFT, padx=10)
        
        # Metric selection frame
        metric_frame = ttk.LabelFrame(control_frame, text="Size Metric")
        metric_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Radiobutton(metric_frame, text="Actual Size", variable=self.size_metric, 
                      value="actual", command=self.on_metric_change).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(metric_frame, text="Space on Disk", variable=self.size_metric, 
                      value="disk", command=self.on_metric_change).pack(side=tk.LEFT, padx=5)
        
        # Main content frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create vertical paned window to separate content and logs
        main_paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True)
        
        # Content area inside the main paned window
        content_frame = ttk.Frame(main_paned)
        main_paned.add(content_frame, weight=3)
        
        # Horizontal paned window for content
        content_paned = ttk.PanedWindow(content_frame, orient=tk.HORIZONTAL)
        content_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left side - folder tree
        left_frame = ttk.Frame(content_paned)
        content_paned.add(left_frame, weight=1)
        
        ttk.Label(left_frame, text="Folders", font=("Arial", 10, "bold")).pack()
        
        # Folder treeview with scrollbar
        folder_scroll = ttk.Scrollbar(left_frame)
        folder_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.folder_tree = ttk.Treeview(
            left_frame,
            columns=("size", "path"),
            yscrollcommand=folder_scroll.set,
            selectmode="browse")
        self.folder_tree.pack(fill=tk.BOTH, expand=True)
        folder_scroll.config(command=self.folder_tree.yview)
        
        self.folder_tree.heading("#0", text="Folder", anchor=tk.W)
        self.folder_tree.heading("size", text="Size", anchor=tk.E)
        self.folder_tree.column("size", width=90, anchor=tk.E, stretch=False)   # NARROWER
        self.folder_tree.column("path", width=0, stretch=False)  # Hidden column for path
        
        self.folder_tree.bind("<<TreeviewOpen>>", self.on_folder_expand)
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_folder_select)
        
        # Right side - file list
        right_frame = ttk.Frame(content_paned)
        content_paned.add(right_frame, weight=1)
        
        ttk.Label(right_frame, text="Top 50 Files", font=("Arial", 10, "bold")).pack()
        
        # File treeview with scrollbar
        file_scroll = ttk.Scrollbar(right_frame)
        file_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_tree = ttk.Treeview(
            right_frame,
            columns=("size", "path"),
            yscrollcommand=file_scroll.set,
            selectmode="browse")
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        file_scroll.config(command=self.file_tree.yview)
        
        self.file_tree.heading("#0", text="File Name", anchor=tk.W)
        self.file_tree.heading("size", text="Size", anchor=tk.E)
        self.file_tree.column("size", width=90, anchor=tk.E, stretch=False)     # NARROWER
        self.file_tree.column("path", width=0, stretch=False)  # NEW hide full path
        
        # Log area frame
        log_frame = ttk.LabelFrame(main_paned, text="Log")
        main_paned.add(log_frame, weight=1)
        
        # Log text area with scrollbar
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=8, width=50, state='disabled', 
                               wrap=tk.WORD, yscrollcommand=log_scroll.set)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        log_scroll.config(command=self.log_text.yview)
        
        # Set up the text handler for logging
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(text_handler)
        
        # Status bar
        self.status_bar = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
        ttk.Button(control_frame, text="Refresh", command=self.refresh).pack(side=tk.LEFT, padx=5)        # NEW
        ttk.Button(control_frame, text="Open Folder", command=self.open_selected_folder).pack(side=tk.RIGHT, padx=5)         # NEW
        ttk.Button(control_frame, text="Open File Location", command=self.open_selected_file_location).pack(side=tk.RIGHT)    # NEW
        
    def on_metric_change(self) -> None:
        """Handle change in size metric selection"""
        metric = self.size_metric.get()
        logger.info(f"Size metric changed to: {metric}")
        
        # If we have a current path, refresh the view
        if self.current_path:
            # Clear the cache for the alternate metric
            if metric == "actual":
                self.analyzer.folder_disk_sizes.clear()
            else:
                self.analyzer.folder_sizes.clear()
                
            # Update displayed sizes for visible items
            self.update_displayed_sizes()
            
    def update_displayed_sizes(self) -> None:
        """Update all displayed sizes based on the current metric"""
        # Update folder tree
        for item in self.folder_tree.get_children(""):
            self.update_item_sizes(item)
            # Also update any expanded children
            for child in self.folder_tree.get_children(item):
                self.update_item_sizes(child)
        
        # Also update file list if any folder is selected
        selection = self.folder_tree.selection()
        if selection:
            item = self.folder_tree.item(selection[0])
            if len(item['values']) > 1:
                folder_path = item['values'][1]
                if folder_path and folder_path in self.analyzer.file_cache:
                    # Reload files with new metric
                    self._update_file_tree(self.analyzer.file_cache[folder_path])
                    
        self._sort_tree_by_size("")                       # NEW – resort root
        for top in self.folder_tree.get_children(""):     # NEW – resort each branch
            self._sort_tree_by_size(top)
            
    def update_item_sizes(self, item: str) -> None:
        """Update size display for a single tree item"""
        item_data = self.folder_tree.item(item)
        path = item_data['values'][1]
        
        if not path:
            return
            
        size = self.get_folder_size(path)
        if size > 0:
            size_str = self.analyzer.format_size(size)
            self.folder_tree.set(item, "size", size_str)
            
    def get_folder_size(self, path: str) -> int:
        """Get folder size based on selected metric"""
        metric = self.size_metric.get()
        if metric == "actual":
            return self.analyzer.get_folder_size_deep(path)
        else:
            return self.analyzer.get_folder_disk_size_deep(path)
            
    def select_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            logger.info(f"Selected folder: {folder}")
            self.current_path = folder
            self.path_label.config(text=f"Current: {folder}")
            self.folder_tree.delete(*self.folder_tree.get_children())
            self.load_folder_contents(folder, parent="")
            
    def load_folder_contents(self, path: str, parent: str = "") -> None:
        """Load only the immediate contents of a folder"""
        logger.debug(f"Loading contents of: {path}")
        self.status_bar.config(text=f"Loading: {os.path.basename(path) or path}")
        
        # Run in thread to avoid freezing UI
        thread = threading.Thread(target=self._load_folder_thread, args=(path, parent))
        thread.daemon = True
        thread.start()
        
    def _load_folder_thread(self, path: str, parent: str) -> None:
        start_time = datetime.now()
        try:
            items = []
            # Only list immediate contents
            for item in os.listdir(path):
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    # Check if we have a cached size for the current metric
                    size = 0
                    if self.size_metric.get() == "actual":
                        size = self.analyzer.folder_sizes.get(item_path, 0)
                    else:
                        size = self.analyzer.folder_disk_sizes.get(item_path, 0)
                    
                    # Initially show just the folder with placeholder or cached size
                    items.append((item, size, item_path, True))
                    
            # Sort by name initially (will resort by size after calculation)
            items.sort(key=lambda x: x[0].lower())
            
            elapsed = (datetime.now() - start_time).total_seconds()
            logger.info(f"Loaded {len(items)} folders from {path} in {elapsed:.2f}s")
            
            # Update UI in main thread
            self.root.after(0, self._update_folder_tree, items, parent, path)
            
        except Exception as e:
            logger.error(f"Error loading folder {path}: {e}")
            self.root.after(0, lambda: self.status_bar.config(text=f"Error: {str(e)}"))
            
    def _update_folder_tree(self, items: List[Any], parent: str, parent_path: str) -> None:
        nodes = []
        for name, size, path, is_dir in items:
            size_str = self.analyzer.format_size(size) if size > 0 else "Calculating..."
            node = self.folder_tree.insert(parent, tk.END, text=name, 
                                         values=(size_str, path), open=False)
            nodes.append((node, path))
            
            # Add dummy child to show expand arrow
            if is_dir:
                self.folder_tree.insert(node, tk.END, text="Loading...", values=("", ""))
                
        self.status_bar.config(text="Ready")
        
        # Start calculating sizes for visible folders
        calculation_needed = []
        for node, folder_path in nodes:
            # Check if we need calculation based on current metric
            needs_calculation = False
            if self.size_metric.get() == "actual":
                needs_calculation = folder_path not in self.analyzer.folder_sizes
            else:
                needs_calculation = folder_path not in self.analyzer.folder_disk_sizes
                
            if needs_calculation:
                calculation_needed.append((node, folder_path))
        
        if calculation_needed:
            self._calc_start[parent] = datetime.now()    # NEW – start timer
            logger.info(f"Calculating sizes for {len(calculation_needed)} folders...")
            for node, folder_path in calculation_needed:
                thread = threading.Thread(target=self._calculate_and_update_size, args=(node, folder_path, parent))
                thread.daemon = True
                thread.start()
        else:
            # If all sizes already calculated, sort nodes by size
            self._sort_tree_by_size(parent)
        
    def _calculate_and_update_size(self, node: str, folder_path: str, parent: str) -> None:
        """Calculate folder size and update the tree"""
        try:
            start = datetime.now()
            size = self.get_folder_size(folder_path)
            size_str = self.analyzer.format_size(size)
            
            # Update UI
            self.root.after(0, lambda: self.folder_tree.set(node, "size", size_str))
            # After a batch of updates, sort the tree
            self.root.after(100, lambda: self._sort_tree_by_size(parent))
            
            # --- NEW : finished-batch detection ---
            def _maybe_done() -> bool:
                for child in self.folder_tree.get_children(parent):
                    val = self.folder_tree.item(child, "values")[0]
                    if val in ("Calculating...", "Error"):
                        return False
                return True

            if _maybe_done() and parent in self._calc_start:
                elapsed = (datetime.now() - self._calc_start.pop(parent)).total_seconds()
                msg = f"Done calculating sizes for '{self.folder_tree.item(parent, 'text')}' in {elapsed:.2f}s"
                logger.info(msg)
                self.status_bar.config(text=msg)
            # --------------------------------------
        except Exception as e:
            logger.error(f"Error calculating size for {folder_path}: {e}")
            self.root.after(0, lambda: self.folder_tree.set(node, "size", "Error"))
    
    def _sort_tree_by_size(self, parent: str) -> None:
        """Sort folders by size in descending order"""
        items = self.folder_tree.get_children(parent)
        if not items:
            return
            
        # Get all items with their sizes
        item_list = []
        for item in items:
            size_str = self.folder_tree.item(item, "values")[0]
            # Skip items still calculating
            if size_str == "Calculating..." or size_str == "Error":
                continue
                
            path = self.folder_tree.item(item, "values")[1]
            # Get size based on current metric
            size = 0
            if self.size_metric.get() == "actual":
                size = self.analyzer.folder_sizes.get(path, 0)
            else:
                size = self.analyzer.folder_disk_sizes.get(path, 0)
                
            item_list.append((item, size))
            
        # Sort by size
        item_list.sort(key=lambda x: x[1], reverse=True)
        
        # Reorder in tree
        for idx, (item, _) in enumerate(item_list):
            self.folder_tree.move(item, parent, idx)
    
    def load_top_files(self, folder_path: str) -> None:
        # Check if we have cached results
        if folder_path in self.analyzer.file_cache:
            logger.info(f"Using cached file list for {folder_path}")
            self._update_file_tree(self.analyzer.file_cache[folder_path])
            return
            
        self.file_tree.delete(*self.file_tree.get_children())
        self.status_bar.config(text=f"Loading files from: {os.path.basename(folder_path)}")
        
        thread = threading.Thread(target=self._load_files_thread, args=(folder_path,))
        thread.daemon = True
        thread.start()
        
    def _load_files_thread(self, folder_path: str) -> None:
        start_time = datetime.now()
        try:
            files = []
            file_count = 0
            error_count = 0
            
            for dirpath, dirnames, filenames in os.walk(folder_path):
                # Skip directories we can't access
                dirnames[:] = [d for d in dirnames if os.access(os.path.join(dirpath, d), os.R_OK)]
                
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        # Get size based on current metric
                        if self.size_metric.get() == "actual":
                            size = os.path.getsize(filepath)
                        else:
                            size = self.analyzer.get_space_on_disk(filepath)
                            
                        rel_path = os.path.relpath(filepath, folder_path)
                        files.append((rel_path, size, filepath))  # Store full path for later metric changes
                        file_count += 1
                    except OSError:
                        error_count += 1
                        
            # Sort by size and get top 50
            files.sort(key=lambda x: x[1], reverse=True)
            top_files = files[:50]
            
            # Cache the results
            self.analyzer.file_cache[folder_path] = top_files
            
            elapsed = (datetime.now() - start_time).total_seconds()
            logger.info(f"Found {file_count} files in {folder_path}, {error_count} errors, showing top 50 (took {elapsed:.2f}s)")
            
            # Update UI in main thread
            self.root.after(0, self._update_file_tree, top_files)
            
        except Exception as e:
            logger.error(f"Error loading files from {folder_path}: {e}")
            self.root.after(0, lambda: self.status_bar.config(text=f"Error: {str(e)}"))
            
    def _update_file_tree(self, files: List[Any]) -> None:
        self.file_tree.delete(*self.file_tree.get_children())
        for filename, size, full_path in files:             # CHANGED – keep full path
            size_str = self.analyzer.format_size(size)
            self.file_tree.insert("", tk.END, text=filename, values=(size_str, full_path))
        self.status_bar.config(text="Ready")
        
    # ---------- NEW METHODS ----------
    def refresh(self) -> None:
        """Clear caches and reload current folder."""
        if not self.current_path:
            messagebox.showinfo("Refresh", "No folder selected.")
            return
        logger.info("Refreshing view…")
        self.analyzer.folder_sizes.clear()
        self.analyzer.folder_disk_sizes.clear()
        self.analyzer.file_cache.clear()
        self.folder_tree.delete(*self.folder_tree.get_children())
        self.file_tree.delete(*self.file_tree.get_children())
        self.load_folder_contents(self.current_path, parent="")

    def open_selected_folder(self) -> None:
        """Open highlighted folder in system file explorer."""
        sel = self.folder_tree.selection()
        if not sel:
            messagebox.showinfo("Open Folder", "No folder selected.")
            return
        path = self.folder_tree.item(sel[0])['values'][1]
        if not path or not os.path.exists(path):
            return
        logger.info(f"Opening folder: {path}")
        if platform.system() == 'Windows':
            os.startfile(path)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', path], check=False)
        else:
            subprocess.run(['xdg-open', path], check=False)

    def open_selected_file_location(self) -> None:
        """Reveal selected file in explorer / finder."""
        sel = self.file_tree.selection()
        if not sel:
            messagebox.showinfo("Open File Location", "No file selected.")
            return
        path = self.file_tree.item(sel[0])['values'][1]
        if not path or not os.path.exists(path):
            return
        logger.info(f"Opening file location: {path}")
        if platform.system() == 'Windows':
            # Normalize path and keep '/select,' in the same argument list
            norm_path = os.path.normpath(path)
            subprocess.run(['explorer', '/select,', norm_path], check=False)
        elif platform.system() == 'Darwin':
            subprocess.run(['open', '-R', path], check=False)
        else:
            subprocess.run(['xdg-open', os.path.dirname(path)], check=False)
    # ---------- END NEW METHODS ----------

    def on_folder_expand(self, event: Any) -> None:
        """Handle folder expansion - load subfolders if not already loaded"""
        item = self.folder_tree.focus()
        if not item:
            return

        # Get children
        children = self.folder_tree.get_children(item)

        # If only has dummy child, load real contents
        if len(children) == 1 and self.folder_tree.item(children[0])['text'] == "Loading...":
            # Remove dummy
            self.folder_tree.delete(children[0])

            # Get folder path
            folder_path = self.folder_tree.item(item)['values'][1]
            logger.debug(f"Expanding folder: {folder_path}")

            # Load contents
            self.load_folder_contents(folder_path, item)

    def on_folder_select(self, event: Any) -> None:
        selection = self.folder_tree.selection()
        if selection:
            item = self.folder_tree.item(selection[0])
            if len(item['values']) > 1:  # Check if values exist
                folder_path = item['values'][1]
                if folder_path:  # Ignore dummy items
                    logger.debug(f"Selected folder: {folder_path}")
                    self.load_top_files(folder_path)

def main() -> None:
    root = tk.Tk()
    app = TreeSizeApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
