#!/usr/bin/env python3
"""
LinkedIn Profile Opener - Refactored version
Opens LinkedIn profiles from Excel data with filtering options.
"""

import argparse
import json
import logging
import sys
import threading
from pathlib import Path
from typing import Dict, Any, Tuple

import pandas as pd
import webbrowser
import easygui
import tkinter as tk
from tkinter import ttk, messagebox


class GuiLogHandler(logging.Handler):
    """Custom logging handler that sends messages to GUI."""

    def __init__(self, gui_callback):
        super().__init__()
        self.gui_callback = gui_callback

    def emit(self, record):
        """Send the log message to the GUI."""
        try:
            msg = self.format(record)
            self.gui_callback(msg)
        except Exception:
            self.handleError(record)


def setup_logging(verbose: bool = False) -> logging.Logger:
    """Setup logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )
    return logging.getLogger(__name__)


def load_config(config_path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        logger.error(f"Configuration file not found: {config_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in configuration file: {e}")
        sys.exit(1)


def update_linkedin_url(url: str, verbose: bool = False) -> str:
    """Update LinkedIn URLs based on the presence of '/company/'."""
    if "/company/" in url:
        url_fix = url + "posts/?feedView=all"
    else:
        url_fix = url + "recent-activity/shares/"
    
    if verbose:
        logger.debug(f"Updated LinkedIn URL: {url_fix}")
    return url_fix


def load_excel_data(file_path: str) -> pd.DataFrame:
    """Load and validate Excel data."""
    logger.info(f"Reading data from: {file_path}")
    
    try:
        df = pd.read_excel(file_path)
        logger.info(f"Successfully loaded {len(df)} rows from Excel file")
        return df
    except PermissionError:
        easygui.msgbox(
            f"Permission denied: Unable to access the file at {file_path}.",
            title="File Access Error"
        )
        logger.error(f"Permission denied accessing file: {file_path}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
        sys.exit(1)


class FilterDialog:
    """Custom tkinter dialog for LinkedIn profile filter selection."""

    def __init__(self, parent, config: Dict[str, Any]):
        self.parent = parent
        self.config = config
        self.result = None
        self.df = None  # Will hold the loaded dataframe
        self.filtered_df = None  # Will hold filtered data
        self.processing = False  # Flag to track if processing is active
        self.continue_event = threading.Event()  # For synchronizing continue actions

        logger.debug("Creating FilterDialog...")

        # Create dialog window without parent to avoid hidden parent issues
        self.dialog = tk.Tk()
        self.dialog.title("LinkedIn Profile Opener - Filter Options")
        self.dialog.geometry("600x600")  # Made wider and taller for logs
        self.dialog.resizable(False, False)

        logger.debug("Dialog window created")

        # Center the dialog on screen
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 600) // 2
        self.dialog.geometry(f"+{x}+{y}")

        logger.debug("Creating widgets...")
        self.create_widgets()
        self.load_defaults()

        # Set up GUI logging handler
        self.gui_handler = GuiLogHandler(self.add_log_message)
        self.gui_handler.setFormatter(logging.Formatter('%(message)s'))  # Clean format for GUI
        # Add the GUI handler to the root logger temporarily
        self.original_handlers = logging.getLogger().handlers[:]
        logging.getLogger().addHandler(self.gui_handler)

        # Handle window close
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)

        # Make sure dialog is visible and focused
        self.dialog.deiconify()
        self.dialog.focus_force()
        self.dialog.lift()

        logger.debug("Waiting for dialog to close...")
        # Wait for dialog to close
        self.dialog.mainloop()

        # Restore original logging handlers
        logging.getLogger().removeHandler(self.gui_handler)
        for handler in self.original_handlers:
            logging.getLogger().addHandler(handler)

        logger.debug("Dialog closed")

    def create_widgets(self):
        """Create all dialog widgets."""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="LinkedIn Profile Opener",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 10))

        # Instructions
        instructions = """Choose your filtering options below. All filters are applied together.
Click OK to start processing and see logs below."""
        instr_label = ttk.Label(main_frame, text=instructions, wraplength=550, justify=tk.LEFT)
        instr_label.pack(pady=(0, 15))

        # Filter frame
        filter_frame = ttk.LabelFrame(main_frame, text="Filter Options", padding="10")
        filter_frame.pack(fill=tk.X, pady=(0, 15))

        # Alert filter
        ttk.Label(filter_frame, text="Alert Filter (IND_ALERT):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.alert_var = tk.StringVar()
        self.alert_combo = ttk.Combobox(filter_frame, textvariable=self.alert_var, state="readonly", width=15)
        self.alert_combo.grid(row=0, column=1, padx=(10, 0), pady=5)

        # Star filter
        ttk.Label(filter_frame, text="Star Filter (IND_STAR):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.star_var = tk.StringVar()
        self.star_combo = ttk.Combobox(filter_frame, textvariable=self.star_var, state="readonly", width=15)
        self.star_combo.grid(row=1, column=1, padx=(10, 0), pady=5)

        # Circle filter
        ttk.Label(filter_frame, text="Circle Filter:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.circle_var = tk.StringVar()
        self.circle_combo = ttk.Combobox(filter_frame, textvariable=self.circle_var, state="readonly", width=15)
        self.circle_combo.grid(row=2, column=1, padx=(10, 0), pady=5)

        # Topic filter
        ttk.Label(filter_frame, text="Topic Filter (FK_TOPIC):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.topic_var = tk.StringVar()
        self.topic_combo = ttk.Combobox(filter_frame, textvariable=self.topic_var, state="readonly", width=20)
        self.topic_combo.grid(row=3, column=1, padx=(10, 0), pady=5)

        # Profiles to open
        ttk.Label(filter_frame, text="Profiles to Open at Once:").grid(row=4, column=0, sticky=tk.W, pady=10)
        self.chunk_size_var = tk.StringVar()
        self.chunk_size_entry = ttk.Entry(filter_frame, textvariable=self.chunk_size_var, width=17)
        self.chunk_size_entry.grid(row=4, column=1, padx=(10, 0), pady=10)

        # Buttons (put at top after instructions for better visibility)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 15))

        # Status label on the left
        self.status_label = ttk.Label(button_frame, text="Ready to start", foreground="blue")
        self.status_label.pack(side=tk.LEFT)

        # Buttons on the right
        self.cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        self.cancel_button.pack(side=tk.RIGHT)

        self.ok_button = ttk.Button(button_frame, text="Start Processing", command=self.start_processing)
        self.ok_button.pack(side=tk.RIGHT, padx=(10, 0))

        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Logs", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Text widget for logs with scrollbar
        self.log_text = tk.Text(log_frame, height=12, width=60, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_defaults(self):
        """Load default values into the widgets."""
        filter_options = self.config.get('filter_options', {})

        # Set combobox values
        self.alert_combo['values'] = filter_options.get('alert_choices', ['True', 'False', 'Any'])
        self.star_combo['values'] = filter_options.get('star_choices', ['True', 'False', 'Any'])
        self.circle_combo['values'] = filter_options.get('circle_choices', ['Any', 'Top5', 'Key50', 'Vital100', 'None'])
        self.topic_combo['values'] = filter_options.get('topic_choices', ['all', 'innovation', 'personal development', 'leadership and management', 'LinkedIn', 'visual illustration'])

        # Set default values
        self.alert_var.set("True")  # Default: alert true
        self.star_var.set("False")  # Default: star false
        self.circle_var.set("Any")  # Default: circle any
        self.topic_var.set("all")  # Default: topic all
        self.chunk_size_var.set(str(self.config.get('default_chunk_size', 10)))  # Default: 10

    def validate_inputs(self) -> bool:
        """Validate user inputs."""
        try:
            chunk_size = int(self.chunk_size_var.get())
            if chunk_size <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid positive number for 'Profiles to Open at Once'")
            return False

        return True

    def add_log_message(self, message: str):
        """Add a message to the log text area."""
        def _add_message():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)  # Auto-scroll to bottom
            self.log_text.config(state=tk.DISABLED)
            self.status_label.config(text=f"Processing: {message[:50]}..." if len(message) > 50 else f"Processing: {message}")
            self.dialog.update_idletasks()  # Force GUI update

        # Schedule the update on the main thread
        self.dialog.after(0, _add_message)

    def start_processing(self):
        """Handle Start Processing button click."""
        if not self.validate_inputs():
            return

        if self.processing:
            return  # Already processing

        self.processing = True
        # Update UI on main thread
        self.dialog.after(0, lambda: self.status_label.config(text="Processing...", foreground="orange"))
        self.dialog.after(0, lambda: self.ok_button.config(state=tk.DISABLED))

        # Get filter values
        filters = {
            'alert': self.alert_var.get(),
            'star': self.star_var.get(),
            'circle': self.circle_var.get(),
            'topic': self.topic_var.get()
        }
        chunk_size = int(self.chunk_size_var.get())

        # Disable filter controls during processing
        self.alert_combo.config(state=tk.DISABLED)
        self.star_combo.config(state=tk.DISABLED)
        self.circle_combo.config(state=tk.DISABLED)
        self.topic_combo.config(state=tk.DISABLED)
        self.chunk_size_entry.config(state=tk.DISABLED)

        # Start processing in a separate thread to avoid blocking GUI
        import threading
        processing_thread = threading.Thread(target=self.process_data, args=(filters, chunk_size))
        processing_thread.daemon = True
        processing_thread.start()

    def process_data(self, filters: Dict[str, str], chunk_size: int):
        """Process the data and open profiles."""
        try:
            # Load data
            self.add_log_message("Loading Excel data...")
            file_path = self.config['file_path']
            self.df = load_excel_data(file_path)

            # Sort data
            self.df = self.df.sort_values(by=['FK_TOPIC', 'DE_PERSON'])
            self.add_log_message("Data sorted by topic and person")

            # Apply filters
            self.filtered_df = apply_filters(self.df, filters)

            # Update URLs
            self.filtered_df = update_urls(self.filtered_df, self.config.get('verbose', False))

            # Open profiles
            self.add_log_message(f"Opening {len(self.filtered_df)} LinkedIn profiles in chunks of {chunk_size}")
            self.open_profiles_gui(chunk_size)

            self.add_log_message("Processing completed successfully!")
            self.status_label.config(text="Completed", foreground="green")
            # Update button on main thread
            self.dialog.after(0, lambda: self.ok_button.config(text="Close", command=self.close_successfully, state=tk.NORMAL))

        except Exception as e:
            self.add_log_message(f"Error during processing: {e}")
            self.status_label.config(text="Error", foreground="red")
            self.ok_button.config(state=tk.NORMAL)

        finally:
            self.processing = False

    def open_profiles_gui(self, chunk_size: int):
        """Open LinkedIn profiles with GUI logging."""
        total_urls = len(self.filtered_df)
        if total_urls == 0:
            self.add_log_message("No profiles to open after filtering")
            return

        num_urls_opened = 0

        for i in range(0, total_urls, chunk_size):
            chunk = self.filtered_df.iloc[i:i+chunk_size]

            for j, (_, row) in enumerate(chunk.iterrows(), start=1):
                num_urls_opened += 1
                url = row['URL_LINKEDIN']
                topic = row['FK_TOPIC']
                person = row['DE_PERSON']

                self.add_log_message(f"Opening URL {num_urls_opened:03} ({topic} - {person})")
                webbrowser.open(url)

            # Ask for continuation if there are more URLs
            if num_urls_opened < total_urls:
                remaining = total_urls - num_urls_opened
                next_chunk = min(chunk_size, remaining)
                self.add_log_message(f"Opened {num_urls_opened} profiles. Press Continue to open next {next_chunk} profiles.")

                # Reset the continue event and set up the button
                self.continue_event.clear()
                self.ok_button.config(text=f"Continue ({next_chunk} more)",
                                    command=self.signal_continue, state=tk.NORMAL)

                # Wait for user to click continue (this blocks the processing thread)
                self.continue_event.wait()

                # Reset button
                self.ok_button.config(text="Processing...", state=tk.DISABLED)

        self.add_log_message("All profiles opened!")

    def signal_continue(self):
        """Signal that the user wants to continue processing."""
        self.continue_event.set()

    def close_successfully(self):
        """Close the dialog after successful completion."""
        self.result = {
            'alert': self.alert_var.get(),
            'star': self.star_var.get(),
            'circle': self.circle_var.get(),
            'topic': self.topic_var.get(),
            'chunk_size': int(self.chunk_size_var.get())
        }
        self.dialog.destroy()

    def cancel(self):
        """Handle Cancel button click."""
        if self.processing:
            # Ask for confirmation if processing
            if messagebox.askyesno("Cancel Processing", "Are you sure you want to cancel? Processing will stop."):
                self.add_log_message("Processing cancelled by user")
                self.processing = False
                self.result = None
                self.dialog.destroy()
        else:
            self.result = None
            self.dialog.destroy()


def get_user_filters_and_chunk_size(config: Dict[str, Any]) -> Tuple[Dict[str, str], int]:
    """Get user filter preferences and chunk size through single GUI screen with dropdowns."""
    logger.info("Opening filter selection dialog...")
    try:
        # Show the filter dialog (now self-contained)
        dialog = FilterDialog(None, config)
        logger.debug("FilterDialog created and completed")

        if dialog.result is None:  # User cancelled
            logger.info("User cancelled the operation")
            sys.exit(0)

        logger.info("Filter selection completed")

        filters = {
            'alert': dialog.result['alert'],
            'star': dialog.result['star'],
            'circle': dialog.result['circle'],
            'topic': dialog.result['topic']
        }

        chunk_size = dialog.result['chunk_size']

        return filters, chunk_size
    except Exception as e:
        logger.error(f"Error in GUI dialog: {e}")
        raise


def apply_filters(df: pd.DataFrame, filters: Dict[str, str]) -> pd.DataFrame:
    """Apply user filters to the dataframe."""
    logger.info("Applying filters to data...")
    
    # Start with basic mask (URL_LINKEDIN not null)
    mask = df['URL_LINKEDIN'].notnull()
    
    # Apply alert filter
    if filters['alert'] != 'Any':
        mask &= (df['IND_ALERT'] == (filters['alert'] == 'True'))
    
    # Apply star filter
    if filters['star'] != 'Any':
        mask &= (df['IND_STAR'] == (filters['star'] == 'True'))
    
    # Apply circle filter
    if filters['circle'] == 'None':
        mask &= df['FK_CIRCLE'].isnull()
    elif filters['circle'] != 'Any':
        mask &= (df['FK_CIRCLE'] == filters['circle'])
    
    # Apply topic filter
    if filters['topic'] != 'all':
        mask &= (df['FK_TOPIC'] == filters['topic'])
    
    filtered_df = df.loc[mask].copy()
    logger.info(f"Filtered data. Number of rows: {len(filtered_df)}")
    
    return filtered_df


def update_urls(df: pd.DataFrame, verbose: bool = False) -> pd.DataFrame:
    """Update LinkedIn URLs in the dataframe."""
    logger.info("Updating LinkedIn URLs...")
    df['URL_LINKEDIN'] = df['URL_LINKEDIN'].apply(lambda url: update_linkedin_url(url, verbose))
    return df


def main():
    """Main execution function."""
    # Get the directory where this script is located
    script_dir = Path(__file__).parent

    parser = argparse.ArgumentParser(
        description="Open LinkedIn profiles from Excel data with filtering options",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python linkedin_open.py
  python linkedin_open.py --config custom_config.json
  python linkedin_open.py --verbose --chunk-size 5
        """
    )

    parser.add_argument(
        '--config',
        default=str(script_dir / 'linkedin_open.json'),
        help='Path to JSON configuration file (default: linkedin_open.json in script directory)'
    )

    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose logging'
    )

    parser.add_argument(
        '--chunk-size',
        type=int,
        help='Number of profiles to open at once (overrides config)'
    )

    args = parser.parse_args()

    # Setup logging
    global logger
    logger = setup_logging(args.verbose)

    # Load configuration
    config = load_config(args.config)
    logger.info(f"Loaded configuration from: {args.config}")

    # Override config with command line arguments
    if args.chunk_size:
        config['default_chunk_size'] = args.chunk_size

    # Launch the GUI dialog (it handles all processing internally; the returned
    # filter selection is not needed here)
    get_user_filters_and_chunk_size(config)

    logger.info("Application finished")


if __name__ == "__main__":
    main()
