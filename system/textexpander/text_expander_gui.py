"""
Text Expander GUI Module

Provides a tkinter-based graphical interface for managing text expander abbreviations.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import logging
from typing import Optional, Dict, List, Tuple
from pathlib import Path

from text_expander_core import TextExpanderCore, ConfigError

# Configure module logger
logger = logging.getLogger(__name__)

class TextExpanderGUI:
    """GUI interface for managing text expander abbreviations."""

    def __init__(self, core: TextExpanderCore) -> None:
        """Initialize the GUI.

        Args:
            core: The text expander core instance.
        """
        self.core = core
        self.root: Optional[tk.Tk] = None
        self.abbreviation_list: Optional[tk.Listbox] = None
        self.abbreviation_entry: Optional[tk.Entry] = None
        self.expansion_text: Optional[scrolledtext.ScrolledText] = None
        self.add_button: Optional[ttk.Button] = None
        self.update_button: Optional[ttk.Button] = None
        self.delete_button: Optional[ttk.Button] = None

        logger.info("🖥️ Initializing Text Expander GUI")

    def show(self) -> None:
        """Show the GUI window."""
        try:
            self.root = tk.Tk()
            self.root.title("Text Expander - Manage Abbreviations")
            self.root.geometry("600x500")
            self.root.resizable(True, True)

            # Set window icon (optional)
            try:
                self.root.iconbitmap(default="")  # Use default icon
            except tk.TclError:
                pass  # Ignore icon errors

            self._create_widgets()
            self._load_abbreviations()
            self._bind_events()

            logger.info("✅ Text Expander GUI window created")
            self.root.mainloop()

        except Exception as e:
            error_msg = f"❌ Failed to create GUI: {e}"
            logger.error(error_msg)
            messagebox.showerror("GUI Error", error_msg)

    def _create_widgets(self) -> None:
        """Create all GUI widgets."""
        if not self.root:
            return

        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Text Expander Abbreviations",
                              font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Left panel - Abbreviations list
        left_frame = ttk.LabelFrame(main_frame, text="Abbreviations", padding="5")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)

        # Abbreviations listbox with scrollbar
        list_frame = ttk.Frame(left_frame)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.abbreviation_list = tk.Listbox(list_frame, height=15, width=25)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.abbreviation_list.yview)
        self.abbreviation_list.configure(yscrollcommand=scrollbar.set)

        self.abbreviation_list.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Right panel - Abbreviation details
        right_frame = ttk.LabelFrame(main_frame, text="Abbreviation Details", padding="5")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)

        # Abbreviation entry
        ttk.Label(right_frame, text="Abbreviation:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.abbreviation_entry = ttk.Entry(right_frame, width=30)
        self.abbreviation_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5), padx=(5, 0))

        # Expansion text
        ttk.Label(right_frame, text="Expansion:").grid(row=1, column=0, sticky=tk.NW, pady=(10, 5))
        self.expansion_text = scrolledtext.ScrolledText(right_frame, width=40, height=10, wrap=tk.WORD)
        self.expansion_text.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 5), padx=(5, 0))

        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=(20, 0))

        self.add_button = ttk.Button(buttons_frame, text="Add New", command=self._add_abbreviation)
        self.add_button.grid(row=0, column=0, padx=(0, 10))

        self.update_button = ttk.Button(buttons_frame, text="Update", command=self._update_abbreviation, state=tk.DISABLED)
        self.update_button.grid(row=0, column=1, padx=(0, 10))

        self.delete_button = ttk.Button(buttons_frame, text="Delete", command=self._delete_abbreviation, state=tk.DISABLED)
        self.delete_button.grid(row=0, column=2, padx=(0, 10))

        ttk.Button(buttons_frame, text="Refresh", command=self._load_abbreviations).grid(row=0, column=3)

    def _bind_events(self) -> None:
        """Bind event handlers to widgets."""
        if self.abbreviation_list:
            self.abbreviation_list.bind('<<ListboxSelect>>', self._on_abbreviation_select)

        if self.abbreviation_entry:
            self.abbreviation_entry.bind('<KeyRelease>', self._on_abbreviation_entry_change)

    def _load_abbreviations(self) -> None:
        """Load abbreviations from core and populate the list."""
        try:
            if not self.abbreviation_list:
                return

            # Clear current list
            self.abbreviation_list.delete(0, tk.END)

            # Get abbreviations from core
            abbreviations = self.core.get_all_abbreviations()

            # Sort abbreviations for consistent display
            sorted_abbreviations = sorted(abbreviations.keys())

            # Populate list
            for abbreviation in sorted_abbreviations:
                display_text = abbreviation  # Could be enhanced to show preview
                self.abbreviation_list.insert(tk.END, display_text)

            logger.debug(f"📋 Loaded {len(abbreviations)} abbreviations into GUI")

        except Exception as e:
            error_msg = f"❌ Failed to load abbreviations: {e}"
            logger.error(error_msg)
            messagebox.showerror("Load Error", error_msg)

    def _on_abbreviation_select(self, event: tk.Event) -> None:
        """Handle abbreviation selection from list."""
        try:
            if not self.abbreviation_list or not self.abbreviation_entry or not self.expansion_text:
                return

            selection = self.abbreviation_list.curselection()
            if not selection:
                return

            index = selection[0]
            abbreviation = self.abbreviation_list.get(index)

            # Get expansion from core
            abbreviations = self.core.get_all_abbreviations()
            expansion = abbreviations.get(abbreviation, "")

            # Update form fields
            self.abbreviation_entry.delete(0, tk.END)
            self.abbreviation_entry.insert(0, abbreviation)

            self.expansion_text.delete(1.0, tk.END)
            self.expansion_text.insert(tk.END, expansion)

            # Enable update and delete buttons
            if self.update_button and self.delete_button:
                self.update_button.config(state=tk.NORMAL)
                self.delete_button.config(state=tk.NORMAL)

            logger.debug(f"🔍 Selected abbreviation: {abbreviation}")

        except Exception as e:
            error_msg = f"❌ Failed to select abbreviation: {e}"
            logger.error(error_msg)
            messagebox.showerror("Selection Error", error_msg)

    def _on_abbreviation_entry_change(self, event: tk.Event) -> None:
        """Handle changes to abbreviation entry field."""
        try:
            if not self.abbreviation_entry or not self.update_button or not self.delete_button:
                return

            abbreviation = self.abbreviation_entry.get().strip()

            # Check if abbreviation exists
            abbreviations = self.core.get_all_abbreviations()
            exists = abbreviation in abbreviations

            # Update button states
            if exists:
                self.update_button.config(state=tk.NORMAL)
                self.delete_button.config(state=tk.NORMAL)
            else:
                self.update_button.config(state=tk.DISABLED)
                self.delete_button.config(state=tk.DISABLED)

        except Exception as e:
            logger.error(f"❌ Error handling abbreviation entry change: {e}")

    def _add_abbreviation(self) -> None:
        """Add a new abbreviation."""
        try:
            if not self.abbreviation_entry or not self.expansion_text:
                return

            abbreviation = self.abbreviation_entry.get().strip()
            expansion = self.expansion_text.get(1.0, tk.END).strip()

            # Validate inputs
            if not abbreviation:
                messagebox.showwarning("Input Error", "Please enter an abbreviation.")
                return

            if not expansion:
                messagebox.showwarning("Input Error", "Please enter expansion text.")
                return

            # Remove trigger key if present
            settings = self.core.get_settings()
            if settings and abbreviation.startswith(settings.trigger_key):
                abbreviation = abbreviation[len(settings.trigger_key):]

            # Add abbreviation
            if self.core.add_abbreviation(abbreviation, expansion):
                messagebox.showinfo("Success", f"Abbreviation '{settings.trigger_key if settings else '/'}{abbreviation}' added successfully!")
                self._load_abbreviations()
                self._clear_form()
                logger.info(f"✅ Added abbreviation: {abbreviation}")
            else:
                messagebox.showerror("Error", "Failed to add abbreviation. It may already exist.")

        except Exception as e:
            error_msg = f"❌ Failed to add abbreviation: {e}"
            logger.error(error_msg)
            messagebox.showerror("Add Error", error_msg)

    def _update_abbreviation(self) -> None:
        """Update an existing abbreviation."""
        try:
            if not self.abbreviation_entry or not self.expansion_text:
                return

            abbreviation = self.abbreviation_entry.get().strip()
            expansion = self.expansion_text.get(1.0, tk.END).strip()

            # Validate inputs
            if not abbreviation:
                messagebox.showwarning("Input Error", "Please enter an abbreviation.")
                return

            if not expansion:
                messagebox.showwarning("Input Error", "Please enter expansion text.")
                return

            # Update abbreviation
            if self.core.update_abbreviation(abbreviation, expansion):
                messagebox.showinfo("Success", f"Abbreviation '{abbreviation}' updated successfully!")
                self._load_abbreviations()
                logger.info(f"✅ Updated abbreviation: {abbreviation}")
            else:
                messagebox.showerror("Error", "Failed to update abbreviation.")

        except Exception as e:
            error_msg = f"❌ Failed to update abbreviation: {e}"
            logger.error(error_msg)
            messagebox.showerror("Update Error", error_msg)

    def _delete_abbreviation(self) -> None:
        """Delete an existing abbreviation."""
        try:
            if not self.abbreviation_entry:
                return

            abbreviation = self.abbreviation_entry.get().strip()

            # Confirm deletion
            if not messagebox.askyesno("Confirm Delete",
                                     f"Are you sure you want to delete the abbreviation '{abbreviation}'?"):
                return

            # Delete abbreviation
            if self.core.delete_abbreviation(abbreviation):
                messagebox.showinfo("Success", f"Abbreviation '{abbreviation}' deleted successfully!")
                self._load_abbreviations()
                self._clear_form()
                logger.info(f"✅ Deleted abbreviation: {abbreviation}")
            else:
                messagebox.showerror("Error", "Failed to delete abbreviation.")

        except Exception as e:
            error_msg = f"❌ Failed to delete abbreviation: {e}"
            logger.error(error_msg)
            messagebox.showerror("Delete Error", error_msg)

    def _clear_form(self) -> None:
        """Clear the form fields."""
        try:
            if self.abbreviation_entry:
                self.abbreviation_entry.delete(0, tk.END)

            if self.expansion_text:
                self.expansion_text.delete(1.0, tk.END)

            if self.update_button and self.delete_button:
                self.update_button.config(state=tk.DISABLED)
                self.delete_button.config(state=tk.DISABLED)

        except Exception as e:
            logger.error(f"❌ Failed to clear form: {e}")

    def close(self) -> None:
        """Close the GUI window."""
        if self.root:
            self.root.quit()
            self.root.destroy()
            logger.info("🖥️ Text Expander GUI closed")
