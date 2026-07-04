#!/usr/bin/env python3
"""
Apple Contacts Converter

This module converts vCard (.vcf) files to Apple-compatible format using a Tkinter GUI.
It automatically adds the _apple suffix to the output filename.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import logging
from pathlib import Path
from typing import List, Dict, Any
import re

from _vcard_fields import extract_type_label

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class AppleContactConverter:
    """Convert vCard files to Apple-compatible format."""
    
    def __init__(self):
        """Initialize the converter with Apple-specific settings."""
        self.vcard_version = "3.0"
        self.required_fields = ["BEGIN:VCARD", "VERSION:", "FN:", "END:VCARD"]
        
    def convert_vcard_to_apple_format(self, input_content: str) -> str:
        """Convert vCard content to Apple-compatible format.
        
        Args:
            input_content: Raw vCard content from input file
            
        Returns:
            Converted vCard content optimized for Apple
        """
        # Split into individual vCard entries
        vcard_entries = input_content.split("BEGIN:VCARD")
        converted_entries = []
        
        for entry in vcard_entries:
            if not entry.strip():
                continue
                
            # Add BEGIN:VCARD back
            entry = "BEGIN:VCARD" + entry
            
            # Convert to Apple format
            converted_entry = self._convert_single_vcard(entry)
            if converted_entry:
                converted_entries.append(converted_entry)
        
        return "\n\n".join(converted_entries)
    
    def _convert_single_vcard(self, vcard_content: str) -> str:
        """Convert a single vCard entry to Apple format.
        
        Args:
            vcard_content: Single vCard entry content
            
        Returns:
            Converted vCard entry
        """
        lines = vcard_content.split('\n')
        converted_lines = []
        
        # Track if we've processed certain fields
        has_version = False
        has_fn = False
        has_n = False
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Handle BEGIN and END
            if line in ["BEGIN:VCARD", "END:VCARD"]:
                converted_lines.append(line)
                continue
            
            # Ensure VERSION is 3.0 for Apple compatibility
            if line.startswith("VERSION:"):
                converted_lines.append("VERSION:3.0")
                has_version = True
                continue
            
            # Ensure FN field exists (required for Apple)
            if line.startswith("FN:"):
                converted_lines.append(line)
                has_fn = True
                continue
            
            # Handle N field (name)
            if line.startswith("N:"):
                converted_lines.append(line)
                has_n = True
                continue
            
            # Handle phone numbers - ensure proper TYPE formatting
            if line.startswith("TEL"):
                converted_line = self._format_phone_line(line)
                converted_lines.append(converted_line)
                continue
            
            # Handle email addresses - ensure proper TYPE formatting
            if line.startswith("EMAIL"):
                converted_line = self._format_email_line(line)
                converted_lines.append(converted_line)
                continue
            
            # Handle address fields
            if line.startswith("ADR"):
                converted_line = self._format_address_line(line)
                converted_lines.append(converted_line)
                continue
            
            # Keep other fields as-is
            converted_lines.append(line)
        
        # Add missing required fields
        if not has_version:
            converted_lines.insert(1, "VERSION:3.0")
        
        # Ensure proper structure
        if not has_fn and has_n:
            # Try to create FN from N field
            n_line = next((line for line in converted_lines if line.startswith("N:")), None)
            if n_line:
                fn_value = self._create_fn_from_n(n_line)
                if fn_value:
                    # Insert FN after N field
                    n_index = converted_lines.index(n_line)
                    converted_lines.insert(n_index + 1, f"FN:{fn_value}")
        
        return '\n'.join(converted_lines)
    
    def _format_phone_number(self, phone_number: str) -> str:
        """Format phone number according to country code rules.

        Args:
            phone_number: Raw phone number

        Returns:
            Formatted phone number
        """
        # Remove any whitespace
        phone_number = phone_number.strip()

        # Rule 1: If starts with 0039, replace with +39
        if phone_number.startswith("0039"):
            return "+39" + phone_number[4:]

        # Rule 2: If starts with +34, keep as is (Spain)
        if phone_number.startswith("+34"):
            return phone_number

        # Rule 3: If starts with 0034, replace with +34 (Spain)
        if phone_number.startswith("0034"):
            return "+34" + phone_number[4:]

        # Rule 4: If starts with other 00XX, convert to +XX
        if phone_number.startswith("00"):
            return "+" + phone_number[2:]

        # Rule 5: If no country code, add +39
        if not phone_number.startswith("+"):
            return "+39" + phone_number

        # If already has other + prefix, keep as is
        return phone_number

    def _format_phone_line(self, line: str) -> str:
        """Format phone line for Apple compatibility.

        Args:
            line: Original phone line

        Returns:
            Formatted phone line
        """
        # Extract phone number and format it
        if ":" in line:
            prefix, phone_number = line.split(":", 1)
            formatted_number = self._format_phone_number(phone_number)
            line = f"{prefix}:{formatted_number}"

        # Ensure proper TYPE formatting
        if "TYPE=" in line and ";" in line:
            # Already properly formatted
            return line
        elif ";" in line and not "TYPE=" in line:
            # Convert old format to new format. Isolate the prefix (before
            # the first ':') from the value first, then extract the type
            # from the prefix alone - not the whole line (dedup: audit issue
            # #64; the old whole-line split embedded the value into the
            # type string for the common single-semicolon shape, see Notes).
            prefix, sep, value = line.partition(":")
            phone_type = extract_type_label(prefix)
            if phone_type:
                return f"TEL;TYPE={phone_type}:{value}"

        return line

    def _format_email_line(self, line: str) -> str:
        """Format email line for Apple compatibility.

        Args:
            line: Original email line

        Returns:
            Formatted email line
        """
        # Ensure proper TYPE formatting
        if "TYPE=" in line and ";" in line:
            return line
        elif ";" in line and not "TYPE=" in line:
            prefix, sep, value = line.partition(":")
            email_type = extract_type_label(prefix)
            if email_type:
                return f"EMAIL;TYPE={email_type}:{value}"

        return line

    def _format_address_line(self, line: str) -> str:
        """Format address line for Apple compatibility.

        Args:
            line: Original address line

        Returns:
            Formatted address line
        """
        # Ensure proper TYPE formatting
        if "TYPE=" in line and ";" in line:
            return line
        elif ";" in line and not "TYPE=" in line:
            prefix, sep, value = line.partition(":")
            addr_type = extract_type_label(prefix)
            if addr_type:
                return f"ADR;TYPE={addr_type}:{value}"

        return line
    
    def _create_fn_from_n(self, n_line: str) -> str:
        """Create FN field from N field.
        
        Args:
            n_line: N field line
            
        Returns:
            Formatted full name
        """
        try:
            # Extract name parts from N field
            name_parts = n_line.split(":")[1].split(";")
            if len(name_parts) >= 2:
                last_name = name_parts[0].strip()
                first_name = name_parts[1].strip()
                if first_name and last_name:
                    return f"{first_name} {last_name}"
                elif first_name:
                    return first_name
                elif last_name:
                    return last_name
        except (IndexError, AttributeError):
            pass
        
        return ""
    
    def validate_vcard_format(self, vcard_content: str) -> bool:
        """Validate vCard format for Apple compatibility.
        
        Args:
            vcard_content: vCard content to validate
            
        Returns:
            True if valid, False otherwise
        """
        for field in self.required_fields:
            if field not in vcard_content:
                logger.warning(f"⚠️ Missing required field: {field}")
                return False
        
        # Check for proper line endings
        if "\r\n" not in vcard_content and "\n" in vcard_content:
            logger.warning("⚠️  Line endings should be CRLF for Apple compatibility")
        
        logger.info("✅ vCard format validation passed")
        return True


class AppleConverterGUI:
    """Tkinter GUI for Apple contact conversion."""
    
    def __init__(self):
        """Initialize the GUI."""
        self.root = tk.Tk()
        self.root.title("Apple Contacts Converter")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.converter = AppleContactConverter()
        self.input_file_path = ""
        self.output_file_path = ""
        
        self._setup_ui()
        
    def _setup_ui(self):
        """Setup the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="🍎 Apple Contacts Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input file selection
        ttk.Label(main_frame, text="Input vCard File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.input_path_var = tk.StringVar()
        input_entry = ttk.Entry(main_frame, textvariable=self.input_path_var, width=50)
        input_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        browse_btn = ttk.Button(main_frame, text="Browse", command=self._browse_input_file)
        browse_btn.grid(row=1, column=2, pady=5)
        
        # Output file display
        ttk.Label(main_frame, text="Output File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(main_frame, textvariable=self.output_path_var, width=50, state="readonly")
        output_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="Convert to Apple Format", 
                                     command=self._convert_file, state="disabled")
        self.convert_btn.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress_var = tk.StringVar(value="Ready to convert")
        progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        progress_label.grid(row=4, column=0, columnspan=3, pady=5)
        
        # Status text
        self.status_text = tk.Text(main_frame, height=10, width=70)
        self.status_text.grid(row=5, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure text widget
        self.status_text.tag_configure("info", foreground="blue")
        self.status_text.tag_configure("success", foreground="green")
        self.status_text.tag_configure("error", foreground="red")
        self.status_text.tag_configure("warning", foreground="orange")
        
        # Initial status
        self._log_status("Ready to convert vCard files to Apple format", "info")
        
    def _browse_input_file(self):
        """Browse for input vCard file."""
        file_path = filedialog.askopenfilename(
            title="Select vCard File",
            filetypes=[("vCard files", "*.vcf"), ("All files", "*.*")]
        )
        
        if file_path:
            self.input_file_path = file_path
            self.input_path_var.set(file_path)
            
            # Generate output path with _apple suffix
            input_path = Path(file_path)
            output_name = f"{input_path.stem}_apple{input_path.suffix}"
            self.output_file_path = str(input_path.parent / output_name)
            self.output_path_var.set(self.output_file_path)
            
            # Enable convert button
            self.convert_btn.config(state="normal")
            
            self._log_status(f"Input file selected: {os.path.basename(file_path)}", "info")
            self._log_status(f"Output will be saved as: {os.path.basename(self.output_file_path)}", "info")
    
    def _convert_file(self):
        """Convert the selected file to Apple format."""
        if not self.input_file_path:
            messagebox.showerror("Error", "Please select an input file first.")
            return
        
        try:
            self.convert_btn.config(state="disabled")
            self.progress_var.set("Converting file...")
            self.root.update()
            
            # Read input file
            self._log_status("Reading input file...", "info")
            with open(self.input_file_path, 'r', encoding='utf-8') as f:
                input_content = f.read()
            
            self._log_status(f"Input file read successfully ({len(input_content)} characters)", "success")
            
            # Convert to Apple format
            self._log_status("Converting to Apple format...", "info")
            converted_content = self.converter.convert_vcard_to_apple_format(input_content)
            
            self._log_status(f"Conversion completed ({len(converted_content)} characters)", "success")
            
            # Validate format
            self._log_status("Validating Apple format...", "info")
            is_valid = self.converter.validate_vcard_format(converted_content)
            
            if is_valid:
                self._log_status("✅ Format validation passed", "success")
            else:
                self._log_status("⚠️  Format validation warnings (file may still work)", "warning")
            
            # Write output file
            self._log_status("Writing output file...", "info")
            with open(self.output_file_path, 'w', encoding='utf-8', newline='\r\n') as f:
                f.write(converted_content)
            
            self._log_status(f"✅ File converted successfully!", "success")
            self._log_status(f"Output saved to: {self.output_file_path}", "success")
            
            # Show success message
            messagebox.showinfo("Success", 
                              f"File converted successfully!\n\n"
                              f"Output saved to:\n{self.output_file_path}")
            
        except Exception as e:
            error_msg = f"Error during conversion: {str(e)}"
            self._log_status(error_msg, "error")
            messagebox.showerror("Conversion Error", error_msg)
            
        finally:
            self.convert_btn.config(state="normal")
            self.progress_var.set("Ready to convert")
    
    def _log_status(self, message: str, level: str = "info"):
        """Log status message to the text widget.
        
        Args:
            message: Message to log
            level: Log level (info, success, error, warning)
        """
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.status_text.insert(tk.END, formatted_message, level)
        self.status_text.see(tk.END)
        self.root.update()
    
    def run(self):
        """Run the GUI application."""
        self.root.mainloop()


def main():
    """Main function to run the Apple converter GUI."""
    try:
        app = AppleConverterGUI()
        app.run()
    except Exception as e:
        print(f"❌ Application failed to start: {e}")
        exit(1)


if __name__ == "__main__":
    main()
