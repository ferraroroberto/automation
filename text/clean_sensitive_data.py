import logging
import re
import json
import os
import sys
from datetime import datetime

log = logging.getLogger(__name__)

# Try to import tkinter, but don't fail if it's not available
try:
    import tkinter as tk
    from tkinter import ttk, scrolledtext
    from tkinter.font import Font
    import pyperclip
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

def has_display():
    """Check if a display is available for GUI"""
    if not TKINTER_AVAILABLE:
        return False
    
    # Check for common display environment variables
    if os.environ.get('DISPLAY') or os.environ.get('WAYLAND_DISPLAY'):
        return True
    
    # On Windows, assume display is available if tkinter is available
    if sys.platform == 'win32':
        return True
    
    # Try to create a simple tkinter window to test
    try:
        test_root = tk.Tk()
        test_root.withdraw()  # Hide the window
        test_root.destroy()
        return True
    except:
        return False

class DataCleanerBase:
    """Base class with common cleaning functionality"""
    
    def __init__(self):
        self.patterns = self.load_patterns()
        # Default cleaning options
        self.clean_paths = True
        self.clean_emails = True
        self.clean_ips = True
        self.clean_urls = True
        self.clean_secrets = True
        self.clean_services = True
        self.clean_numbers = True
        self.clean_guids = True

    def load_patterns(self):
        """Load regex patterns from the JSON file or use defaults if file doesn't exist"""
        patterns_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cleaning_patterns.json")
        
        default_patterns = {
            "paths": r'([A-Za-z]:)?(\\|\/)[^\s:*?"<>|]+',
            "emails": r'[\w\.-]+@[\w\.-]+\.\w+',
            "ips": r'\b\d{1,3}(?:\.\d{1,3}){3}\b',
            "urls": r'https?:\/\/[^\s]+',
            "secrets": r'(?i)(api[_-]?key|token|secret|password|pw|pwd|auth)[\'"\s:=]+[A-Za-z0-9\-_\.]+',
            "services": r'\b(internal_db|prod_db|auth_service|user_service|admin_service|payment_service)\b',
            "numbers": r'\b\d{5,}\b',
            "guids": r'\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b'
        }
        
        try:
            if os.path.exists(patterns_file):
                with open(patterns_file, 'r') as f:
                    return json.load(f)
            else:
                # Create the default file if it doesn't exist
                with open(patterns_file, 'w') as f:
                    json.dump(default_patterns, f, indent=4)
                return default_patterns
        except Exception as e:
            log.warning("Error loading patterns: %s", e)
            return default_patterns

    def clean_sensitive_data(self, text):
        """Clean sensitive data from the text"""
        cleaned = text
        
        if self.clean_paths:
            cleaned = re.sub(self.patterns["paths"], '[REDACTED_PATH]', cleaned)
        
        if self.clean_emails:
            cleaned = re.sub(self.patterns["emails"], '[REDACTED_EMAIL]', cleaned)
        
        if self.clean_ips:
            cleaned = re.sub(self.patterns["ips"], '[REDACTED_IP]', cleaned)
        
        if self.clean_urls:
            cleaned = re.sub(self.patterns["urls"], '[REDACTED_URL]', cleaned)
        
        if self.clean_secrets:
            cleaned = re.sub(self.patterns["secrets"], r'\1=[REDACTED_SECRET]', cleaned)
        
        if self.clean_services:
            cleaned = re.sub(self.patterns["services"], '[REDACTED_SERVICE]', cleaned)
            
        if self.clean_numbers:
            cleaned = re.sub(self.patterns["numbers"], '[REDACTED_NUMBER]', cleaned)
            
        if self.clean_guids:
            cleaned = re.sub(self.patterns["guids"], '[REDACTED_GUID]', cleaned)
            
        return cleaned

class CLIDataCleaner(DataCleanerBase):
    """Command-line interface for data cleaning"""
    
    def __init__(self):
        super().__init__()
    
    def interactive_mode(self):
        """Run in interactive CLI mode"""
        print("=== Sensitive Data Cleaner (CLI Mode) ===")
        print("Enter text to clean (press Ctrl+D or Ctrl+Z when finished):")
        print("Type 'config' to change cleaning options, 'quit' to exit")
        
        while True:
            try:
                print("\n> ", end="")
                command = input().strip().lower()
                
                if command == 'quit':
                    break
                elif command == 'config':
                    self.configure_options()
                    continue
                elif command == 'help':
                    self.show_help()
                    continue
                
                # If it's not a command, treat as text to clean
                if command:
                    # Read multiline input
                    lines = [command]
                    print("Continue entering text (empty line to finish):")
                    while True:
                        try:
                            line = input()
                            if not line:
                                break
                            lines.append(line)
                        except EOFError:
                            break
                    
                    text = '\n'.join(lines)
                    cleaned = self.clean_sensitive_data(text)
                    print("\n--- Cleaned Text ---")
                    print(cleaned)
                    print("--- End ---\n")
                
            except (EOFError, KeyboardInterrupt):
                print("\nExiting...")
                break
    
    def configure_options(self):
        """Configure cleaning options interactively"""
        options = [
            ('clean_paths', 'File Paths'),
            ('clean_emails', 'Email Addresses'),
            ('clean_ips', 'IP Addresses'),
            ('clean_urls', 'URLs'),
            ('clean_secrets', 'Secrets/Keys'),
            ('clean_services', 'Service Names'),
            ('clean_numbers', 'Numeric Sequences'),
            ('clean_guids', 'GUIDs/UUIDs')
        ]
        
        print("\nCurrent cleaning options:")
        for i, (attr, name) in enumerate(options, 1):
            status = "ON" if getattr(self, attr) else "OFF"
            print(f"{i}. {name}: {status}")
        
        print("\nEnter option number to toggle (1-8), or press Enter to finish:")
        while True:
            try:
                choice = input("> ").strip()
                if not choice:
                    break
                
                num = int(choice)
                if 1 <= num <= len(options):
                    attr, name = options[num - 1]
                    current_value = getattr(self, attr)
                    setattr(self, attr, not current_value)
                    new_status = "ON" if not current_value else "OFF"
                    print(f"{name} is now {new_status}")
                else:
                    print("Invalid option number")
            except ValueError:
                print("Please enter a valid number")
            except (EOFError, KeyboardInterrupt):
                break
    
    def show_help(self):
        """Show help information"""
        print("\nAvailable commands:")
        print("  config - Configure cleaning options")
        print("  help   - Show this help")
        print("  quit   - Exit the program")
        print("\nOr enter text to clean it according to current settings")
    
    def clean_file(self, input_path, output_path=None):
        """Clean a file and optionally save to output file"""
        try:
            with open(input_path, 'r', encoding='utf-8') as f:
                text = f.read()
            
            cleaned = self.clean_sensitive_data(text)
            
            if output_path:
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(cleaned)
                log.info("Cleaned text saved to: %s", output_path)
            else:
                print("--- Cleaned Text ---")
                print(cleaned)
                print("--- End ---")
        
        except FileNotFoundError:
            log.error("Error: File not found: %s", input_path)
        except Exception as e:
            log.error("Error processing file: %s", e)

class SensitiveDataCleaner(DataCleanerBase):
    """GUI version of the data cleaner"""
    
    def __init__(self, master):
        super().__init__()
        self.master = master
        self.master.title("Sensitive Data Cleaner")
        self.master.geometry("900x700")
        self.master.minsize(700, 600)
        
        # Set theme
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use 'clam' theme which works well across platforms
        
        # Configure styles
        self.style.configure('TFrame', background='#f5f5f5')
        self.style.configure('TButton', font=('Arial', 10), background='#4a86e8')
        self.style.configure('Copy.TButton', font=('Arial', 10, 'bold'), background='#43a047')
        self.style.configure('TLabel', font=('Arial', 11), background='#f5f5f5')
        self.style.configure('Header.TLabel', font=('Arial', 14, 'bold'), background='#f5f5f5')
        
        # Create main container
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create title
        self.title_label = ttk.Label(self.main_frame, text="Sensitive Data Cleaner", style='Header.TLabel')
        self.title_label.pack(pady=(0, 10))
        
        # Create input frame
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create left panel (input)
        self.left_panel = ttk.Frame(self.input_frame)
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.input_label = ttk.Label(self.left_panel, text="Input Text:")
        self.input_label.pack(anchor=tk.W)
        
        self.input_text = scrolledtext.ScrolledText(self.left_panel, wrap=tk.WORD, width=40, height=20)
        self.input_text.pack(fill=tk.BOTH, expand=True)
        self.input_text.config(font=Font(family="Consolas", size=10))
        
        # Button frame under input
        self.input_buttons = ttk.Frame(self.left_panel)
        self.input_buttons.pack(fill=tk.X, pady=(5, 0))
        
        self.paste_button = ttk.Button(self.input_buttons, text="Paste", command=self.paste_from_clipboard)
        self.paste_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.clear_input_button = ttk.Button(self.input_buttons, text="Clear", command=self.clear_input)
        self.clear_input_button.pack(side=tk.LEFT)
        
        # Create right panel (output)
        self.right_panel = ttk.Frame(self.input_frame)
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.output_label = ttk.Label(self.right_panel, text="Cleaned Text:")
        self.output_label.pack(anchor=tk.W)
        
        self.output_text = scrolledtext.ScrolledText(self.right_panel, wrap=tk.WORD, width=40, height=20)
        self.output_text.pack(fill=tk.BOTH, expand=True)
        self.output_text.config(font=Font(family="Consolas", size=10), state="disabled")
        
        # Button frame under output
        self.output_buttons = ttk.Frame(self.right_panel)
        self.output_buttons.pack(fill=tk.X, pady=(5, 0))
        
        self.copy_button = ttk.Button(self.output_buttons, text="Copy to Clipboard", command=self.copy_to_clipboard, style='Copy.TButton')
        self.copy_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.clear_output_button = ttk.Button(self.output_buttons, text="Clear", command=self.clear_output)
        self.clear_output_button.pack(side=tk.LEFT)
        
        # Middle frame with clean button
        self.middle_frame = ttk.Frame(self.main_frame)
        self.middle_frame.pack(fill=tk.X, pady=10)
        
        self.clean_button = ttk.Button(self.middle_frame, text="Clean Text", command=self.clean_text)
        self.clean_button.pack(pady=5, fill=tk.X)
        
        # Status frame
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill=tk.X)
        
        # Configuration frame
        self.config_frame = ttk.Frame(self.main_frame)
        self.config_frame.pack(fill=tk.X, pady=10)
        
        self.config_label = ttk.Label(self.config_frame, text="Cleaning Options:")
        self.config_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Checkboxes for different cleaning options
        self.options_frame = ttk.Frame(self.config_frame)
        self.options_frame.pack(fill=tk.X)
        
        # Create variables for checkboxes
        self.clean_paths_var = tk.BooleanVar(value=self.clean_paths)
        self.clean_emails_var = tk.BooleanVar(value=self.clean_emails)
        self.clean_ips_var = tk.BooleanVar(value=self.clean_ips)
        self.clean_urls_var = tk.BooleanVar(value=self.clean_urls)
        self.clean_secrets_var = tk.BooleanVar(value=self.clean_secrets)
        self.clean_services_var = tk.BooleanVar(value=self.clean_services)
        self.clean_numbers_var = tk.BooleanVar(value=self.clean_numbers)
        self.clean_guids_var = tk.BooleanVar(value=self.clean_guids)
        
        # First row of options
        self.option1 = ttk.Checkbutton(self.options_frame, text="File Paths", variable=self.clean_paths_var)
        self.option1.grid(row=0, column=0, sticky=tk.W, padx=5)
        
        self.option2 = ttk.Checkbutton(self.options_frame, text="Email Addresses", variable=self.clean_emails_var)
        self.option2.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        self.option3 = ttk.Checkbutton(self.options_frame, text="IP Addresses", variable=self.clean_ips_var)
        self.option3.grid(row=0, column=2, sticky=tk.W, padx=5)
        
        self.option4 = ttk.Checkbutton(self.options_frame, text="URLs", variable=self.clean_urls_var)
        self.option4.grid(row=0, column=3, sticky=tk.W, padx=5)
        
        # Second row of options
        self.option5 = ttk.Checkbutton(self.options_frame, text="Secrets/Keys", variable=self.clean_secrets_var)
        self.option5.grid(row=1, column=0, sticky=tk.W, padx=5)
        
        self.option6 = ttk.Checkbutton(self.options_frame, text="Service Names", variable=self.clean_services_var)
        self.option6.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        self.option7 = ttk.Checkbutton(self.options_frame, text="Numeric Sequences", variable=self.clean_numbers_var)
        self.option7.grid(row=1, column=2, sticky=tk.W, padx=5)
        
        self.option8 = ttk.Checkbutton(self.options_frame, text="GUIDs/UUIDs", variable=self.clean_guids_var)
        self.option8.grid(row=1, column=3, sticky=tk.W, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Check for initial clipboard content
        self.master.after(500, self.check_clipboard)

    def update_cleaning_options(self):
        """Update cleaning options from GUI checkboxes"""
        self.clean_paths = self.clean_paths_var.get()
        self.clean_emails = self.clean_emails_var.get()
        self.clean_ips = self.clean_ips_var.get()
        self.clean_urls = self.clean_urls_var.get()
        self.clean_secrets = self.clean_secrets_var.get()
        self.clean_services = self.clean_services_var.get()
        self.clean_numbers = self.clean_numbers_var.get()
        self.clean_guids = self.clean_guids_var.get()

    def clean_text(self):
        """Clean sensitive data from the input text and show in the output"""
        input_text = self.input_text.get("1.0", tk.END).strip()
        if not input_text:
            self.status_var.set("No input text to clean")
            return
        
        try:
            start_time = datetime.now()
            self.update_cleaning_options()
            cleaned_text = self.clean_sensitive_data(input_text)
            end_time = datetime.now()
            
            # Update output text
            self.output_text.config(state="normal")
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", cleaned_text)
            self.output_text.config(state="disabled")
            
            duration = (end_time - start_time).total_seconds()
            self.status_var.set(f"Text cleaned in {duration:.2f} seconds")
        except Exception as e:
            self.status_var.set(f"Error during cleaning: {str(e)}")

    def copy_to_clipboard(self):
        """Copy the cleaned text to clipboard"""
        cleaned_text = self.output_text.get("1.0", tk.END).strip()
        if cleaned_text:
            pyperclip.copy(cleaned_text)
            self.status_var.set("Cleaned text copied to clipboard")
        else:
            self.status_var.set("No cleaned text to copy")

    def paste_from_clipboard(self):
        """Paste text from clipboard to input"""
        try:
            clipboard_text = pyperclip.paste()
            self.input_text.delete("1.0", tk.END)
            self.input_text.insert("1.0", clipboard_text)
            self.status_var.set("Text pasted from clipboard")
        except Exception as e:
            self.status_var.set(f"Error pasting from clipboard: {str(e)}")

    def clear_input(self):
        """Clear the input text"""
        self.input_text.delete("1.0", tk.END)
        self.status_var.set("Input cleared")

    def clear_output(self):
        """Clear the output text"""
        self.output_text.config(state="normal")
        self.output_text.delete("1.0", tk.END)
        self.output_text.config(state="disabled")
        self.status_var.set("Output cleared")
    
    def check_clipboard(self):
        """Check if clipboard contains text and offer to paste it"""
        try:
            clipboard_text = pyperclip.paste()
            if clipboard_text and not self.input_text.get("1.0", tk.END).strip():
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", clipboard_text)
                self.status_var.set("Text automatically pasted from clipboard")
        except:
            # Ignore clipboard errors
            pass

def main():
    # Parse command line arguments
    if len(sys.argv) > 1:
        # CLI file mode
        cli_cleaner = CLIDataCleaner()
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        cli_cleaner.clean_file(input_file, output_file)
        return
    
    # Check if display is available
    if has_display():
        # GUI mode
        root = tk.Tk()
        app = SensitiveDataCleaner(root)
        root.mainloop()
    else:
        # CLI interactive mode
        cli_cleaner = CLIDataCleaner()
        cli_cleaner.interactive_mode()

if __name__ == "__main__":
    main()
