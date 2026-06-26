#!/usr/bin/env python3
"""
QuickDeck - Minimal always-on-top command launcher
A tiny desktop app that behaves like a mini Stream Deck with configurable buttons.

How to run:
    pip install PySimpleGUI pyautogui pyperclip
    python quickdeck.py [optional path to config]

Features:
- Always-on-top window with optional borderless mode
- Configurable button grid loaded from JSON
- Multiple action types: shell, url, copy, python, reload
- Cross-platform support (Windows, macOS, Linux)
- Auto-creates default config if missing
"""

import json
import logging
import os
import subprocess
import sys
import webbrowser
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List, Optional

import PySimpleGUI as sg
import pyperclip
from PIL import Image, ImageDraw, ImageFont
import io


# Configure logging to console
def setup_logging() -> logging.Logger:
    """Set up console logging with emoji indicators."""
    logger = logging.getLogger("quickdeck")
    logger.setLevel(logging.INFO)

    # Remove existing handlers to avoid duplicates
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)

    # Create console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    # Create formatter with emoji indicators
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )
    console_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    return logger


logger = setup_logging()


class QuickDeckConfig:
    """Configuration GUI for QuickDeck."""
    
    def __init__(self, quickdeck_instance):
        """Initialize configuration GUI."""
        self.qd = quickdeck_instance
        self.config_window = None
    
    def show_config_window(self) -> None:
        """Show the main configuration window."""
        grid_config = self.qd.config.get("grid", {})
        rows = grid_config.get("rows", 2)
        cols = grid_config.get("cols", 4)
        total_buttons = rows * cols
        
        # Parse geometry for window size settings
        geometry = self.qd.config.get("geometry", "420x300+40+40")
        try:
            geom_parts = geometry.split('+')
            size = geom_parts[0].split('x')
            config_width, config_height = int(size[0]), int(size[1])
        except (ValueError, IndexError):
            config_width, config_height = 420, 300
        
        # Get theme settings
        theme_config = self.qd.config.get("theme", {})
        
        # Main configuration layout
        layout = [
            [sg.Text("QuickDeck Configuration", font=("Arial", 16, "bold"))],
            [sg.HSeparator()],
            [sg.Text("General Settings:", font=("Arial", 12, "bold"))],
            [sg.Text("Title:"), sg.Input(self.qd.config.get("title", "🚀 QuickDeck"), key="title", size=(30, 1))],
            [sg.Checkbox("Always on top", key="topmost", default=self.qd.config.get("topmost", True)),
             sg.Checkbox("Borderless window", key="borderless", default=self.qd.config.get("borderless", False))],
            [sg.HSeparator()],
            [sg.Text("Window Settings:", font=("Arial", 12, "bold"))],
            [sg.Text("Width:"), sg.Input(str(config_width), key="window_width", size=(8, 1)),
             sg.Text("Height:"), sg.Input(str(config_height), key="window_height", size=(8, 1))],
            [sg.HSeparator()],
            [sg.Text("Grid Settings:", font=("Arial", 12, "bold"))],
            [sg.Text("Rows:"), sg.Input(str(rows), key="grid_rows", size=(5, 1)),
             sg.Text("Cols:"), sg.Input(str(cols), key="grid_cols", size=(5, 1))],
            [sg.Text("Emoji Size:"), sg.Input(str(grid_config.get("emoji_size", 48)), key="emoji_size", size=(5, 1)),
             sg.Text("Text Size:"), sg.Input(str(grid_config.get("text_size", 12)), key="text_size", size=(5, 1))],
            [sg.Text("Text Position:"), sg.Combo(["below", "above", "left", "right"], 
                                                 default_value=grid_config.get("text_position", "below"), 
                                                 key="text_position", size=(10, 1))],
            [sg.Text("Padding:"), sg.Input(str(grid_config.get("pad", 8)), key="grid_pad", size=(5, 1)),
             sg.Text("Button Height:"), sg.Input(str(grid_config.get("button_height", 80)), key="button_height", size=(5, 1))],
            [sg.HSeparator()],
            [sg.Text("Theme Colors:", font=("Arial", 12, "bold"))],
            [sg.Text("Background:"), sg.Input(theme_config.get("bg", "#111827"), key="theme_bg", size=(10, 1)),
             sg.Text("Text:"), sg.Input(theme_config.get("fg", "#E5E7EB"), key="theme_fg", size=(10, 1))],
            [sg.Text("Button BG:"), sg.Input(theme_config.get("button_bg", "#1F2937"), key="theme_button_bg", size=(10, 1))],
            [sg.HSeparator()],
            [sg.Text("Button Management:", font=("Arial", 12, "bold"))],
            [sg.Text(f"Configure individual buttons one by one ({total_buttons} total):")],
        ]
        
        # Create button edit grid matching the main app layout (rows x cols)
        buttons = self.qd.config.get("buttons", [])
        for row in range(rows):
            button_row = []
            for col in range(cols):
                btn_index = row * cols + col
                if btn_index < len(buttons) and buttons[btn_index]:
                    # Show actual button name if configured
                    button_config = buttons[btn_index]
                    label = button_config.get("label", f"Button {btn_index + 1}")
                    button_text = f"Edit {label}"
                else:
                    # Show generic name for empty buttons
                    button_text = f"Edit Button {btn_index + 1}"
                
                button_row.append(sg.Button(button_text, key=f"edit_btn_{btn_index}", size=(12, 1)))
            layout.append(button_row)
        
        layout.extend([
            [sg.HSeparator()],
            [sg.Button("Save Configuration", key="save_config", button_color=("white", "green")), 
             sg.Button("Cancel", key="cancel", button_color=("white", "red"))]
        ])
        
        # Calculate window size to accommodate the grid properly
        # Width: enough for cols buttons + padding
        window_width = max(cols * 120 + 100, 700)  # 120px per button + padding, increased min width
        # Height: header + settings + button rows + footer + extra space for visibility
        base_height = 600  # Significantly increased base height for all sections and button visibility
        button_row_height = 45  # Increased height per button row for better spacing
        window_height = base_height + (rows * button_row_height)
        
        self.config_window = sg.Window(
            "QuickDeck Configuration",
            layout,
            size=(window_width, min(window_height, 800)),  # Cap at 800px height
            resizable=True,
            modal=True
        )
        
        # Event loop for configuration window
        while True:
            event, values = self.config_window.read()
            
            if event == sg.WIN_CLOSED or event == "cancel":
                break
            elif event == "save_config":
                self._save_grid_configuration(values)
                break
            elif event.startswith("edit_btn_"):
                btn_index = int(event.split("_")[2])
                self._edit_single_button(btn_index)
        
        if self.config_window:
            self.config_window.close()
    
    def _edit_single_button(self, btn_index: int) -> None:
        """Edit a single button with detailed configuration."""
        buttons = self.qd.config.get("buttons", [])
        grid_config = self.qd.config.get("grid", {})
        total_buttons = grid_config.get("rows", 2) * grid_config.get("cols", 4)
        
        if btn_index >= total_buttons:
            sg.popup_error(f"Button {btn_index + 1} is outside the grid range!", title="Error")
            return
        
        # Get existing button config or create new one
        if btn_index < len(buttons):
            btn_config = buttons[btn_index]
            label = btn_config.get("label", f"Button {btn_index + 1}")
            emoji = btn_config.get("emoji", "")
            btn_type = btn_config.get("type", "shell")
            hint = btn_config.get("hint", label)
        else:
            # Create empty button config for new buttons
            btn_config = {}
            label = f"Button {btn_index + 1}"
            emoji = ""
            btn_type = "shell"
            hint = label
        
        # Create button edit layout
        layout = [
            [sg.Text(f"Edit Button {btn_index + 1}", font=("Arial", 14, "bold"))],
            [sg.HSeparator()],
            [sg.Text("Label:"), sg.Input(label, key="label", size=(20, 1))],
            [sg.Text("Emoji:"), sg.Input(emoji, key="emoji", size=(10, 1))],
            [sg.Text("Tooltip:"), sg.Input(hint, key="hint", size=(30, 1))],
            [sg.Text("Type:"), sg.Combo(["shell", "url", "copy", "python"], 
                                        default_value=btn_type, 
                                        key="type", 
                                        size=(15, 1))],
            [sg.HSeparator()],
            [sg.Text("Action Configuration:", font=("Arial", 12, "bold"))],
            [sg.Text("Action:"), sg.Input(self._get_action_value(btn_config, btn_type), 
                                          key="action", size=(40, 1))],
            [sg.HSeparator()],
            [sg.Button("Save Button", key="save_button", button_color=("white", "green")), 
             sg.Button("Delete Button", key="delete_button", button_color=("white", "red")),
             sg.Button("Cancel", key="cancel")]
        ]
        
        edit_window = sg.Window(f"Edit Button {btn_index + 1}", layout, modal=True)
        
        while True:
            event, values = edit_window.read()
            if event in (sg.WIN_CLOSED, "cancel"):
                break
            elif event == "save_button":
                self._save_button_config(btn_index, values)
                break
            elif event == "delete_button":
                self._delete_button_config(btn_index)
                break
        
        edit_window.close()
    
    def _get_action_value(self, btn_config: Dict[str, Any], btn_type: str) -> str:
        """Get the action value for display in the config GUI."""
        if btn_type == "url":
            return btn_config.get("url", "")
        elif btn_type == "copy":
            return btn_config.get("text", "")
        elif btn_type == "python":
            return btn_config.get("code", "")
        elif btn_type == "shell":
            cmd = btn_config.get("cmd", "")
            args = btn_config.get("args", [])
            if args:
                return f"{cmd} {' '.join(args)}"
            return cmd
        return ""
    
    def _save_button_config(self, btn_index: int, values: Dict[str, Any]) -> None:
        """Save a single button configuration."""
        try:
            label = values.get("label", "").strip()
            emoji = values.get("emoji", "").strip()
            hint = values.get("hint", "").strip()
            btn_type = values.get("type", "shell")
            action = values.get("action", "").strip()
            
            if not label:
                # Use default label instead of blocking with error
                label = f"Button {btn_index + 1}"
            
            # Create button configuration
            button_config = {
                "label": label,
                "emoji": emoji,
                "hint": hint,
                "type": btn_type
            }
            
            # Add type-specific configuration
            if btn_type == "url":
                if not action.startswith(("http://", "https://")):
                    # Provide default URL instead of blocking with error
                    button_config["url"] = "https://www.google.com"
                else:
                    button_config["url"] = action
            elif btn_type == "copy":
                button_config["text"] = action
            elif btn_type == "python":
                button_config["code"] = action
            elif btn_type == "shell":
                parts = action.split()
                if parts:
                    button_config["cmd"] = parts[0]
                    if len(parts) > 1:
                        button_config["args"] = parts[1:]
                else:
                    # Provide a default command instead of blocking with error
                    button_config["cmd"] = "echo"
                    button_config["args"] = ["Button configured - add your command"]
            
            # Update buttons list
            buttons = self.qd.config.get("buttons", [])
            
            # Ensure buttons list is long enough
            while len(buttons) <= btn_index:
                buttons.append({})
            
            buttons[btn_index] = button_config
            
            # Save to file
            self.qd.config["buttons"] = buttons
            with open(self.qd.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.qd.config, f, indent=2, ensure_ascii=False)
            
            sg.popup(f"✅ Button {btn_index + 1} saved successfully!", title="Success")
            
        except Exception as e:
            sg.popup_error(f"❌ Failed to save button: {e}", title="Error")
    
    def _delete_button_config(self, btn_index: int) -> None:
        """Delete a button configuration."""
        try:
            buttons = self.qd.config.get("buttons", [])
            
            if btn_index < len(buttons):
                del buttons[btn_index]
                
                # Save to file
                self.qd.config["buttons"] = buttons
                with open(self.qd.config_path, 'w', encoding='utf-8') as f:
                    json.dump(self.qd.config, f, indent=2, ensure_ascii=False)
                
                sg.popup(f"✅ Button {btn_index + 1} deleted successfully!", title="Success")
            else:
                sg.popup("Button is already empty!", title="Info")
                
        except Exception as e:
            sg.popup_error(f"❌ Failed to delete button: {e}", title="Error")
    
    def _save_grid_configuration(self, values: Dict[str, Any]) -> None:
        """Save the grid configuration."""
        try:
            # Update general settings
            self.qd.config["title"] = values.get("title", "🚀 QuickDeck").strip()
            self.qd.config["topmost"] = values.get("topmost", True)
            self.qd.config["borderless"] = values.get("borderless", False)
            
            # Update grid settings
            grid_config = self.qd.config.get("grid", {})
            grid_config["rows"] = int(values.get("grid_rows", 2))
            grid_config["cols"] = int(values.get("grid_cols", 4))
            grid_config["emoji_size"] = int(values.get("emoji_size", 48))
            grid_config["text_size"] = int(values.get("text_size", 12))
            grid_config["text_position"] = values.get("text_position", "below")
            grid_config["pad"] = int(values.get("grid_pad", 8))
            grid_config["button_height"] = int(values.get("button_height", 80))
            
            # Update theme settings
            theme_config = self.qd.config.get("theme", {})
            theme_config["bg"] = values.get("theme_bg", "#111827").strip()
            theme_config["fg"] = values.get("theme_fg", "#E5E7EB").strip()
            theme_config["button_bg"] = values.get("theme_button_bg", "#1F2937").strip()
            
            # Update window size settings
            window_width = int(values.get("window_width", 420))
            window_height = int(values.get("window_height", 300))
            
            # Validate grid settings
            if grid_config["rows"] < 1 or grid_config["cols"] < 1:
                sg.popup_error("Rows and columns must be at least 1!", title="Error")
                return
            
            if grid_config["emoji_size"] < 16 or grid_config["emoji_size"] > 128:
                sg.popup_error("Emoji size must be between 16 and 128!", title="Error")
                return
            
            if grid_config["text_size"] < 8 or grid_config["text_size"] > 24:
                sg.popup_error("Text size must be between 8 and 24!", title="Error")
                return
            
            if grid_config["pad"] < 0 or grid_config["pad"] > 50:
                sg.popup_error("Padding must be between 0 and 50!", title="Error")
                return
            
            if grid_config["button_height"] < 40 or grid_config["button_height"] > 200:
                sg.popup_error("Button height must be between 40 and 200!", title="Error")
                return
            
            # Validate window size settings
            if window_width < 200 or window_width > 2000:
                sg.popup_error("Window width must be between 200 and 2000!", title="Error")
                return
            
            if window_height < 150 or window_height > 1500:
                sg.popup_error("Window height must be between 150 and 1500!", title="Error")
                return
            
            # Validate color format (basic hex color validation)
            import re
            color_pattern = r'^#[0-9A-Fa-f]{6}$'
            if not re.match(color_pattern, theme_config["bg"]):
                sg.popup_error("Background color must be a valid hex color (e.g., #111827)!", title="Error")
                return
            
            if not re.match(color_pattern, theme_config["fg"]):
                sg.popup_error("Text color must be a valid hex color (e.g., #E5E7EB)!", title="Error")
                return
            
            if not re.match(color_pattern, theme_config["button_bg"]):
                sg.popup_error("Button background color must be a valid hex color (e.g., #1F2937)!", title="Error")
                return
            
            # Update geometry with new window size
            current_geometry = self.qd.config.get("geometry", "420x300+40+40")
            try:
                geom_parts = current_geometry.split('+')
                if len(geom_parts) >= 3:
                    x, y = int(geom_parts[1]), int(geom_parts[2])
                else:
                    x, y = 40, 40
            except (ValueError, IndexError):
                x, y = 40, 40
            
            new_geometry = f"{window_width}x{window_height}+{x}+{y}"
            self.qd.config["geometry"] = new_geometry
            
            # Update configuration
            self.qd.config["grid"] = grid_config
            self.qd.config["theme"] = theme_config
            
            # Save to file
            with open(self.qd.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.qd.config, f, indent=2, ensure_ascii=False)
            
            # Reload the main window
            self.qd.execute_reload_action()
            
            sg.popup("✅ Configuration saved and applied!", title="Success")
            
        except ValueError as e:
            sg.popup_error(f"❌ Invalid input values: {e}", title="Error")
        except Exception as e:
            sg.popup_error(f"❌ Failed to save configuration: {e}", title="Error")


class QuickDeck:
    """Main QuickDeck application class."""
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize QuickDeck with config file."""
        self.config_path = Path(config_path) if config_path else Path("quickdeck.json")
        self.config: Dict[str, Any] = {}
        self.window: Optional[sg.Window] = None
        self.status_msg = "🚀 QuickDeck ready"
        self.emoji_cache: Dict[str, bytes] = {}  # Cache for emoji images
        self.config_gui = None  # Will be initialized after config is loaded
        
        # Load or create config
        self.load_config()
        
        # Set PySimpleGUI theme
        self.setup_theme()
        
        # Initialize configuration GUI
        self.config_gui = QuickDeckConfig(self)
    
    def setup_theme(self) -> None:
        """Configure PySimpleGUI theme based on config."""
        theme_config = self.config.get("theme", {})
        bg_color = theme_config.get("bg", "#111827")
        fg_color = theme_config.get("fg", "#E5E7EB")
        button_bg = theme_config.get("button_bg", "#1F2937")
        
        # Create custom theme
        sg.theme_add_new('QuickDeckDark', {
            'BACKGROUND': bg_color,
            'TEXT': fg_color,
            'INPUT': button_bg,
            'TEXT_INPUT': fg_color,
            'SCROLL': button_bg,
            'BUTTON': (fg_color, button_bg),
            'PROGRESS': (fg_color, button_bg),
            'BORDER': 1,
            'SLIDER_DEPTH': 0,
            'PROGRESS_DEPTH': 0,
        })
        sg.theme('QuickDeckDark')
    
    def load_config(self) -> None:
        """Load configuration from JSON file or create default."""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                logger.info("✅ Configuration loaded successfully")
            else:
                self.create_default_config()
                logger.info("📝 Created default configuration")
        except Exception as e:
            logger.error(f"❌ Config loading failed: {e}")
            self.create_default_config()
    
    def create_default_config(self) -> None:
        """Create default configuration file."""
        default_config = {
            "title": "🚀 QuickDeck",
            "geometry": "420x300+40+40",
            "topmost": True,
            "borderless": False,
            "theme": {
                "bg": "#111827",
                "fg": "#E5E7EB", 
                "button_bg": "#1F2937"
            },
            "grid": {
                "rows": 2,
                "cols": 4,
                "pad": 8,
                "button_height": 80,
                "emoji_size": 48,
                "text_size": 12,
                "text_position": "below"
            },
            "buttons": [
                {
                    "label": "Browser",
                    "emoji": "🌐",
                    "type": "url",
                    "url": "https://www.google.com"
                },
                {
                    "label": "Explorer", 
                    "emoji": "📂",
                    "type": "shell",
                    "cmd": "explorer" if os.name == 'nt' else "nautilus",
                    "args": ["~"]
                },
                {
                    "label": "Copy email",
                    "emoji": "📧", 
                    "type": "copy",
                    "text": "user@example.com"
                },
                {
                    "label": "Ping DNS",
                    "emoji": "📡",
                    "type": "shell",
                    "cmd": "ping",
                    "args": ["-n", "4", "8.8.8.8"] if os.name == 'nt' else ["-c", "4", "8.8.8.8"]
                },
                {
                    "label": "Snippet",
                    "emoji": "🐍",
                    "type": "python", 
                    "code": "print('Hello from QuickDeck!')"
                }
            ]
        }
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            self.config = default_config
        except Exception as e:
            logger.error(f"❌ Failed to create default config: {e}")
            self.config = default_config
    
    def expand_path(self, path: str) -> str:
        """Expand ~ and environment variables in path."""
        # Expand ~
        if path.startswith('~'):
            path = os.path.expanduser(path)
        
        # Expand environment variables
        if os.name == 'nt':  # Windows
            path = os.path.expandvars(path)
        else:  # Unix-like
            import os
            path = os.path.expandvars(path)
        
        return path
    
    def create_emoji_image(self, emoji: str, size: int = 32) -> bytes:
        """Create a colored emoji image using PIL."""
        cache_key = f"{emoji}_{size}"
        if cache_key in self.emoji_cache:
            return self.emoji_cache[cache_key]
        
        try:
            # Create image with extra padding to prevent cutting
            padding = 8  # Increased padding to prevent complex emoji cutting
            img_size = size + (padding * 2)
            img = Image.new('RGBA', (img_size, img_size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # Try to use a system emoji font
            try:
                # Windows emoji font
                if os.name == 'nt':
                    font_paths = [
                        "C:/Windows/Fonts/seguiemj.ttf",  # Segoe UI Emoji
                        "C:/Windows/Fonts/segmdl2.ttf",  # Segoe MDL2 Assets
                    ]
                else:
                    font_paths = [
                        "/System/Library/Fonts/Apple Color Emoji.ttc",  # macOS
                        "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",  # Linux
                    ]
                
                font = None
                for font_path in font_paths:
                    if os.path.exists(font_path):
                        try:
                            font = ImageFont.truetype(font_path, size)
                            break
                        except OSError:
                            continue

                if font is None:
                    # Fallback to default font
                    font = ImageFont.load_default()

            except Exception:
                font = ImageFont.load_default()
            
            # Calculate text position to center the emoji with padding
            bbox = draw.textbbox((0, 0), emoji, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            # Center the emoji in the padded image with extra vertical padding
            x = (img_size - text_width) // 2
            y = (img_size - text_height) // 2 + 2  # Slight vertical adjustment for better centering
            
            # Draw the emoji
            draw.text((x, y), emoji, font=font, embedded_color=True)
            
            # Convert to bytes
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            # Cache the result
            self.emoji_cache[cache_key] = img_bytes.getvalue()
            return img_bytes.getvalue()
            
        except Exception as e:
            logger.warning(f"⚠️ Failed to create emoji image for {emoji}: {e}")
            # Return a simple colored square as fallback
            img = Image.new('RGBA', (size, size), (255, 255, 255, 255))
            draw = ImageDraw.Draw(img)
            draw.rectangle([0, 0, size-1, size-1], outline=(100, 100, 100, 255))
            
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            return img_bytes.getvalue()
    
    def _create_combined_button_image(self, emoji_data: bytes, text: str, emoji_size: int, text_size: int, text_color: str, bg_color: str) -> bytes:
        """Create a combined image with emoji on top and text below."""
        try:
            # Calculate dimensions with proper text width calculation
            padding = 10
            text_height = text_size + 10
            
            # Calculate text width first to determine total width needed
            try:
                if os.name == 'nt':
                    font_paths = [
                        "C:/Windows/Fonts/arial.ttf",
                        "C:/Windows/Fonts/calibri.ttf",
                    ]
                else:
                    font_paths = [
                        "/System/Library/Fonts/Arial.ttf",  # macOS
                        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
                    ]
                
                font = None
                for font_path in font_paths:
                    if os.path.exists(font_path):
                        try:
                            font = ImageFont.truetype(font_path, text_size)
                            break
                        except OSError:
                            continue

                if font is None:
                    font = ImageFont.load_default()

                # Calculate actual text width
                temp_img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
                temp_draw = ImageDraw.Draw(temp_img)
                bbox = temp_draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]

            except Exception:
                # Fallback: estimate text width (rough approximation)
                text_width = len(text) * (text_size * 0.6)
            
            # Calculate total dimensions - use the larger of emoji width or text width
            # Account for emoji padding (emoji image is actually larger due to padding)
            padded_emoji_size = emoji_size + 16  # 8px padding on each side
            total_width = max(padded_emoji_size, text_width) + (padding * 2)
            total_height = padded_emoji_size + text_height + (padding * 2)
            
            # Create image with background color
            img = Image.new('RGBA', (total_width, total_height), bg_color)
            draw = ImageDraw.Draw(img)
            
            # Load emoji image (it already has padding)
            emoji_img = Image.open(io.BytesIO(emoji_data))
            # Don't resize - use the emoji at its original padded size
            
            # Paste emoji at the top, centered
            emoji_x = (total_width - padded_emoji_size) // 2
            emoji_y = padding
            img.paste(emoji_img, (emoji_x, emoji_y), emoji_img)
            
            # Add text below emoji
            try:
                # Try to use a system font
                if os.name == 'nt':
                    font_paths = [
                        "C:/Windows/Fonts/arial.ttf",
                        "C:/Windows/Fonts/calibri.ttf",
                    ]
                else:
                    font_paths = [
                        "/System/Library/Fonts/Arial.ttf",  # macOS
                        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
                    ]
                
                font = None
                for font_path in font_paths:
                    if os.path.exists(font_path):
                        try:
                            font = ImageFont.truetype(font_path, text_size)
                            break
                        except OSError:
                            continue

                if font is None:
                    font = ImageFont.load_default()

            except Exception:
                font = ImageFont.load_default()
            
            # Calculate text position
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_x = (total_width - text_width) // 2
            text_y = emoji_y + padded_emoji_size + 5
            
            # Draw text
            draw.text((text_x, text_y), text, font=font, fill=text_color)
            
            # Convert to bytes
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            return img_bytes.getvalue()
            
        except Exception as e:
            logger.warning(f"⚠️ Failed to create combined button image: {e}")
            # Return the original emoji data as fallback
            return emoji_data
    
    def set_status(self, msg: str) -> None:
        """Update status message."""
        self.status_msg = msg
        logger.info(msg)
    
    def create_window(self) -> sg.Window:
        """Create the main window with button grid."""
        grid_config = self.config.get("grid", {"rows": 2, "cols": 4, "pad": 8, "button_height": 80, "emoji_size": 48, "text_size": 12, "text_position": "below"})
        buttons = self.config.get("buttons", [])
        emoji_size = grid_config.get("emoji_size", 48)
        text_size = grid_config.get("text_size", 12)
        text_position = grid_config.get("text_position", "below")
        
        # Create fixed grid based on rows x cols
        rows = grid_config.get("rows", 2)
        cols = grid_config.get("cols", 4)
        total_buttons = rows * cols
        
        # Create button grid
        button_rows = []
        
        for row in range(rows):
            button_row = []
            for col in range(cols):
                btn_index = row * cols + col
                
                if btn_index < len(buttons):
                    button_config = buttons[btn_index]
                    label = button_config.get("label", f"Button {btn_index+1}")
                    emoji = button_config.get("emoji", "")
                    hint = button_config.get("hint", label)
                else:
                    label = f"Button {btn_index+1}"
                    emoji = ""
                    hint = label
                
                # Create button with emoji and text (always text below emoji like Elgato)
                if emoji:
                    # Create emoji image
                    emoji_data = self.create_emoji_image(emoji, emoji_size)
                    
                    # Create button with image and text (text always below emoji)
                    theme_config = self.config.get("theme", {})
                    button_bg = theme_config.get("button_bg", "#1F2937")
                    button_fg = theme_config.get("fg", "#E5E7EB")
                    
                    # Create a combined image with emoji on top and text below
                    combined_image_data = self._create_combined_button_image(
                        emoji_data, label, emoji_size, text_size, button_fg, button_bg
                    )
                    
                    # Calculate proper image size based on text width
                    try:
                        # Calculate text width for proper sizing
                        if os.name == 'nt':
                            font_paths = ["C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/calibri.ttf"]
                        else:
                            font_paths = ["/System/Library/Fonts/Arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
                        
                        font = None
                        for font_path in font_paths:
                            if os.path.exists(font_path):
                                try:
                                    font = ImageFont.truetype(font_path, text_size)
                                    break
                                except OSError:
                                    continue

                        if font is None:
                            font = ImageFont.load_default()

                        temp_img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
                        temp_draw = ImageDraw.Draw(temp_img)
                        bbox = temp_draw.textbbox((0, 0), label, font=font)
                        text_width = bbox[2] - bbox[0]

                    except Exception:
                        text_width = len(label) * (text_size * 0.6)
                    
                    # Use proper sizing - width should accommodate text, height for emoji + text
                    # Account for emoji padding (8px on each side = 16px total)
                    padded_emoji_size = emoji_size + 16
                    image_width = max(padded_emoji_size, text_width) + 20
                    image_height = padded_emoji_size + text_size + 20
                    
                    btn = sg.Button(
                        "",
                        key=f"btn_{btn_index}",
                        size=(None, None),
                        tooltip=hint,
                        expand_x=True,
                        expand_y=True,
                        image_data=combined_image_data,
                        image_size=(image_width, image_height),
                        image_subsample=1,
                        button_color=(button_fg, button_bg)
                    )
                else:
                    # Create button without emoji
                    btn = sg.Button(
                        label,
                        key=f"btn_{btn_index}",
                        size=(None, None),  # Let PySimpleGUI auto-size based on content
                        tooltip=hint,
                        expand_x=True,
                        expand_y=True,
                        font=("Arial", text_size)
                    )
                
                button_row.append(btn)
            button_rows.append(button_row)
        
        # Create layout
        layout = [
            [sg.Text(self.config.get("title", "🚀 QuickDeck"), 
                    font=("Arial", 16, "bold"), 
                    justification="center", 
                    expand_x=True)],
            [sg.HSeparator()],
            *button_rows,
            [sg.HSeparator()],
            [sg.Button("⚙️ Config", key="config", size=(10, 1)), 
             sg.Text("", key="status", size=(None, 1), expand_x=True),
             sg.Button("🔄", key="reload", size=(3, 1))]
        ]
        
        # Calculate proper window size based on grid
        # Account for emoji padding (8px on each side = 16px total)
        padded_emoji_size = emoji_size + 16
        button_height = max(padded_emoji_size + text_size + 20, 60)  # Minimum button height
        
        # Calculate button width based on longest text in buttons
        max_text_width = 0
        for btn_index in range(min(len(buttons), total_buttons)):
            if btn_index < len(buttons) and buttons[btn_index]:
                button_config = buttons[btn_index]
                label = button_config.get("label", f"Button {btn_index+1}")
                # Estimate text width (rough approximation)
                estimated_width = len(label) * (text_size * 0.6)
                max_text_width = max(max_text_width, estimated_width)
        
        # Use the larger of padded emoji size or max text width, plus padding
        button_width = max(padded_emoji_size, max_text_width) + 20
        
        # Calculate window dimensions
        window_width = max(cols * button_width + 40, 300)  # Add padding
        window_height = max(rows * button_height + 120, 200)  # Add space for title, separators, config button
        
        # Window properties
        geometry = self.config.get("geometry", f"{window_width}x{window_height}+40+40")
        topmost = self.config.get("topmost", True)
        borderless = self.config.get("borderless", False)
        
        # Parse geometry, but use calculated size if it's too small
        try:
            geom_parts = geometry.split('+')
            size = geom_parts[0].split('x')
            config_width, config_height = int(size[0]), int(size[1])
            if len(geom_parts) > 2:
                x, y = int(geom_parts[1]), int(geom_parts[2])
            else:
                x, y = 40, 40
            
            # Use the larger of configured or calculated size
            width = max(config_width, window_width)
            height = max(config_height, window_height)
        except (ValueError, IndexError):
            width, height, x, y = window_width, window_height, 40, 40
        
        # Create window
        window = sg.Window(
            self.config.get("title", "🚀 QuickDeck"),
            layout,
            size=(width, height),
            location=(x, y),
            keep_on_top=topmost,
            no_titlebar=borderless,
            grab_anywhere=borderless,
            resizable=True,
            finalize=True
        )
        
        return window
    
    def execute_shell_action(self, button_config: Dict[str, Any]) -> None:
        """Execute shell command action."""
        try:
            cmd = button_config.get("cmd", "")
            args = button_config.get("args", [])
            cwd = button_config.get("cwd", ".")
            
            # Expand paths and environment variables
            cmd = self.expand_path(cmd)
            args = [self.expand_path(arg) for arg in args]
            cwd = self.expand_path(cwd)
            
            # Prepare command
            if args:
                full_cmd = [cmd] + args
            else:
                full_cmd = cmd
            
            self.set_status(f"🔄 Running: {cmd}")
            
            # Execute command
            process = subprocess.Popen(
                full_cmd,
                cwd=cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Don't wait for completion to avoid blocking UI
            self.set_status(f"✅ Started: {cmd}")
            
        except Exception as e:
            error_msg = f"❌ Shell command failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="Shell Command Error")
    
    def execute_url_action(self, button_config: Dict[str, Any]) -> None:
        """Execute URL action."""
        try:
            url = button_config.get("url", "")
            if not url:
                raise ValueError("No URL specified")
            
            self.set_status(f"🌐 Opening: {url}")
            webbrowser.open(url)
            self.set_status(f"✅ Opened: {url}")
            
        except Exception as e:
            error_msg = f"❌ URL action failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="URL Error")
    
    def execute_copy_action(self, button_config: Dict[str, Any]) -> None:
        """Execute copy to clipboard action."""
        try:
            text = button_config.get("text", "")
            if not text:
                raise ValueError("No text specified")
            
            pyperclip.copy(text)
            self.set_status(f"📋 Copied: {text[:30]}{'...' if len(text) > 30 else ''}")
            
        except Exception as e:
            error_msg = f"❌ Copy action failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="Copy Error")
    
    def execute_python_action(self, button_config: Dict[str, Any]) -> None:
        """Execute Python code snippet safely."""
        try:
            code = button_config.get("code", "")
            if not code:
                raise ValueError("No Python code specified")
            
            self.set_status("🐍 Executing Python code...")
            
            # Capture stdout and stderr
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            stdout_capture = StringIO()
            stderr_capture = StringIO()
            
            try:
                sys.stdout = stdout_capture
                sys.stderr = stderr_capture
                
                # Execute in restricted namespace
                namespace = {
                    '__builtins__': {
                        'print': print,
                        'len': len,
                        'str': str,
                        'int': int,
                        'float': float,
                        'list': list,
                        'dict': dict,
                        'tuple': tuple,
                        'set': set,
                        'range': range,
                        'enumerate': enumerate,
                        'zip': zip,
                        'min': min,
                        'max': max,
                        'sum': sum,
                        'abs': abs,
                        'round': round,
                        'sorted': sorted,
                        'reversed': reversed,
                    }
                }
                
                exec(code, namespace)
                
            finally:
                sys.stdout = old_stdout
                sys.stderr = old_stderr
            
            # Get output
            stdout_output = stdout_capture.getvalue()
            stderr_output = stderr_capture.getvalue()
            
            # Show results
            if stdout_output or stderr_output:
                result_text = ""
                if stdout_output:
                    result_text += f"Output:\n{stdout_output}\n"
                if stderr_output:
                    result_text += f"Errors:\n{stderr_output}\n"
                
                sg.popup_scrolled(
                    result_text,
                    title="Python Execution Result",
                    size=(50, 20)
                )
            
            self.set_status("✅ Python code executed successfully")
            
        except Exception as e:
            error_msg = f"❌ Python execution failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="Python Execution Error")
    
    def execute_reload_action(self) -> None:
        """Reload configuration and rebuild window."""
        try:
            self.set_status("🔄 Reloading configuration...")
            self.load_config()
            self.setup_theme()
            
            # Close current window
            if self.window:
                self.window.close()
            
            # Create new window
            self.window = self.create_window()
            self.set_status("✅ Configuration reloaded")
            
        except Exception as e:
            error_msg = f"❌ Reload failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="Reload Error")
    
    def handle_button_click(self, event: str) -> None:
        """Handle button click events."""
        try:
            # Extract button index
            if event.startswith("btn_"):
                btn_index = int(event.split("_")[1])
                buttons = self.config.get("buttons", [])
                
                if btn_index < len(buttons):
                    button_config = buttons[btn_index]
                    action_type = button_config.get("type", "shell")
                    
                    # Execute appropriate action
                    if action_type == "shell":
                        self.execute_shell_action(button_config)
                    elif action_type == "url":
                        self.execute_url_action(button_config)
                    elif action_type == "copy":
                        self.execute_copy_action(button_config)
                    elif action_type == "python":
                        self.execute_python_action(button_config)
                    elif action_type == "reload":
                        self.execute_reload_action()
                    else:
                        self.set_status(f"❌ Unknown action type: {action_type}")
                        
        except Exception as e:
            error_msg = f"❌ Button action failed: {e}"
            self.set_status(error_msg)
            sg.popup_error(error_msg, title="Button Action Error")
    
    def run(self) -> None:
        """Main application loop."""
        try:
            # Create initial window
            self.window = self.create_window()
            
            logger.info("🚀 QuickDeck started")
            
            # Main event loop
            while True:
                event, values = self.window.read(timeout=100)
                
                if event == sg.WIN_CLOSED:
                    break
                elif event == "Escape:27":  # ESC key
                    break
                elif event == "r:17":  # Ctrl+R
                    self.execute_reload_action()
                elif event == "config":
                    self.config_gui.show_config_window()
                elif event == "reload":
                    self.execute_reload_action()
                elif event.startswith("btn_"):
                    self.handle_button_click(event)
                
                # Update status display
                if self.window and "status" in self.window.AllKeysDict:
                    self.window["status"].update(self.status_msg)
            
        except Exception as e:
            logger.error(f"❌ Application error: {e}")
            sg.popup_error(f"Application error: {e}", title="QuickDeck Error")
        
        finally:
            if self.window:
                self.window.close()
            logger.info("🛑 QuickDeck stopped")


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description="QuickDeck - Minimal command launcher")
    parser.add_argument("config", nargs="?", help="Path to config file (default: quickdeck.json)")
    
    args = parser.parse_args()
    
    try:
        app = QuickDeck(args.config)
        app.run()
    except KeyboardInterrupt:
        print("\n👋 QuickDeck stopped by user")
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()