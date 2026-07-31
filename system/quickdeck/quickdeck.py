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

Layout (audit issue #92 split this module into four):
- `quickdeck.py`            - this file: config load, main window, event loop, CLI
- `quickdeck_config_gui.py` - the settings/button editor windows
- `quickdeck_render.py`     - PIL emoji + label button-image rendering
- `quickdeck_actions.py`    - the shell/url/copy/python action executors
"""

import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict, Optional

import PySimpleGUI as sg

from quickdeck_actions import ActionContext, execute_action
from quickdeck_config_gui import QuickDeckConfig
from quickdeck_render import button_image_size, create_combined_button_image, create_emoji_image


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


class QuickDeck:
    """Main QuickDeck application class."""

    def __init__(self, config_path: Optional[str] = None):
        """Initialize QuickDeck with config file."""
        self.config_path = Path(config_path) if config_path else Path("quickdeck.json")
        self.config: Dict[str, Any] = {}
        self.window: Optional[sg.Window] = None
        self.status_msg = "🚀 QuickDeck ready"
        self.emoji_cache: Dict[str, bytes] = {}  # Cache for rendered emoji images
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

    def set_status(self, msg: str) -> None:
        """Update status message."""
        self.status_msg = msg
        logger.info(msg)

    def _action_context(self) -> ActionContext:
        """Bind this window's status line and popups for the action executors."""
        return ActionContext(
            set_status=self.set_status,
            show_error=lambda message, title: sg.popup_error(message, title=title),
            show_output=lambda body, title: sg.popup_scrolled(body, title=title, size=(50, 20)),
            reload=self.execute_reload_action,
        )

    def _build_button(self, btn_index: int, button_config: Dict[str, Any],
                      emoji_size: int, text_size: int) -> sg.Button:
        """Build one grid button, with a rendered emoji+label image when it has an emoji."""
        label = button_config.get("label", f"Button {btn_index + 1}")
        emoji = button_config.get("emoji", "")
        hint = button_config.get("hint", label)

        if not emoji:
            return sg.Button(
                label,
                key=f"btn_{btn_index}",
                size=(None, None),  # Let PySimpleGUI auto-size based on content
                tooltip=hint,
                expand_x=True,
                expand_y=True,
                font=("Arial", text_size)
            )

        theme_config = self.config.get("theme", {})
        button_bg = theme_config.get("button_bg", "#1F2937")
        button_fg = theme_config.get("fg", "#E5E7EB")

        emoji_data = create_emoji_image(emoji, emoji_size, self.emoji_cache)
        combined_image_data = create_combined_button_image(
            emoji_data, label, emoji_size, text_size, button_fg, button_bg
        )
        image_width, image_height = button_image_size(label, emoji_size, text_size)

        return sg.Button(
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

    def create_window(self) -> sg.Window:
        """Create the main window with button grid."""
        grid_config = self.config.get("grid", {"rows": 2, "cols": 4, "pad": 8, "button_height": 80, "emoji_size": 48, "text_size": 12, "text_position": "below"})
        buttons = self.config.get("buttons", [])
        emoji_size = grid_config.get("emoji_size", 48)
        text_size = grid_config.get("text_size", 12)

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
                button_config = buttons[btn_index] if btn_index < len(buttons) else {}
                button_row.append(self._build_button(btn_index, button_config, emoji_size, text_size))
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

        # Size the grid off the widest label and the padded emoji box
        max_text_width = 0.0
        for btn_index in range(min(len(buttons), total_buttons)):
            if buttons[btn_index]:
                label = buttons[btn_index].get("label", f"Button {btn_index + 1}")
                # Estimate text width (rough approximation)
                max_text_width = max(max_text_width, len(label) * (text_size * 0.6))

        cell_width, cell_height = button_image_size("", emoji_size, text_size)
        button_width = max(cell_width, max_text_width + 20)
        button_height = max(cell_height, 60)  # Minimum button height

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

        return sg.Window(
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
            btn_index = int(event.split("_")[1])
        except (IndexError, ValueError):
            self.set_status(f"❌ Unrecognised button event: {event}")
            return

        buttons = self.config.get("buttons", [])
        if btn_index < len(buttons):
            execute_action(buttons[btn_index], self._action_context())

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


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(description="QuickDeck - Minimal command launcher")
    parser.add_argument("config", nargs="?", help="Path to config file (default: quickdeck.json)")

    args = parser.parse_args()

    try:
        app = QuickDeck(args.config)
        app.run()
    except KeyboardInterrupt:
        logger.info("👋 QuickDeck stopped by user")
    except Exception as e:
        logger.error("❌ Fatal error: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
