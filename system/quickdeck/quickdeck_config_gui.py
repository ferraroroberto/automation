#!/usr/bin/env python3
"""
QuickDeck configuration editor.

A self-contained PySimpleGUI editor for `quickdeck.json`: general/window/grid/
theme settings plus a per-button edit dialog. It talks to the running app
through the small surface documented on `QuickDeckConfig.__init__` (audit
issue #92), so it can be opened against any object exposing those attributes.
"""

import json
import re
from typing import Any, Dict

import PySimpleGUI as sg

# Hex colour literal, e.g. #111827 - the only format the theme accepts.
_COLOR_PATTERN = r'^#[0-9A-Fa-f]{6}$'


class QuickDeckConfig:
    """Configuration GUI for QuickDeck."""

    def __init__(self, quickdeck_instance):
        """
        Initialize configuration GUI.

        `quickdeck_instance` must expose `.config` (the parsed dict),
        `.config_path` (where to write it back), and `.execute_reload_action()`
        (rebuild the main window after a save).
        """
        self.qd = quickdeck_instance
        self.config_window = None

    def _write_config(self) -> None:
        """Persist the in-memory config back to its JSON file."""
        with open(self.qd.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.qd.config, f, indent=2, ensure_ascii=False)

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
            self._write_config()

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
                self._write_config()

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
            if not re.match(_COLOR_PATTERN, theme_config["bg"]):
                sg.popup_error("Background color must be a valid hex color (e.g., #111827)!", title="Error")
                return

            if not re.match(_COLOR_PATTERN, theme_config["fg"]):
                sg.popup_error("Text color must be a valid hex color (e.g., #E5E7EB)!", title="Error")
                return

            if not re.match(_COLOR_PATTERN, theme_config["button_bg"]):
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

            self.qd.config["geometry"] = f"{window_width}x{window_height}+{x}+{y}"

            # Update configuration
            self.qd.config["grid"] = grid_config
            self.qd.config["theme"] = theme_config

            # Save to file
            self._write_config()

            # Reload the main window
            self.qd.execute_reload_action()

            sg.popup("✅ Configuration saved and applied!", title="Success")

        except ValueError as e:
            sg.popup_error(f"❌ Invalid input values: {e}", title="Error")
        except Exception as e:
            sg.popup_error(f"❌ Failed to save configuration: {e}", title="Error")
