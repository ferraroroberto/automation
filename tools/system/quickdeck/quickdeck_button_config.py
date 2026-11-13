"""
Pure button-config edits for the QuickDeck configuration editor.

Kept free of PySimpleGUI so the slot and field rules can be tested headlessly.
The button grid is positional (the example config pads unused slots with `{}`),
so a slot is emptied in place rather than removed from the list.
"""

import subprocess
from typing import Any, Dict, List

# Keys that belong to exactly one button type; switching type drops the others.
_TYPE_KEYS = {
    "url": {"url"},
    "copy": {"text"},
    "python": {"code"},
    "shell": {"cmd", "args", "cwd"},
}

_DEFAULT_SHELL = {"cmd": "echo", "args": ["Button configured - add your command"]}


def split_command(action: str) -> List[str]:
    """Split a shell action into argv, keeping a quoted path with spaces as one argument."""
    argv = []
    token = ""
    quote = ""
    started = False
    for char in action:
        if quote:
            if char == quote:
                quote = ""
            else:
                token += char
        elif char in ('"', "'"):
            quote = char
            started = True
        elif char.isspace():
            if started or token:
                argv.append(token)
                token = ""
                started = False
        else:
            token += char
    if started or token:
        argv.append(token)
    return argv


def join_command(cmd: str, args: List[str]) -> str:
    """Inverse of `split_command`: quote any part that contains whitespace."""
    return subprocess.list2cmdline([cmd, *args])


def apply_button_values(existing: Dict[str, Any], values: Dict[str, Any], slot_number: int) -> Dict[str, Any]:
    """Return `existing` updated from the edit dialog's `values`.

    The existing dict is the starting point, so keys the dialog does not show
    (a shell button's `cwd`) survive a save. Raises ValueError for a URL that is
    not http(s) instead of silently substituting another address.
    """
    btn_type = values.get("type", "shell")
    action = values.get("action", "").strip()

    config = {
        key: value
        for key, value in existing.items()
        if not any(key in keys for kind, keys in _TYPE_KEYS.items() if kind != btn_type)
    }
    config["label"] = values.get("label", "").strip() or f"Button {slot_number}"
    config["emoji"] = values.get("emoji", "").strip()
    config["hint"] = values.get("hint", "").strip()
    config["type"] = btn_type

    if btn_type == "url":
        if not action.startswith(("http://", "https://")):
            raise ValueError("URL must start with http:// or https://")
        config["url"] = action
    elif btn_type == "copy":
        config["text"] = action
    elif btn_type == "python":
        config["code"] = action
    elif btn_type == "shell":
        parts = split_command(action)
        if not parts:
            parts = [_DEFAULT_SHELL["cmd"], *_DEFAULT_SHELL["args"]]
        config["cmd"] = parts[0]
        if len(parts) > 1:
            config["args"] = parts[1:]
        else:
            config.pop("args", None)
    return config


def clear_slot(buttons: List[Dict[str, Any]], btn_index: int) -> bool:
    """Empty one slot in place so later buttons keep their positions; False if already empty."""
    if btn_index >= len(buttons) or not buttons[btn_index]:
        return False
    buttons[btn_index] = {}
    return True
