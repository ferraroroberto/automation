"""GUI-free logic of the region-screenshot tool: config, hotkey parsing, file naming, crop maths.

Everything the tool decides lives here so it is unit-testable without a display;
``snap_tool.py`` only wires it to Win32, Tk and the tray.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, NamedTuple, Optional, Tuple

DEFAULT_HOTKEY = "ctrl+shift+a"
DEFAULT_JPEG_QUALITY = 90
#: Smaller selections are accidental clicks, not captures.
MIN_SELECTION_SIDE = 3

# RegisterHotKey modifier flags.
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
#: Holding the keys down fires once, not once per key repeat.
MOD_NOREPEAT = 0x4000

_MODIFIERS = {
    "ctrl": MOD_CONTROL,
    "control": MOD_CONTROL,
    "shift": MOD_SHIFT,
    "alt": MOD_ALT,
    "win": MOD_WIN,
}
_NAMED_KEYS = {"printscreen": 0x2C, "space": 0x20}
_FUNCTION_KEY = re.compile(r"^f([1-9]|1\d|2[0-4])$")
_KNOWN_CONFIG_KEYS = {"folder", "hotkey", "jpeg_quality", "copy_to_clipboard"}


class ConfigError(ValueError):
    """The config file or hotkey text cannot be used; the message says which setting and why."""


class Rect(NamedTuple):
    """A pixel rectangle in image coordinates: ``right``/``bottom`` are exclusive."""

    left: int
    top: int
    right: int
    bottom: int


@dataclass(frozen=True)
class SnapConfig:
    folder: Path
    hotkey: str
    jpeg_quality: int
    copy_to_clipboard: bool


def parse_hotkey(text: str) -> Tuple[int, int]:
    """Turn ``"ctrl+shift+a"`` into ``(modifier flags, virtual-key code)`` for RegisterHotKey."""
    parts = [part.strip().lower() for part in text.split("+") if part.strip()]
    if len(parts) < 2:
        raise ConfigError(f"hotkey {text!r} needs at least one modifier and a key, e.g. 'ctrl+shift+a'")
    *modifier_names, key_name = parts
    modifiers = 0
    for name in modifier_names:
        if name not in _MODIFIERS:
            raise ConfigError(f"hotkey {text!r}: unknown modifier {name!r} (use ctrl, shift, alt or win)")
        modifiers |= _MODIFIERS[name]
    if len(key_name) == 1 and key_name.isalnum():
        key = ord(key_name.upper())
    elif _FUNCTION_KEY.match(key_name):
        key = 0x70 + int(key_name[1:]) - 1
    elif key_name in _NAMED_KEYS:
        key = _NAMED_KEYS[key_name]
    else:
        raise ConfigError(f"hotkey {text!r}: unknown key {key_name!r} (a letter, digit, f1-f24, printscreen or space)")
    return modifiers, key


def _default_folder() -> Path:
    return Path.home() / "Pictures" / "snaps"


def parse_config(raw: Dict[str, Any]) -> SnapConfig:
    """Validate a decoded config object; unknown keys are rejected so a typo is not silently ignored."""
    unknown = sorted(set(raw) - _KNOWN_CONFIG_KEYS)
    if unknown:
        raise ConfigError(f"unknown config key(s) {unknown}; valid keys are {sorted(_KNOWN_CONFIG_KEYS)}")

    folder = raw.get("folder")
    if folder is None:
        folder_path = _default_folder()
    elif isinstance(folder, str) and folder.strip():
        folder_path = Path(os.path.expandvars(folder)).expanduser()
    else:
        raise ConfigError("'folder' must be a non-empty path string")

    hotkey = raw.get("hotkey", DEFAULT_HOTKEY)
    if not isinstance(hotkey, str):
        raise ConfigError("'hotkey' must be a string such as 'ctrl+shift+a'")
    parse_hotkey(hotkey)

    quality = raw.get("jpeg_quality", DEFAULT_JPEG_QUALITY)
    if isinstance(quality, bool) or not isinstance(quality, int) or not 1 <= quality <= 100:
        raise ConfigError(f"'jpeg_quality' must be a whole number from 1 to 100, got {quality!r}")

    clipboard = raw.get("copy_to_clipboard", True)
    if not isinstance(clipboard, bool):
        raise ConfigError("'copy_to_clipboard' must be true or false")

    return SnapConfig(folder=folder_path, hotkey=hotkey.lower(), jpeg_quality=quality, copy_to_clipboard=clipboard)


def load_config(path: Path) -> SnapConfig:
    """Read ``path``; a missing file means all defaults, an unreadable or invalid one raises ``ConfigError``."""
    if not path.exists():
        return parse_config({})
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"cannot read {path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"{path} must hold a JSON object")
    return parse_config(raw)


def unique_path(folder: Path, moment: datetime) -> Path:
    """``snap_YYYY-MM-DD_HH-MM-SS-mmm.jpg`` in ``folder``; a taken name gets ``_2``, ``_3``... so nothing is overwritten."""
    stem = f"snap_{moment:%Y-%m-%d_%H-%M-%S}-{moment.microsecond // 1000:03d}"
    candidate = folder / f"{stem}.jpg"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{stem}_{counter}.jpg"
        counter += 1
    return candidate


def window_geometry(left: int, top: int, width: int, height: int) -> str:
    """Tk geometry string for a window at a screen position; monitors left of or above the primary have negative origins."""
    return f"{width}x{height}{left:+d}{top:+d}"


def screen_to_image(x: int, y: int, origin: Tuple[int, int]) -> Tuple[int, int]:
    """Map a screen position to a pixel of the captured image whose top-left sits at ``origin`` on the virtual screen."""
    return x - origin[0], y - origin[1]


def selection_rect(
    start: Tuple[int, int], end: Tuple[int, int], image_size: Tuple[int, int]
) -> Optional[Rect]:
    """The dragged rectangle in image pixels: ordered, clamped to the image, ``None`` when too small to be a capture."""
    width, height = image_size
    left = max(0, min(start[0], end[0]))
    top = max(0, min(start[1], end[1]))
    right = min(width, max(start[0], end[0]))
    bottom = min(height, max(start[1], end[1]))
    if right - left < MIN_SELECTION_SIDE or bottom - top < MIN_SELECTION_SIDE:
        return None
    return Rect(left, top, right, bottom)
