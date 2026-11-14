"""GUI-free logic of the region-screenshot tool: config, hotkey parsing, file naming, crop and loupe maths.

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
from typing import Any, Dict, NamedTuple, Optional, Sequence, Tuple

DEFAULT_HOTKEY = "ctrl+shift+a"
DEFAULT_JPEG_QUALITY = 90
DEFAULT_FILE_PREFIX = "snap"
#: The tray's quality menu offers these; the config file accepts any 1-100.
QUALITY_CHOICES = (70, 80, 90, 95, 100)
MAX_PREFIX_LENGTH = 40
_INVALID_FILENAME_CHARS = frozenset('<>:"/\\|?*')
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
_KNOWN_CONFIG_KEYS = {"folder", "hotkey", "jpeg_quality", "copy_to_clipboard", "file_prefix", "show_loupe"}

#: Loupe: a square of source pixels around the cursor, magnified by a whole factor so every pixel stays a crisp block.
LOUPE_RADIUS = 12
LOUPE_ZOOM = 8
LOUPE_CELLS = 2 * LOUPE_RADIUS + 1
LOUPE_IMAGE_SIDE = LOUPE_CELLS * LOUPE_ZOOM
LOUPE_TEXT_HEIGHT = 22
#: Space between the cursor and the loupe box so the box never covers the pixel being placed.
LOUPE_CURSOR_GAP = 24
ARROW_KEYS = ("Left", "Right", "Up", "Down")
_ARROW_DIRECTIONS = {"Left": (-1, 0), "Right": (1, 0), "Up": (0, -1), "Down": (0, 1)}
NUDGE_STEP = 1
NUDGE_STEP_FAST = 10


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
    file_prefix: str = DEFAULT_FILE_PREFIX
    show_loupe: bool = True


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


def validate_prefix(text: str) -> str:
    """The filename prefix, stripped; raises ``ConfigError`` when it is empty, too long or holds a character Windows refuses in a filename."""
    prefix = text.strip()
    if not prefix:
        raise ConfigError("the filename prefix cannot be empty")
    if len(prefix) > MAX_PREFIX_LENGTH:
        raise ConfigError(f"the filename prefix is limited to {MAX_PREFIX_LENGTH} characters")
    bad = sorted({ch for ch in prefix if ch in _INVALID_FILENAME_CHARS or ord(ch) < 32})
    if bad:
        shown = " ".join(repr(ch) if ord(ch) < 32 else ch for ch in bad)
        raise ConfigError(f"the filename prefix cannot contain {shown}")
    return prefix


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

    prefix = raw.get("file_prefix", DEFAULT_FILE_PREFIX)
    if not isinstance(prefix, str):
        raise ConfigError("'file_prefix' must be a string such as 'snap'")

    show_loupe = raw.get("show_loupe", True)
    if not isinstance(show_loupe, bool):
        raise ConfigError("'show_loupe' must be true or false")

    return SnapConfig(
        folder=folder_path,
        hotkey=hotkey.lower(),
        jpeg_quality=quality,
        copy_to_clipboard=clipboard,
        file_prefix=validate_prefix(prefix),
        show_loupe=show_loupe,
    )


def _read_raw_config(path: Path) -> Dict[str, Any]:
    """The decoded JSON object at ``path``; a missing file is an empty object."""
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"cannot read {path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"{path} must hold a JSON object")
    return raw


def load_config(path: Path) -> SnapConfig:
    """Read ``path``; a missing file means all defaults, an unreadable or invalid one raises ``ConfigError``."""
    return parse_config(_read_raw_config(path))


def update_config_file(path: Path, key: str, value: Any) -> SnapConfig:
    """Set one key in the file and return the resulting config.

    Only that key changes (the other values stay as the user wrote them, e.g. a ``%VAR%`` folder), the
    result is validated before anything is written, and the file is replaced in one step so a crash
    never leaves half a config behind.
    """
    raw = _read_raw_config(path)
    raw[key] = value
    config = parse_config(raw)
    staging = path.with_name(path.name + ".tmp")
    # runtime-write-exempt: gitignored per-user config.json beside the tool, which the user also edits by hand and the tool reads at start
    staging.write_text(json.dumps(raw, indent=2) + "\n", encoding="utf-8")
    os.replace(staging, path)
    return config


def unique_path(folder: Path, moment: datetime, prefix: str = DEFAULT_FILE_PREFIX) -> Path:
    """``<prefix>_YYYY-MM-DD_HH-MM-SS-mmm.jpg`` in ``folder``; a taken name gets ``_2``, ``_3``... so nothing is overwritten."""
    stem = f"{prefix}_{moment:%Y-%m-%d_%H-%M-%S}-{moment.microsecond // 1000:03d}"
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


def _clamped_rect(start: Tuple[int, int], end: Tuple[int, int], image_size: Tuple[int, int]) -> Rect:
    width, height = image_size
    return Rect(
        max(0, min(start[0], end[0])),
        max(0, min(start[1], end[1])),
        min(width, max(start[0], end[0])),
        min(height, max(start[1], end[1])),
    )


def selection_rect(
    start: Tuple[int, int], end: Tuple[int, int], image_size: Tuple[int, int]
) -> Optional[Rect]:
    """The dragged rectangle in image pixels: ordered, clamped to the image, ``None`` when too small to be a capture."""
    rect = _clamped_rect(start, end, image_size)
    if rect.right - rect.left < MIN_SELECTION_SIDE or rect.bottom - rect.top < MIN_SELECTION_SIDE:
        return None
    return rect


def selection_size(
    start: Tuple[int, int], end: Tuple[int, int], image_size: Tuple[int, int]
) -> Tuple[int, int]:
    """Width and height of the drag as it would be cropped, small or not, so the loupe can show it live."""
    rect = _clamped_rect(start, end, image_size)
    return rect.right - rect.left, rect.bottom - rect.top


def loupe_source_box(center: Tuple[int, int]) -> Rect:
    """The image pixels the loupe magnifies: a square centred on ``center``. It may reach past the image; the crop pads that with black."""
    return Rect(
        center[0] - LOUPE_RADIUS,
        center[1] - LOUPE_RADIUS,
        center[0] + LOUPE_RADIUS + 1,
        center[1] + LOUPE_RADIUS + 1,
    )


def monitor_for(point: Tuple[int, int], monitors: Sequence[Rect]) -> Rect:
    """The monitor holding ``point``, else the nearest one, so the loupe is placed on a real screen."""

    def distance(rect: Rect) -> int:
        dx = max(rect.left - point[0], 0, point[0] - (rect.right - 1))
        dy = max(rect.top - point[1], 0, point[1] - (rect.bottom - 1))
        return dx * dx + dy * dy

    return min(monitors, key=distance)


def place_loupe(
    cursor: Tuple[int, int], box_size: Tuple[int, int], bounds: Rect, gap: int = LOUPE_CURSOR_GAP
) -> Tuple[int, int]:
    """Top-left of the loupe box: below-right of the cursor, flipped to the other side near the right/bottom edge of ``bounds``, then clamped inside it."""
    width, height = box_size
    x = cursor[0] + gap
    y = cursor[1] + gap
    if x + width > bounds.right:
        x = cursor[0] - gap - width
    if y + height > bounds.bottom:
        y = cursor[1] - gap - height
    x = max(bounds.left, min(x, bounds.right - width))
    y = max(bounds.top, min(y, bounds.bottom - height))
    return x, y


def loupe_label(cursor: Tuple[int, int], selection: Optional[Tuple[int, int]]) -> str:
    """``x,y`` under the cursor, plus ``(w×h)`` while a selection is being dragged."""
    text = f"{cursor[0]},{cursor[1]}"
    if selection is not None:
        text += f" ({selection[0]}×{selection[1]})"
    return text


def nudge_point(point: Tuple[int, int], key: str, fast: bool, bounds: Rect) -> Tuple[int, int]:
    """``point`` moved one pixel (ten when ``fast``) in the arrow ``key``'s direction, kept on the last pixel of ``bounds``."""
    dx, dy = _ARROW_DIRECTIONS[key]
    step = NUDGE_STEP_FAST if fast else NUDGE_STEP
    return (
        max(bounds.left, min(point[0] + dx * step, bounds.right - 1)),
        max(bounds.top, min(point[1] + dy * step, bounds.bottom - 1)),
    )
