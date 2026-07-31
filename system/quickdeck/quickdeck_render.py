#!/usr/bin/env python3
"""
QuickDeck button-image rendering.

Pure PIL compositing: turn an emoji glyph plus a label into the PNG bytes
PySimpleGUI paints onto a button. No GUI toolkit is imported here, so this
module can be exercised (and eyeballed) without a display or a PySimpleGUI
licence — which was the point of splitting it out of `quickdeck.py`
(audit issue #92).
"""

import io
import logging
import os
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger("quickdeck")

# Transparent padding added around an emoji glyph on every side, in pixels.
# Complex emoji (flags, ZWJ sequences) overflow their nominal box and get
# clipped without it; every size calculation has to account for it too.
EMOJI_PADDING = 8


def _emoji_font_paths() -> List[str]:
    """Platform fallback list of TTF paths for rendering emoji glyphs."""
    if os.name == 'nt':
        return [
            "C:/Windows/Fonts/seguiemj.ttf",  # Segoe UI Emoji
            "C:/Windows/Fonts/segmdl2.ttf",  # Segoe MDL2 Assets
        ]
    return [
        "/System/Library/Fonts/Apple Color Emoji.ttc",  # macOS
        "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf",  # Linux
    ]


def _text_font_paths() -> List[str]:
    """Platform fallback list of TTF paths for rendering button label text."""
    if os.name == 'nt':
        return [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
        ]
    return [
        "/System/Library/Fonts/Arial.ttf",  # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
    ]


def _load_font(font_paths: List[str], size: int) -> ImageFont.FreeTypeFont:
    """
    Try each TTF path in order, returning the first that loads at `size`;
    falls back to PIL's built-in default font if none are usable.

    Consolidates the "pick a TTF from a platform fallback list, load with
    ImageFont.truetype, fall through to load_default()" pattern previously
    copy-pasted four times across QuickDeck (dedup: audit issue #64).
    """
    font = None
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                font = ImageFont.truetype(font_path, size)
                break
            except OSError:
                continue
    if font is None:
        font = ImageFont.load_default()
    return font


def measure_text_width(text: str, text_size: int) -> float:
    """
    Width in pixels of `text` rendered at `text_size` with the label font.

    Falls back to a character-count approximation if no usable font is found,
    matching the estimate the three former inline copies of this measurement
    used (audit issue #92).
    """
    try:
        font = _load_font(_text_font_paths(), text_size)
        temp_img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
        temp_draw = ImageDraw.Draw(temp_img)
        bbox = temp_draw.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0]
    except Exception:
        return len(text) * (text_size * 0.6)


def button_image_size(label: str, emoji_size: int, text_size: int) -> Tuple[float, int]:
    """
    (width, height) PySimpleGUI should reserve for a rendered button image.

    Width has to fit whichever is wider — the padded emoji or the label — so
    long labels aren't clipped; height stacks emoji over text.
    """
    padded_emoji_size = emoji_size + (EMOJI_PADDING * 2)
    width = max(padded_emoji_size, measure_text_width(label, text_size)) + 20
    height = padded_emoji_size + text_size + 20
    return width, height


def _to_png_bytes(img: Image.Image) -> bytes:
    """Encode a PIL image to in-memory PNG bytes."""
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    return img_bytes.getvalue()


def create_emoji_image(
    emoji: str,
    size: int = 32,
    cache: Optional[Dict[str, bytes]] = None,
) -> bytes:
    """
    Render `emoji` to PNG bytes at `size`, padded so complex glyphs aren't cut.

    `cache` is an optional caller-owned dict keyed by "<emoji>_<size>"; the app
    passes its own so repeated redraws don't re-rasterise.
    """
    cache_key = f"{emoji}_{size}"
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    try:
        img_size = size + (EMOJI_PADDING * 2)
        img = Image.new('RGBA', (img_size, img_size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Try to use a system emoji font
        try:
            font = _load_font(_emoji_font_paths(), size)
        except Exception:
            font = ImageFont.load_default()

        # Calculate text position to center the emoji with padding
        bbox = draw.textbbox((0, 0), emoji, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        # Center the emoji in the padded image with extra vertical padding
        x = (img_size - text_width) // 2
        y = (img_size - text_height) // 2 + 2  # Slight vertical adjustment for better centering

        draw.text((x, y), emoji, font=font, embedded_color=True)

        data = _to_png_bytes(img)
        if cache is not None:
            cache[cache_key] = data
        return data

    except Exception as e:
        logger.warning(f"⚠️ Failed to create emoji image for {emoji}: {e}")
        # Return a simple colored square as fallback
        img = Image.new('RGBA', (size, size), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.rectangle([0, 0, size - 1, size - 1], outline=(100, 100, 100, 255))
        return _to_png_bytes(img)


def create_combined_button_image(
    emoji_data: bytes,
    text: str,
    emoji_size: int,
    text_size: int,
    text_color: str,
    bg_color: str,
) -> bytes:
    """Create a combined image with emoji on top and text below."""
    try:
        padding = 10
        text_height = text_size + 10
        text_width = measure_text_width(text, text_size)

        # Total dimensions - use the larger of emoji width or text width.
        # The emoji image is already padded, so account for that here too.
        padded_emoji_size = emoji_size + (EMOJI_PADDING * 2)
        total_width = int(max(padded_emoji_size, text_width) + (padding * 2))
        total_height = int(padded_emoji_size + text_height + (padding * 2))

        img = Image.new('RGBA', (total_width, total_height), bg_color)
        draw = ImageDraw.Draw(img)

        # Paste the (already padded) emoji at the top, centered
        emoji_img = Image.open(io.BytesIO(emoji_data))
        emoji_x = (total_width - padded_emoji_size) // 2
        emoji_y = padding
        img.paste(emoji_img, (emoji_x, emoji_y), emoji_img)

        # Add text below emoji
        try:
            font = _load_font(_text_font_paths(), text_size)
        except Exception:
            font = ImageFont.load_default()

        bbox = draw.textbbox((0, 0), text, font=font)
        text_x = int((total_width - (bbox[2] - bbox[0])) // 2)
        text_y = int(emoji_y + padded_emoji_size + 5)

        draw.text((text_x, text_y), text, font=font, fill=text_color)

        return _to_png_bytes(img)

    except Exception as e:
        logger.warning(f"⚠️ Failed to create combined button image: {e}")
        # Return the original emoji data as fallback
        return emoji_data
