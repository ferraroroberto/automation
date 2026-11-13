"""
Shared vCard TEL/EMAIL/ADR TYPE-label parsing, used by both unify_vcards.py
and convert_to_apple.py so the two never drift apart. A position-based split
on the whole line mis-parses the common single-semicolon legacy shape like
"TEL;HOME:+123": splitting on ";" before isolating the value from the prefix
leaves "HOME:+123" as the extracted type, or lets it leak into the
reconstructed value for ADR.

extract_type_label() extracts TYPE= via regex first (correct per the vCard
spec, since TYPE= may appear anywhere in the property parameter list),
falling back to the legacy positional convention
(`PROPERTY;TYPE:VALUE`, using the second semicolon-separated segment of the
*prefix only* - i.e. the part before the first ':') when no TYPE= parameter
is present at all. Both callers must isolate `prefix` (everything before the
first ':') before calling this - do not pass the whole raw line.
"""

import re
from typing import Optional


def has_type_param(text: str) -> bool:
    """True when `text` carries a TYPE= parameter; the name is case-insensitive in vCard (iCloud writes `type=`)."""
    return re.search(r'TYPE=', text, re.IGNORECASE) is not None


def extract_type_label(prefix: str, *, legacy_positional_fallback: bool = True) -> Optional[str]:
    """
    Extract the TYPE value from a vCard property prefix (the part before the
    first ':'), e.g. "TEL;TYPE=HOME" -> "HOME", or the legacy "TEL;HOME" ->
    "HOME". Returns None when neither shape matches - callers should leave
    their own default type label untouched in that case (mirrors the
    original per-field parsing's if/elif structure exactly).

    legacy_positional_fallback: unify_vcards.py's original parsing only ever
    applied the "second semicolon-separated segment" fallback for TEL, not
    for EMAIL/ADR (which only ever matched the regex path) - pass False to
    reproduce that narrower EMAIL/ADR behavior; leave True (the default) for
    TEL, and for convert_to_apple.py's callers (which only invoke this once
    they've already confirmed no TYPE= is present, so the fallback is the
    only path that can fire there).
    """
    if has_type_param(prefix):
        type_match = re.search(r'TYPE=([^;:]+)', prefix, re.IGNORECASE)
        if type_match:
            return type_match.group(1).upper()
        return None

    if legacy_positional_fallback and ";" in prefix:
        parts = prefix.split(";")
        if len(parts) > 1:
            return parts[1].upper()

    return None
