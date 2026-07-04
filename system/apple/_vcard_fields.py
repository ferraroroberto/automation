"""
Shared vCard TEL/EMAIL/ADR TYPE-label parsing.

Consolidates the TYPE= extraction logic previously duplicated - and drifted -
between unify_vcards.py (regex `TYPE=([^;:]+)`) and convert_to_apple.py
(position-based split on the WHOLE line, which mis-parsed the common
single-semicolon legacy shape like "TEL;HOME:+123" - because it split on
";" before isolating the value from the prefix, "HOME:+123" ended up as the
extracted type, or leaked into the reconstructed value for ADR - see the
audit issue #64 dedup Notes for the concrete before/after traces).

extract_type_label() implements unify_vcards.py's behavior (regex-first,
correct per the vCard spec since TYPE= may appear anywhere in the property
parameter list), falling back to the legacy positional convention
(`PROPERTY;TYPE:VALUE`, using the second semicolon-separated segment of the
*prefix only* - i.e. the part before the first ':') when no TYPE= parameter
is present at all. Both callers must isolate `prefix` (everything before the
first ':') before calling this - do not pass the whole raw line.
"""

import re
from typing import Optional


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
    if "TYPE=" in prefix:
        type_match = re.search(r'TYPE=([^;:]+)', prefix)
        if type_match:
            return type_match.group(1).upper()
        return None

    if legacy_positional_fallback and ";" in prefix:
        parts = prefix.split(";")
        if len(parts) > 1:
            return parts[1].upper()

    return None
