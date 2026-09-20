"""JWT helpers. The token is a credential: nothing here ever logs its value."""

from __future__ import annotations

import base64
import json
from datetime import datetime, timedelta, timezone
from typing import Optional


def normalize_token(raw: Optional[str]) -> Optional[str]:
    """Strip quotes / a leading 'Bearer ' from a stored token; None when empty."""
    if not raw:
        return None
    token = raw.strip().strip('"').strip("'")
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    return token or None


def jwt_expiry(token: str) -> Optional[datetime]:
    """UTC expiry from the `exp` claim, or None when the token is not a decodable JWT."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return datetime.fromtimestamp(int(claims["exp"]), tz=timezone.utc)
    except (IndexError, KeyError, ValueError, TypeError):
        return None


def time_left(token: str, now: datetime) -> Optional[timedelta]:
    expiry = jwt_expiry(token)
    return None if expiry is None else expiry - now
