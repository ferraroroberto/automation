"""Persistent run state: backoff and alert de-duplication."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger("parking.state")


@dataclass
class State:
    failures: int = 0
    next_allowed: Optional[datetime] = None
    alerts: Dict[str, str] = field(default_factory=dict)  # alert key -> ISO time last sent

    def in_backoff(self, now: datetime) -> bool:
        return self.next_allowed is not None and now < self.next_allowed

    def alert_due(self, key: str, now: datetime, cooldown: timedelta) -> bool:
        last = self.alerts.get(key)
        return last is None or now - datetime.fromisoformat(last) >= cooldown

    def mark_alert(self, key: str, now: datetime) -> None:
        self.alerts[key] = now.isoformat()


def load_state(path: Path) -> State:
    if not path.exists():
        return State()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        nxt = raw.get("next_allowed")
        return State(
            failures=int(raw.get("failures", 0)),
            next_allowed=datetime.fromisoformat(nxt) if nxt else None,
            alerts=dict(raw.get("alerts", {})),
        )
    except (ValueError, OSError) as exc:
        logger.warning("⚠️ unreadable state file (%s); starting fresh", exc)
        return State()


def save_state(path: Path, state: State) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {"failures": state.failures,
             "next_allowed": state.next_allowed.isoformat() if state.next_allowed else None,
             "alerts": state.alerts},
            indent=2,
        ),
        encoding="utf-8",
    )
