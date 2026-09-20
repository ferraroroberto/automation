"""Telegram notifications through the fleet notifier (notify_send.py).

The bot token and chat routing live in fleet-config; this project only needs
the interpreter + script path (both in .env) and, optionally, a category.
Without them a notification degrades to a log line.
"""

from __future__ import annotations

import logging
import subprocess
import sys
from typing import Dict, Protocol

logger = logging.getLogger("parking.notify")

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0


class Notifier(Protocol):
    def send(self, text: str) -> bool: ...


class FleetNotifier:
    def __init__(self, env: Dict[str, str]) -> None:
        self._python = env.get("NOTIFY_PYTHON", "")
        self._script = env.get("NOTIFY_SCRIPT", "")
        self._category = env.get("NOTIFY_CATEGORY", "attention")
        self._chat = env.get("NOTIFY_CHAT", "")

    def command(self, text: str) -> list:
        """Notifier argv: an explicit chat id wins over the routing category."""
        target = ["--chat", self._chat] if self._chat else ["--category", self._category]
        return [self._python, self._script, *target, "--text", text]

    def send(self, text: str) -> bool:
        if not (self._python and self._script):
            logger.warning("⚠️ NOTIFY_PYTHON/NOTIFY_SCRIPT not set; notification only logged: %s", text)
            return False
        try:
            result = subprocess.run(
                self.command(text),
                capture_output=True, text=True, timeout=60, creationflags=CREATE_NO_WINDOW,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            logger.error("❌ notifier failed to start: %s", exc)
            return False
        if result.returncode != 0:
            logger.error("❌ notifier exit %s", result.returncode)
            return False
        logger.info("✅ notification sent")
        return True
