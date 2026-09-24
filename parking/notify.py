"""Telegram notifications through the fleet notifier (notify_send.py).

The bot token and chat routing live in fleet-config; this project only needs
the interpreter + script path (both in .env) and, optionally, a category.
Without them a notification degrades to a log line.

Chat cleanup: with a `ledger_path`, the ids of every *disposable* message
(sweep reports, burst pings, verbose logs) are recorded, and the first send of
the next run deletes them - so the chat shows the latest run, not a scroll of
screenshots. Messages sent without `disposable=True` (bookings, failures,
login alerts) are never recorded and so never deleted.
"""

from __future__ import annotations

import json
import logging
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Dict, List, Optional, Protocol

logger = logging.getLogger("parking.notify")

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0

# Telegram allows a bot ~20 messages a minute in a group; pings stay under it
# by batching whatever queued up while waiting into one message.
PING_MIN_INTERVAL_SECONDS = 3.2
PING_FLUSH_TIMEOUT_SECONDS = 60.0


class Notifier(Protocol):
    def send(self, text: str, disposable: bool = False) -> bool: ...
    def send_file(self, text: str, path: str, disposable: bool = False) -> bool: ...
    def send_files(self, text: str, paths: List[str], disposable: bool = False) -> bool: ...
    def ping(self, text: str) -> None: ...
    def close(self) -> None: ...


class FleetNotifier:
    def __init__(self, env: Dict[str, str], ledger_path: Optional[Path] = None) -> None:
        self._python = env.get("NOTIFY_PYTHON", "")
        self._script = env.get("NOTIFY_SCRIPT", "")
        self._category = env.get("NOTIFY_CATEGORY", "attention")
        self._chat = env.get("NOTIFY_CHAT", "")
        self._ledger_path = ledger_path
        self._lock = threading.Lock()
        self._cleaned = False
        self._pings: "queue.Queue[Optional[str]]" = queue.Queue()
        self._ping_thread: Optional[threading.Thread] = None
        self._last_sent = float("-inf")  # time.monotonic() of the last send; paces the pings

    def _target(self) -> list:
        """An explicit chat id wins over the routing category."""
        return ["--chat", self._chat] if self._chat else ["--category", self._category]

    def command(self, text: str) -> list:
        return [self._python, self._script, *self._target(), "--text", text]

    def file_command(self, text: str, path: str) -> list:
        """Argv for a text message with an attached file (e.g. a screenshot)."""
        return [self._python, self._script, *self._target(), "--file", path, "--text", text]

    def files_command(self, text: str, paths: List[str]) -> list:
        """Argv for a text message with 2-10 attached files, sent as one Telegram message."""
        return [self._python, self._script, *self._target(), "--files", *paths, "--text", text]

    def delete_command(self, ids: List[int]) -> list:
        """Argv deleting the bot's own messages `ids` from the chat."""
        return [self._python, self._script, *self._target(), "--delete-ids", *map(str, ids)]

    def _exec(self, argv: list, timeout: float) -> Optional[subprocess.CompletedProcess]:
        if not (self._python and self._script):
            logger.warning("⚠️ NOTIFY_PYTHON/NOTIFY_SCRIPT not set; notification only logged: %s", argv[-1])
            return None
        try:
            return subprocess.run(
                argv, capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=timeout, creationflags=CREATE_NO_WINDOW,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            logger.error("❌ notifier failed to start: %s", exc)
            return None

    def _run(self, argv: list, timeout: float, disposable: bool = False) -> bool:
        self._cleanup_previous_run()
        if self._ledger_path is not None:
            argv = [*argv[:2], "--print-ids", *argv[2:]]  # the text stays last, for the log line above
        result = self._exec(argv, timeout)
        self._last_sent = time.monotonic()
        if result is None:
            return False
        if disposable and self._ledger_path is not None:
            self._record(parse_ids(result.stdout))
        if result.returncode != 0:
            logger.error("❌ notifier exit %s", result.returncode)
            return False
        logger.info("✅ notification sent")
        return True

    def _read_ledger(self) -> List[int]:
        try:
            raw = json.loads(self._ledger_path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return []
        except (OSError, ValueError) as exc:
            logger.warning("⚠️ message ledger unreadable (%s); starting a new one", exc)
            return []
        return [i for i in raw.get("disposable", []) if isinstance(i, int)]

    def _write_ledger(self, ids: List[int]) -> None:
        self._ledger_path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._ledger_path.with_suffix(".tmp")
        tmp.write_text(json.dumps({"disposable": ids}), encoding="utf-8")
        tmp.replace(self._ledger_path)

    def _record(self, ids: List[int]) -> None:
        if not ids:
            return
        with self._lock:
            self._write_ledger(self._read_ledger() + ids)

    def _cleanup_previous_run(self) -> None:
        """Once per run, before its first send: delete the earlier runs' disposable messages.

        The ledger is cleared whatever the delete returns - Telegram skips ids it
        can no longer delete (older than 48h), so retrying them would never help.
        """
        if self._ledger_path is None:
            return
        with self._lock:
            if self._cleaned:
                return
            self._cleaned = True
            ids = self._read_ledger()
            if not ids:
                return
            result = self._exec(self.delete_command(ids), timeout=60)
            if result is not None and result.returncode == 0:
                logger.info("🧹 deleted %d message(s) from the previous run", len(ids))
            else:
                logger.warning("⚠️ could not delete %d old message(s); dropping them from the ledger", len(ids))
            self._write_ledger([])

    def send(self, text: str, disposable: bool = False) -> bool:
        return self._run(self.command(text), timeout=60, disposable=disposable)

    def send_file(self, text: str, path: str, disposable: bool = False) -> bool:
        return self._run(self.file_command(text, path), timeout=120, disposable=disposable)

    def send_files(self, text: str, paths: List[str], disposable: bool = False) -> bool:
        return self._run(self.files_command(text, paths), timeout=180, disposable=disposable)

    def ping(self, text: str) -> None:
        """Queue a short disposable message without waiting for it to send.

        One background thread sends them in order, at most one Telegram message
        per `PING_MIN_INTERVAL_SECONDS`; pings that queue up meanwhile go out
        together as one message, so none is dropped and the caller's loop never
        waits on the network.
        """
        if self._ping_thread is None:
            self._ping_thread = threading.Thread(target=self._ping_worker, name="parking-pings", daemon=True)
            self._ping_thread.start()
        self._pings.put(text)

    def _ping_worker(self) -> None:
        done = False
        while not done:
            first = self._pings.get()
            if first is None:
                return
            time.sleep(max(0.0, self._last_sent + PING_MIN_INTERVAL_SECONDS - time.monotonic()))
            lines = [first]
            while True:
                try:
                    item = self._pings.get_nowait()
                except queue.Empty:
                    break
                if item is None:
                    done = True
                    break
                lines.append(item)
            self.send("\n".join(lines), disposable=True)

    def close(self) -> None:
        """Wait (bounded) for queued pings to go out before the process exits."""
        if self._ping_thread is None:
            return
        self._pings.put(None)
        self._ping_thread.join(PING_FLUSH_TIMEOUT_SECONDS)
        if self._ping_thread.is_alive():
            logger.warning("⚠️ pings still sending after %.0fs; exiting anyway", PING_FLUSH_TIMEOUT_SECONDS)


class LogNotifier:
    """Sends nothing: every message only goes to the log (the burst rehearsal, which must stay out of the chat)."""

    def send(self, text: str, disposable: bool = False) -> bool:
        logger.info("ℹ️ [not sent] %s", text)
        return True

    def send_file(self, text: str, path: str, disposable: bool = False) -> bool:
        return self.send(text)

    def send_files(self, text: str, paths: List[str], disposable: bool = False) -> bool:
        return self.send(text)

    def ping(self, text: str) -> None:
        self.send(text)

    def close(self) -> None:
        pass


def parse_ids(stdout: str) -> List[int]:
    """The JSON id list `notify_send.py --print-ids` prints as its last stdout line."""
    for line in reversed((stdout or "").strip().splitlines()):
        try:
            ids = json.loads(line)
        except ValueError:
            continue
        if isinstance(ids, list):
            return [i for i in ids if isinstance(i, int)]
    return []
