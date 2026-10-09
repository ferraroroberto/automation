"""Saturday-morning run of the weekly Photos album, for an app-launcher job (Playwright, no LLM).

    python -m google.photos_weekly_album_job            # run from the repo root; google/run_photos_weekly_album.bat does

Runs ``google.photos_weekly_album`` for the week that just ended (album, then the share link and Gmail
draft when the config lists recipients) and sends one short Telegram message through the fleet
notifier: done, no-op (the album was already there), or a failure that says what to fix. It never
passes ``--send``; mail goes out only if the config says ``"send": true``. Setup, outcomes and the job
entry: ``google/photos_weekly_album.md``.

Exit codes: 0 done / no-op / nothing to album, 1 error or count mismatch, 2 config error,
3 profile signed out, 4 desktop locked or Chrome window hidden, 5 a Google page didn't respond.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Optional

from dotenv import dotenv_values

from google import photos_weekly_album as album
from parking.notify import FleetNotifier

logger = logging.getLogger("photos_weekly_album_job")

REPO_ROOT = Path(__file__).resolve().parent.parent
DESKTOP_SWITCHDESKTOP = 0x0100

# exit code per failure kind (see the module docstring)
EXIT = {"config": 2, "desktop-locked": 4, **album.FAILURE_EXIT, "count-mismatch": 1}
DRAFT_TEXT = {None: "not shared (no recipients configured)", "drafted": "saved, not sent",
              "draft-exists": "already there, not sent", "sent": "sent", "already-sent": "already sent earlier"}
FAILURE_TEXT = {
    "not-signed-in": ("the Chrome profile is signed out of Google.",
                      "Sign in again: .venv\\Scripts\\python.exe -m google.photos_weekly_album --login"),
    "desktop-locked": ("the desktop is locked, or nobody is signed in.",
                       "Unlock it, then run the job from the Jobs tab."),
    "window-hidden": ("the Chrome window was hidden or minimised.",
                      "Keep it on screen, then run the job again."),
    "page-timeout": ("a Google page didn't respond as expected.",
                     "Google may have changed its pages: see the run log."),
    "count-mismatch": ("the album lists a different number of items than were selected.",
                       "Check the album by hand."),
}


@dataclass
class Outcome:
    kind: str  # done | no-op | empty | one of EXIT
    text: str
    exit_code: int

    @property
    def failed(self) -> bool:
        return self.exit_code != 0


def desktop_locked() -> Optional[bool]:
    """True when the input desktop can't be opened (screen locked, or no one signed in); None if unknown."""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        user32 = ctypes.windll.user32
        user32.OpenInputDesktop.restype = ctypes.c_void_p
        user32.OpenInputDesktop.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]
        user32.CloseDesktop.argtypes = [ctypes.c_void_p]
        handle = user32.OpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP)
        if not handle:
            return True
        user32.CloseDesktop(handle)
        return False
    except (OSError, AttributeError) as exc:
        logger.warning("⚠️ could not tell whether the desktop is locked: %s", exc)
        return None


def describe(code: int, record: dict) -> Outcome:
    """The notification text and exit code for one finished run, from its record."""
    title = record.get("title", "?")
    link = record.get("share_url") or record.get("url")
    if code == 0:
        outcome = record.get("outcome")
        if outcome == "empty":
            return Outcome("empty", f"📸 No album this week: nothing to add for {record['week_start']} to "
                                    f"{record['week_end']}.", 0)
        draft = DRAFT_TEXT.get(record.get("mail"), record.get("mail"))
        noop = outcome == "exists" and record.get("mail") in (None, "draft-exists", "already-sent")
        head = "already there" if noop else "done"
        return Outcome("no-op" if noop else "done", f"📸 Weekly album {head}: {title}\n{link}\nDraft: {draft}", 0)
    kind = record.get("failure") or ("count-mismatch" if record.get("outcome") == "count-mismatch" else "error")
    reason, hint = FAILURE_TEXT.get(kind) or (f"{(record.get('error') or 'unexpected error')[:200]}",
                                              "See the run log; re-running is safe.")
    if record.get("outcome") in ("created", "exists") and kind == "error":  # the album is fine, a later step broke
        reason = f"the album is made but the share or draft step failed: {reason}"
    text = f"❌ Weekly album failed: {reason}\n{hint}"
    if record.get("url"):
        text += f"\nAlbum: {record['url']}"
    return Outcome(kind, text, code or EXIT.get(kind, 1))


def notify(env: dict, outcome: Outcome) -> bool:
    """One message: failures to the attention chat, everything else to the log chat."""
    category = "attention" if outcome.failed else "log"
    return FleetNotifier({**env, "NOTIFY_CATEGORY": category}).send(outcome.text)


def execute(config_path: Path, today: date, env: dict, locked: Callable[[], Optional[bool]] = desktop_locked,
            run: Callable[..., int] = album.run) -> int:
    """Config, desktop check, the album run, one notification. Returns the exit code."""
    try:
        cfg = album.Config.load(config_path)
        start, end = album.week_range(today)
        title = album.build_title(cfg.title_template, cfg.counters, start)
    except ValueError as exc:
        logger.error("❌ %s", exc)
        outcome = Outcome("config", f"❌ Weekly album failed: config problem: {str(exc)[:200]}\n"
                                    "Fix google/photos_weekly_album.json.", EXIT["config"])
        notify(env, outcome)
        return outcome.exit_code
    record: dict = {"title": title, "week_start": start, "week_end": end}
    if locked():
        record.update(outcome="desktop-locked", failure="desktop-locked", dry_run=False,
                      at=datetime.now().isoformat(timespec="seconds"))
        album.append_run_log(cfg.log_file, record)
        logger.error("❌ the desktop is locked; Chrome would not render Google Photos. Nothing was started.")
        code = EXIT["desktop-locked"]
    else:
        # sending follows the config alone: this job never passes --send
        code = run(cfg, start, end, title, False, send=bool(cfg.email and cfg.email.send), record=record)
    outcome = describe(code, record)
    logger.info("ℹ️ outcome: %s", outcome.kind)
    if not notify(env, outcome):
        logger.warning("⚠️ the notification was not delivered (NOTIFY_PYTHON / NOTIFY_SCRIPT in .env?)")
    return outcome.exit_code


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Run the weekly Photos album and notify the outcome.")
    parser.add_argument("--config", type=Path,
                        default=Path(os.environ.get("PHOTOS_ALBUM_CONFIG", album.DEFAULT_CONFIG)))
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    env = {**dotenv_values(REPO_ROOT / ".env"), **{k: v for k, v in os.environ.items() if k.startswith("NOTIFY_")}}
    try:
        return execute(args.config, date.today(), env)
    except Exception as exc:  # last resort: a scheduled run must not end without a word
        logger.exception("❌ unexpected error")
        notify(env, Outcome("error", f"❌ Weekly album failed: unexpected error: {str(exc)[:200]}\n"
                                     "See the run log; re-running is safe.", 1))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
