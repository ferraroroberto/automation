"""Saturday-morning run of the weekly Photos album, for an app-launcher job (Playwright, no LLM).

    python -m tools.google.photos_weekly_album_job            # run from the repo root; tools/google/run_photos_weekly_album.bat does
    python -m tools.google.photos_weekly_album_job --dry-run  # the same entry point, creates nothing, sends no message

Runs ``tools.google.photos_weekly_album`` for the week that just ended (album, then the share link and Gmail
draft when the config lists recipients), but only after a preflight: the same run in its report-only mode,
whose counts must agree, else the job stops before selecting anything. It sends one short Telegram message through the fleet
notifier: done, no-op (the album was already there), or a failure that says what to fix. It never
passes ``--send``; mail goes out only if the config says ``"send": true``. With ``--dry-run`` the album run
only counts (nothing is created, shared or drafted) and no notification is sent. Setup, outcomes and the job
entry: ``tools/google/photos_weekly_album.md``.

Exit codes: 0 done / no-op / nothing to album, 1 error or count mismatch, 2 config error,
3 profile signed out, 4 desktop locked or Chrome window hidden, 5 a Google page didn't respond,
6 the preflight dry run found a count problem (nothing was created).
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

from tools.google import photos_weekly_album as album
from tools.system._lib.notifier import FleetNotifier

logger = logging.getLogger("photos_weekly_album_job")

REPO_ROOT = Path(__file__).resolve().parents[2]
DESKTOP_SWITCHDESKTOP = 0x0100

# exit code per failure kind (see the module docstring)
EXIT = {"config": 2, "desktop-locked": 4, **album.FAILURE_EXIT, "count-mismatch": 1, "preflight-failed": 6}
DRAFT_TEXT = {None: "not shared (no recipients configured)", "drafted": "saved, not sent",
              "draft-exists": "already there, not sent", "sent": "sent", "already-sent": "already sent earlier"}
FAILURE_TEXT = {
    "not-signed-in": ("the Chrome profile is signed out of Google.",
                      "Sign in again: .venv\\Scripts\\python.exe -m tools.google.photos_weekly_album --login"),
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


def preflight(run: Callable[..., int], cfg: album.Config, start: date, end: date, title: str) -> Optional[Outcome]:
    """The real run's own dry run: None when it ends clean, else the outcome that stops the job.

    A dry run counts the week and clears the selection, so a mismatch is found before anything is
    created. Anything short of exit 0 with outcome ``dry-run`` stops the real pass; a count problem is
    ``preflight-failed``, while a signed-out profile, hidden window or unresponsive page keeps its own
    kind (the same condition would stop the real pass the same way).
    """
    pre: dict = {"title": title, "week_start": start, "week_end": end, "preflight": True}
    code = run(cfg, start, end, title, True, send=False, record=pre)
    if code == 0 and pre.get("outcome") == "dry-run":
        logger.info("✅ preflight clean: %s selected, per day %s, tiles per day %s", pre.get("selected"),
                    pre.get("per_day"), pre.get("tiles_per_day"))
        return None
    if code not in (0, 1):
        logger.error("❌ preflight stopped the run (exit %s): nothing was created", code)
        return describe(code, pre)
    detail = str(pre.get("error") or f"the dry run ended as '{pre.get('outcome')}'")[:300]
    logger.error("❌ preflight stopped the run: %s", detail)
    return Outcome("preflight-failed", f"❌ Weekly album preflight failed: {detail}\n"
                                       "Nothing was created, shared or sent, and the selection was cleared. "
                                       "Check the library by hand.", EXIT["preflight-failed"])


def execute(config_path: Path, today: date, env: dict, locked: Optional[Callable[[], Optional[bool]]] = None,
            run: Optional[Callable[..., int]] = None, dry_run: bool = False) -> int:
    """Config, desktop check, the album run, one notification (none when ``dry_run``). Returns the exit code."""
    locked = locked or desktop_locked  # resolved per call so a test can swap the module's own check
    run = run or album.run
    try:
        cfg = album.Config.load(config_path)
        start, end = album.week_range(today)
        title = album.build_title(cfg.title_template, cfg.counters, start)
    except ValueError as exc:
        logger.error("❌ %s", exc)
        outcome = Outcome("config", f"❌ Weekly album failed: config problem: {str(exc)[:200]}\n"
                                    "Fix tools/google/photos_weekly_album.json.", EXIT["config"])
        if not dry_run:
            notify(env, outcome)
        return outcome.exit_code
    record: dict = {"title": title, "week_start": start, "week_end": end}
    stopped: Optional[Outcome] = None
    if locked():
        record.update(outcome="desktop-locked", failure="desktop-locked", dry_run=dry_run,
                      at=datetime.now().isoformat(timespec="seconds"))
        album.append_run_log(cfg.log_file, record)
        logger.error("❌ the desktop is locked; Chrome would not render Google Photos. Nothing was started.")
        code = EXIT["desktop-locked"]
    else:
        send = bool(cfg.email and cfg.email.send)  # sending follows the config alone: this job never passes --send
        if not dry_run:
            stopped = preflight(run, cfg, start, end, title)
        if stopped:
            code = stopped.exit_code
        else:
            code = run(cfg, start, end, title, dry_run, send=send, record=record)
    outcome = stopped or describe(code, record)
    logger.info("ℹ️ outcome: %s", outcome.kind)
    if dry_run:
        logger.info("ℹ️ dry run: no notification sent")
    elif not notify(env, outcome):
        logger.warning("⚠️ the notification was not delivered (NOTIFY_PYTHON / NOTIFY_SCRIPT in .env?)")
    return outcome.exit_code


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Run the weekly Photos album and notify the outcome.")
    parser.add_argument("--config", type=Path,
                        default=Path(os.environ.get("PHOTOS_ALBUM_CONFIG", album.DEFAULT_CONFIG)))
    parser.add_argument("--dry-run", action="store_true",
                        help="the real entry point, but create, share and send nothing, and send no message")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    env = {**dotenv_values(REPO_ROOT / ".env"), **{k: v for k, v in os.environ.items() if k.startswith("NOTIFY_")}}
    try:
        return execute(args.config, date.today(), env, dry_run=args.dry_run)
    except Exception as exc:  # last resort: a scheduled run must not end without a word
        logger.exception("❌ unexpected error")
        if args.dry_run:
            return 1
        notify(env, Outcome("error", f"❌ Weekly album failed: unexpected error: {str(exc)[:200]}\n"
                                     "See the run log; re-running is safe.", 1))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
