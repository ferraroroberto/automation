"""Build the weekly Google Photos album, share it by link and draft the email (Playwright, no LLM).

    python -m google.photos_weekly_album --login        # once: sign the dedicated profile in by hand
    python -m google.photos_weekly_album --dry-run      # week, title, per-day counts; creates nothing
    python -m google.photos_weekly_album                # album + share link + Gmail draft
    python -m google.photos_weekly_album --send         # ... and send this week's draft
    python -m google.photos_weekly_album --from 2026-10-03 --to 2026-10-09 [--dry-run]

Sharing and the draft happen only when the config lists email recipients. Drafting never sends;
``--send`` (or ``"send": true`` in the config) sends only the draft carrying this week's subject
and share link.

Run from the repo root with the repo's .venv. The week is the most recent complete Saturday
through Friday; run it on Saturday morning. Screenshots and screen recordings are left out.
Google's Photos APIs can't list the library or share since 2025, hence the browser. Flow, selectors
and failure modes: ``google/photos_weekly_album.md``. Config: ``google/photos_weekly_album.json``
(gitignored; copy the ``.sample``). Every run appends one JSON line to the local run log.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import string
import sys
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin

from browser_stealth import launch_persistent
from google.gmail_web import EMAIL_RE, GmailNotSignedIn, GmailWeb

logger = logging.getLogger("photos_weekly_album")

HERE = Path(__file__).resolve().parent
DEFAULT_CONFIG = HERE / "photos_weekly_album.json"
DEFAULT_LOG = HERE / "photos_weekly_album.log"
DEFAULT_PROFILE = Path.home() / ".photos-weekly-album-profile"

PHOTOS_URL = "https://photos.google.com/"
ALBUMS_URL = "https://photos.google.com/albums"
SCREENSHOTS_LINK = "Screenshots & recordings"

SATURDAY = 5  # date.weekday()
LABEL_RE = re.compile(r"^(Photo|Video)\b.* - ([A-Z][a-z]{2} \d{1,2}, \d{4}, \d{1,2}:\d{2}:\d{2} [AP]M)$")
COUNT_RE = re.compile(r"^(\d+) items?\b")
SHARE_LINK_RE = re.compile(r"https://(?:photos\.app\.goo\.gl/[A-Za-z0-9]+"
                           r"|photos\.google\.com/share/[A-Za-z0-9_-]+\?key=[A-Za-z0-9_-]+)")
EMAIL_FIELDS = ("title", "link", "week_start", "week_end")
WEEK_FIELDS = ("title", "week_start", "week_end")  # at least one keeps each week's subject distinct

# Scan the rendered part of a day-grouped grid (library timeline, collection or album). Day
# headers are dated by the first tile after them (their own text is relative: "Today",
# "Wednesday"). Elements get data-pwa-* marks so Playwright can click them with real input.
SCAN_JS = r"""() => {
  const HEAD = '[role=checkbox][aria-label^="Select all photos from "]';
  document.querySelectorAll('[data-pwa-head]').forEach(e => e.removeAttribute('data-pwa-head'));
  const nodes = [...document.querySelectorAll(HEAD + ', a[href*="/photo/"]')];
  const headers = [], tiles = [];
  let pending = [];
  for (const n of nodes) {
    if (n.tagName !== 'A') { pending.push(n); continue; }
    const label = n.getAttribute('aria-label') || '';
    const id = (n.getAttribute('href').split('/photo/')[1] || '').split(/[?#/]/)[0];
    for (const h of pending) {
      h.setAttribute('data-pwa-head', id);
      headers.push({key: id, label, checked: h.getAttribute('aria-checked') === 'true'});
    }
    pending = [];
    let box = null;
    for (let p = n.parentElement, i = 0; p && i < 5 && !box; p = p.parentElement, i++) {
      box = p.querySelector('[role=checkbox]:not(' + HEAD + ')');
    }
    if (box && box.getAttribute('aria-label') !== label) box = null;
    if (box) box.setAttribute('data-pwa-tile', id);
    tiles.push({id, label, checked: box ? box.getAttribute('aria-checked') === 'true' : null, box: !!box});
  }
  return {headers, tiles, visible: document.visibilityState === 'visible'};
}"""

# The grid scrolls inside its largest c-wiz, not the window (the album list may use either).
SCROLL_POS_JS = r"""() => {
  const sc = [...document.querySelectorAll('c-wiz')].sort((a, b) => b.scrollHeight - a.scrollHeight)[0];
  return Math.max(sc ? sc.scrollTop : 0, window.scrollY);
}"""

COUNTER_JS = r"""() => {
  for (const e of document.querySelectorAll('div, span')) {
    if (e.childElementCount) continue;
    const m = (e.textContent || '').trim().match(/^(\d+) selected$/);
    if (m && e.offsetParent !== null) return Number(m[1]);
  }
  return 0;
}"""

SCROLL_TOP_JS = r"""(px) => {
  const sc = [...document.querySelectorAll('c-wiz')].sort((a, b) => b.scrollHeight - a.scrollHeight)[0];
  if (sc) sc.scrollTop = px;
}"""

HEAD_NEAR_TOP_JS = r"""(key) => {
  const el = document.querySelector(`[data-pwa-head="${key}"]`);
  const sc = [...document.querySelectorAll('c-wiz')].sort((a, b) => b.scrollHeight - a.scrollHeight)[0];
  if (!el || !sc) return false;
  el.scrollIntoView({block: 'start'});
  sc.scrollTop = Math.max(0, sc.scrollTop - 150);
  return true;
}"""

# A link-shared album's card links to /share/<id>?key=... instead of /album/<id>. Albums others
# shared with the account are /share/ cards too; a title match there refuses creation (fails safe).
ALBUM_CARDS_JS = r"""() => [...document.querySelectorAll('a[href*="/album/"], a[href*="/share/"]')].map(a => ({
  href: a.href, lines: (a.innerText || '').split('\n').map(s => s.trim()).filter(Boolean)}))"""

# The visible dialogs, in DOM order (the Share dialog, then the "Create link to share" confirmation
# stacked over it). Buttons are marked data-pwa-share="<dialog>:<label>" for a real click. ``texts``
# is everywhere a link can show: the text, the inputs and the markup (an existing link is only in
# the markup).
SHARE_DIALOGS_JS = r"""() => {
  const shown = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  document.querySelectorAll('[data-pwa-share]').forEach(e => e.removeAttribute('data-pwa-share'));
  return [...document.querySelectorAll('[role=dialog]')].filter(shown).map((d, i) => ({
    heading: (d.innerText || '').trim().split('\n')[0],
    buttons: [...d.querySelectorAll('button, [role=button]')].filter(shown).map(b => {
      const label = (b.innerText || '').trim() || b.getAttribute('aria-label') || '';
      b.setAttribute('data-pwa-share', i + ':' + label);
      return label;
    }),
    texts: [d.innerText || '', ...[...d.querySelectorAll('input')].map(x => x.value), d.innerHTML]}));
}"""
CREATE_LINK, COPY_LINK = "Create link", "Copy link"
CONFIRM_HEADING = "Create link to share"
SHARED_URL_RE = re.compile(r"/share/[^?]+\?key=")


class FlowError(RuntimeError):
    """The page did not behave as the flow expects; nothing was created past this point."""


class NotSignedIn(FlowError):
    pass


class WindowHidden(FlowError):
    """Chrome is minimised or covered by a locked screen, so Photos won't render the grid."""


# How a failed run is told apart: the run record's ``failure`` and the process exit code.
FAILURE_EXIT = {"not-signed-in": 3, "window-hidden": 4, "page-timeout": 5, "error": 1}


def failure_kind(exc: Exception) -> str:
    """``window-hidden``, ``page-timeout`` (a Playwright wait ran out: a selector or page changed) or ``error``."""
    if isinstance(exc, WindowHidden):
        return "window-hidden"
    if type(exc).__name__ == "TimeoutError" and type(exc).__module__.startswith("playwright"):
        return "page-timeout"
    return "error"


# -- pure logic (unit-tested) -------------------------------------------------------------------

def week_range(today: date) -> tuple[date, date]:
    """The most recent complete Saturday-Friday week before ``today`` (a Friday is still open)."""
    end = today - timedelta(days=(today.weekday() - 4) % 7 or 7)
    return end - timedelta(days=6), end


def counter_value(value: int, base_week_start: date, week_start: date) -> int:
    """A counter that goes up by one per Saturday-Friday week from its base week."""
    if base_week_start.weekday() != SATURDAY:
        raise ValueError(f"counter base week_start {base_week_start} is not a Saturday")
    return value + (week_start - base_week_start).days // 7


def build_title(template: str, counters: dict, week_start: date) -> str:
    """``template`` with each ``{name}`` replaced by that counter's value for the week."""
    values = {name: counter_value(int(spec["value"]), date.fromisoformat(spec["week_start"]), week_start)
              for name, spec in counters.items()}
    try:
        return template.format(**values)
    except (KeyError, IndexError) as exc:
        raise ValueError(f"title template placeholder {exc} has no counter in config") from exc


def label_datetime(label: str) -> Optional[datetime]:
    """Date taken from a tile label like ``Photo - Landscape - Oct 3, 2026, 8:23:55 AM``."""
    match = LABEL_RE.match(" ".join(label.split()))  # also folds the narrow no-break space before AM/PM
    return datetime.strptime(match.group(2), "%b %d, %Y, %I:%M:%S %p") if match else None


def label_date(label: str) -> Optional[date]:
    taken = label_datetime(label)
    return taken.date() if taken else None


def days_of(start: date, end: date) -> list[date]:
    return [start + timedelta(days=n) for n in range((end - start).days + 1)]


def find_share_link(texts: list[str]) -> Optional[str]:
    """The first album share link (short or full form) in ``texts``."""
    for text in texts:
        match = SHARE_LINK_RE.search(text or "")
        if match:
            return match.group(0)
    return None


def _fields(template: str) -> set[str]:
    return {name for _, name, _, _ in string.Formatter().parse(template) if name}


@dataclass
class EmailConfig:
    to: list[str]
    subject: str
    body: str
    send: bool = False

    @classmethod
    def load(cls, raw: Optional[dict]) -> Optional["EmailConfig"]:
        """None when no recipients are configured: then nothing is shared or drafted."""
        if not raw or not raw.get("to"):
            return None
        to = [a.strip() for a in raw["to"]]
        if any(not EMAIL_RE.fullmatch(a) for a in to):
            raise ValueError("email.to must list plain addresses (name@example.com)")
        if len({a.lower() for a in to}) != len(to):
            raise ValueError("email.to lists an address twice")
        subject, body = raw.get("subject") or "", raw.get("body") or ""
        try:
            unknown = (_fields(subject) | _fields(body)) - set(EMAIL_FIELDS)
        except ValueError as exc:
            raise ValueError(f"email template is malformed: {exc}") from exc
        if unknown:
            raise ValueError(f"email templates use unknown placeholders {sorted(unknown)}; "
                             f"allowed: {', '.join(EMAIL_FIELDS)}")
        if not _fields(subject) & set(WEEK_FIELDS):
            raise ValueError("email.subject needs {title}, {week_start} or {week_end}, so each week's "
                             "draft is found by its own subject")
        if "link" not in _fields(body):
            raise ValueError("email.body needs a {link} placeholder")
        return cls(to=to, subject=subject, body=body, send=bool(raw.get("send", False)))

    def render(self, title: str, link: str, start: date, end: date) -> tuple[str, str]:
        """(subject, body) for the week; the subject is folded to one line."""
        values = {"title": title, "link": link, "week_start": start.isoformat(), "week_end": end.isoformat()}
        return " ".join(self.subject.format(**values).split()), self.body.format(**values)


@dataclass
class Config:
    title_template: str
    counters: dict
    profile_dir: Path
    log_file: Path
    email: Optional[EmailConfig] = None

    @classmethod
    def load(cls, path: Path) -> "Config":
        if not path.is_file():
            raise ValueError(f"config not found: {path} (copy {path.name}.sample and fill it in)")
        raw = json.loads(path.read_text(encoding="utf-8"))
        if not raw.get("title_template") or not raw.get("counters"):
            raise ValueError("config needs title_template and counters")
        return cls(title_template=raw["title_template"], counters=raw["counters"],
                   profile_dir=Path(raw.get("profile_dir") or DEFAULT_PROFILE).expanduser(),
                   log_file=Path(raw.get("log_file") or DEFAULT_LOG).expanduser(),
                   email=EmailConfig.load(raw.get("email")))


@dataclass
class Selection:
    per_day: dict[date, int] = field(default_factory=dict)
    week_tiles: int = 0
    excluded: int = 0
    excluded_missing: int = 0
    selected: int = 0

    @property
    def day_total(self) -> int:
        return sum(self.per_day.values())


# -- browser flow ---------------------------------------------------------------------------------

class PhotosWeb:
    """One headed Chrome session on the dedicated profile. Use as a context manager."""

    def __init__(self, profile_dir: Path) -> None:
        self.profile_dir = profile_dir
        self._pw = None
        self._context = None
        self.page = None

    def __enter__(self) -> "PhotosWeb":
        from playwright.sync_api import sync_playwright
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        self._pw = sync_playwright().start()
        # Hidden or covered windows stop rendering the grid; keep Chrome from treating a covered
        # window as hidden. A minimised one still is, so the flow checks visibility itself.
        self._context = launch_persistent(self._pw, str(self.profile_dir), headless=False,
                                          disabled_features=("CalculateNativeWinOcclusion",))
        self.page = self._context.pages[0] if self._context.pages else self._context.new_page()
        return self

    def __exit__(self, *exc) -> None:
        try:
            if self._context:
                self._context.close()
        finally:
            if self._pw:
                self._pw.stop()

    # -- navigation & waits -------------------------------------------------------------------

    def _open(self, url: str) -> None:
        self.page.goto(url, wait_until="domcontentloaded")
        self.page.bring_to_front()
        self.page.wait_for_timeout(2_500)

    def _scan(self) -> dict:
        scan = self.page.evaluate(SCAN_JS)
        if not scan["visible"]:
            self.page.bring_to_front()
            self.page.wait_for_timeout(2_000)
            scan = self.page.evaluate(SCAN_JS)
            if not scan["visible"]:
                raise WindowHidden("the Chrome window is hidden or minimised; Google Photos only renders "
                                   "the grid in a visible window - keep it on screen while this runs")
        return scan

    def _wheel(self) -> bool:
        """Scroll the grid one step with the real mouse wheel; False at the end of the list."""
        before = self.page.evaluate(SCROLL_POS_JS)
        self.page.mouse.move(640, 520)
        self.page.mouse.wheel(0, 650)
        self.page.wait_for_timeout(700)
        return self.page.evaluate(SCROLL_POS_JS) > before

    def _counter(self) -> int:
        return int(self.page.evaluate(COUNTER_JS))

    # -- session ------------------------------------------------------------------------------

    def _signed_in(self) -> bool:
        url = self.page.url
        return url.startswith(PHOTOS_URL) and "/about" not in url

    def login(self, wait_seconds: int = 600) -> bool:
        """Wait for the owner to sign in by hand; True once the library has loaded."""
        self._open(PHOTOS_URL)
        deadline = time.monotonic() + wait_seconds
        while time.monotonic() < deadline:
            if self._signed_in():
                return True
            self.page.wait_for_timeout(2_000)
        return False

    def require_signed_in(self) -> None:
        self._open(PHOTOS_URL)
        if not self._signed_in():
            raise NotSignedIn("the dedicated Chrome profile is not signed in to Google Photos")

    # -- albums -------------------------------------------------------------------------------

    def find_album(self, title: str) -> Optional[tuple[str, Optional[int]]]:
        """(url, item count) of an album titled ``title`` in the album list, else None."""
        self._open(ALBUMS_URL)
        wanted = title.strip().casefold()
        idle = 0
        while idle < 3:  # the list loads lazily; stop after three scrolls that add nothing
            cards = self.page.evaluate(ALBUM_CARDS_JS)
            for card in cards:
                lines = card["lines"]
                if lines and lines[0].casefold() == wanted:
                    count = next((int(m.group(1)) for m in map(COUNT_RE.match, lines[1:]) if m), None)
                    return card["href"], count
            moved = self._wheel()
            idle = 0 if moved or len(self.page.evaluate(ALBUM_CARDS_JS)) > len(cards) else idle + 1
        return None

    def screenshot_labels(self, start: date, end: date) -> dict[str, str]:
        """{tile id: label} of the week's items in the "Screenshots & recordings" collection."""
        self._open(ALBUMS_URL)
        link = self.page.locator("a", has_text=SCREENSHOTS_LINK).first
        try:
            href = link.get_attribute("href", timeout=15_000)
        except Exception as exc:
            raise FlowError(f"no '{SCREENSHOTS_LINK}' collection link on the albums page") from exc
        self._open(urljoin(self.page.url, href))
        found: dict[str, str] = {}
        while True:
            scan = self._scan()
            dates = [d for d in (label_date(t["label"]) for t in scan["tiles"]) if d]
            for tile in scan["tiles"]:
                taken = label_date(tile["label"])
                if taken and start <= taken <= end:
                    found[tile["id"]] = tile["label"]
            if (dates and min(dates) < start) or not self._wheel():
                return found

    # -- selection ----------------------------------------------------------------------------

    def _click_checked(self, selector: str, reveal: str, want: bool, what: str) -> None:
        """Real click on a checkbox, asserting the new state; one retry (clicks can be dropped).

        A day checkbox has no size until the pointer is over its header row, and Playwright won't
        move the mouse to an invisible element, so ``reveal`` is hovered first when needed. (Once
        something is selected, tile checkboxes are shown and cover the tile, so no hover there.)
        """
        state = "true" if want else "false"
        for _ in range(2):
            box = self.page.locator(selector).first
            if not box.is_visible():
                self.page.locator(reveal).first.hover()
            box.click()
            try:  # aria-checked follows the click by a moment
                self.page.locator(f'{selector}[aria-checked="{state}"]').first.wait_for(timeout=3_000)
                return
            except Exception:
                continue
        raise FlowError(f"{what}: the checkbox did not change after two clicks")

    def select_week(self, start: date, end: date) -> dict[date, int]:
        """Tick each day header of the week in the library timeline; {day: items it added}."""
        self._open(PHOTOS_URL)
        if self._counter():
            raise FlowError("something is already selected in the library; clear it and retry")
        per_day: dict[date, int] = {}
        while True:
            scan = self._scan()
            target = next((h for h in scan["headers"]
                           if (d := label_date(h["label"])) and start <= d <= end and d not in per_day), None)
            if target:
                day = label_date(target["label"])
                before = self._counter()
                self.page.evaluate(HEAD_NEAR_TOP_JS, target["key"])
                self.page.wait_for_timeout(500)
                self.page.evaluate(SCAN_JS)  # re-mark: the scroll may have re-rendered the header
                head = f'[data-pwa-head="{target["key"]}"]'
                self._click_checked(head, f"{head} >> xpath=..", True, f"day {day}")
                self.page.wait_for_timeout(500)  # let the "N selected" counter catch up
                per_day[day] = self._counter() - before
                logger.info("ℹ️ %s %s: %d items", day.strftime("%a"), day, per_day[day])
                continue
            dates = [d for d in (label_date(t["label"]) for t in scan["tiles"]) if d]
            if (dates and min(dates) < start) or not self._wheel():
                break
        return {day: per_day.get(day, 0) for day in days_of(start, end)}

    def drop_excluded(self, start: date, end: date, excluded: dict[str, str]) -> tuple[int, int, int]:
        """Untick the week's tiles whose label matches an excluded item.

        Returns (week tiles seen, tiles unticked, excluded items not found in the timeline).
        The collection and the timeline use different tile ids, so tiles are matched by label:
        type, orientation and the second the item was taken. More timeline tiles with a label
        than excluded items carrying it can't be told apart, so that aborts.
        """
        wanted = Counter(excluded.values())
        self.page.evaluate(SCROLL_TOP_JS, 0)
        self.page.wait_for_timeout(1_000)
        week: dict[str, str] = {}
        dropped: dict[str, str] = {}
        while True:
            scan = self._scan()
            for tile in scan["tiles"]:
                taken = label_date(tile["label"])
                if not taken or not start <= taken <= end:
                    continue
                week[tile["id"]] = tile["label"]
                if tile["label"] not in wanted or tile["id"] in dropped:
                    continue
                if Counter(dropped.values())[tile["label"]] >= wanted[tile["label"]]:
                    raise FlowError("more items in the week share a timestamp label than there are "
                                    "screenshots with it; can't tell them apart - remove them by hand")
                if not tile["box"]:
                    raise FlowError("a screenshot tile has no checkbox to untick")
                if tile["checked"]:
                    self._click_checked(f'[data-pwa-tile="{tile["id"]}"]', f'a[href*="/photo/{tile["id"]}"]',
                                        False, "screenshot tile")
                dropped[tile["id"]] = tile["label"]
            dates = [d for d in (label_date(t["label"]) for t in scan["tiles"]) if d]
            if (dates and min(dates) < start) or not self._wheel():
                break
        return len(week), len(dropped), sum((wanted - Counter(dropped.values())).values())

    def clear_selection(self) -> None:
        self.page.keyboard.press("Escape")
        self.page.wait_for_timeout(800)
        if self._counter():
            raise FlowError("could not clear the selection")

    def create_album(self, title: str) -> str:
        """Selected items -> new album titled ``title``; returns the album URL."""
        self.page.get_by_role("button", name=re.compile(r"add to album", re.I)).first.click()
        self.page.get_by_role("menuitem", name=re.compile(r"^Album$", re.I)).first.click()
        dialog = self.page.get_by_role("dialog")
        dialog.get_by_text("New album", exact=True).first.click()
        name = self.page.get_by_role("textbox", name="Edit album name")
        name.wait_for(state="visible", timeout=30_000)
        name.click()
        name.press_sequentially(title, delay=60)
        self.page.wait_for_timeout(500)
        self.page.get_by_role("button", name=re.compile(r"^(Done|Save)$", re.I)).first.click()
        self.page.wait_for_url(re.compile(r"/album/"), timeout=30_000)
        self.page.wait_for_timeout(3_000)
        return self.page.url

    def _open_album(self, album_url: str) -> None:
        """Open an album and let a link-shared one finish redirecting to its ``/share/`` URL (a
        click on Share during the redirect is lost)."""
        self._open(album_url)
        try:
            self.page.wait_for_url(SHARED_URL_RE, timeout=5_000)
        except Exception:  # not link-shared: it stays at /album/
            pass

    def _share_dialogs(self, done, timeout_ms: int) -> list[dict]:
        """Snapshots of the visible dialogs, polled until ``done(dialogs)`` or the timeout."""
        for _ in range(max(1, timeout_ms // 500)):
            dialogs = self.page.evaluate(SHARE_DIALOGS_JS)
            if done(dialogs):
                break
            self.page.wait_for_timeout(500)
        return dialogs

    def _click_share_button(self, dialogs: list[dict], label: str, heading: Optional[str] = None) -> bool:
        """Click the button ``label`` in the topmost dialog that has one (and ``heading``, if given)."""
        for index in reversed(range(len(dialogs))):
            if label in dialogs[index]["buttons"] and heading in (None, dialogs[index]["heading"]):
                self.page.locator(f'[data-pwa-share="{index}:{label}"]').first.click()
                return True
        return False

    def share_link(self, album_url: str) -> str:
        """The album's share link: the existing one, else one made with "Create link".

        An album has at most one link, and once it exists its Share dialog offers "Copy link"
        instead of "Create link", so a re-run never makes a second one. "Create link" opens a
        second dialog, "Create link to share", whose own "Create link" button makes the link and
        then shows it. The link is taken from the first place it shows up: the dialogs (the
        confirmation's input, or the Share dialog's markup once a link exists), a response from
        Google, the page URL (a link-shared album opens at ``/share/<id>?key=...``), or the
        clipboard after "Copy link" (this overwrites the clipboard).
        """
        self._open_album(album_url)
        self._context.grant_permissions(["clipboard-read", "clipboard-write"], origin=PHOTOS_URL.rstrip("/"))
        seen: list[str] = []
        tried = ["dialogs", "responses", "page URL"]

        def on_response(response) -> None:
            if response.request.resource_type in ("xhr", "fetch"):
                try:
                    seen.append(response.text())
                except Exception:  # bodies of redirects and aborted requests can't be read
                    pass

        def read(dialogs: list[dict]) -> Optional[str]:
            return find_share_link([*(t for d in dialogs for t in d["texts"]), *seen, self.page.url])

        def offers(*labels: str):
            return lambda dialogs: any(label in d["buttons"] for d in dialogs for label in labels)

        created = False
        self.page.on("response", on_response)
        try:
            self.page.get_by_role("button", name=re.compile(r"^Share$", re.I)).first.click()
            dialogs = self._share_dialogs(offers(CREATE_LINK, COPY_LINK), 20_000)
            has_link = offers(COPY_LINK)(dialogs)
            if not has_link and self._click_share_button(dialogs, CREATE_LINK):
                dialogs = self._share_dialogs(lambda ds: any(d["heading"] == CONFIRM_HEADING and CREATE_LINK
                                                             in d["buttons"] for d in ds), 10_000)
                created = self._click_share_button(dialogs, CREATE_LINK, CONFIRM_HEADING)
                if created:
                    logger.info("ℹ️ share link created")
                    dialogs = self._share_dialogs(read, 30_000)
            link = read(dialogs)
            if not link and self._click_share_button(dialogs, COPY_LINK):
                tried.append("clipboard")
                self.page.wait_for_timeout(1_500)
                clipboard = self.page.evaluate("() => navigator.clipboard.readText().catch(() => '')")
                link = find_share_link([clipboard])
        finally:
            self.page.remove_listener("response", on_response)
            self.page.keyboard.press("Escape")
        if not link:  # a freshly shared album may only open at its /share/ URL once reloaded
            tried.append("reloaded page URL")
            self._open_album(album_url)
            link = find_share_link([self.page.url])
        if not link:
            logger.warning("⚠️ share link not found; tried: %s; dialog buttons: %s", ", ".join(tried),
                           " / ".join(f'{d["heading"]!r}: {d["buttons"]}' for d in dialogs) or "no dialog open")
            state = ("the album is now shared by link, but the script could not read it" if created else
                     "the album has a share link, but the script could not read it" if has_link else
                     "the album's Share dialog gave no link to create or copy")
            raise FlowError(f"{state}; re-run, or copy it from the album's Share dialog (its buttons are in the log)")
        return link

    def collect(self, start: date, end: date) -> Selection:
        """Select the week minus screenshots and check the counts add up."""
        excluded = self.screenshot_labels(start, end)
        logger.info("ℹ️ %d screenshots/recordings in the week to leave out", len(excluded))
        result = Selection(per_day=self.select_week(start, end))
        result.week_tiles, result.excluded, result.excluded_missing = self.drop_excluded(start, end, excluded)
        result.selected = self._counter()
        if result.week_tiles != result.day_total:
            raise FlowError(f"saw {result.week_tiles} tiles in the week but the day headers selected "
                            f"{result.day_total}; the sweep missed items, so screenshots may remain")
        if result.selected != result.day_total - result.excluded:
            raise FlowError(f"selection counter reads {result.selected}, expected "
                            f"{result.day_total} - {result.excluded} screenshots")
        return result


# -- run ------------------------------------------------------------------------------------------

def append_run_log(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")


def report(start: date, end: date, title: str, sel: Selection) -> None:
    logger.info("ℹ️ week %s (Sat) .. %s (Fri) -> title: %s", start, end, title)
    for day, count in sel.per_day.items():
        logger.info("ℹ️   %s %s  %4d", day.strftime("%a"), day, count)
    logger.info("ℹ️ day total %d - screenshots %d = %d selected%s", sel.day_total, sel.excluded,
                sel.selected, f" ({sel.excluded_missing} screenshots not in the timeline)"
                if sel.excluded_missing else "")


def deliver(mail, to: list[str], subject: str, body: str, link: str, send: bool) -> str:
    """Make sure this week's message exists as exactly one draft, or was sent; send it if asked.

    Returns the mail outcome: ``drafted``, ``draft-exists``, ``sent`` or ``already-sent``. The
    week's subject is the key: a message already in Sent is never drafted or sent again, and an
    existing draft is reused, never duplicated. ``send`` only ever sends that one draft.
    """
    if mail.subjects("sent", subject):
        logger.info("✅ this week's message is already in Sent; no draft, nothing sent")
        return "already-sent"
    drafts = len(mail.subjects("draft", subject))
    if drafts > 1:
        raise FlowError(f"{drafts} drafts carry this week's subject; delete the extras by hand and re-run")
    if drafts:
        logger.info("ℹ️ this week's draft already exists; not drafting another")
        outcome = "draft-exists"
    else:
        mail.create_draft(to, subject, body)
        if len(mail.subjects("draft", subject, expect=True)) != 1:
            raise FlowError("the draft was saved but isn't in Drafts under this week's subject")
        logger.info("✅ draft saved for %d recipients", len(to))
        outcome = "drafted"
    if not send:
        return outcome
    mail.send_draft(subject, to, link)
    if not mail.subjects("sent", subject, expect=True):
        raise FlowError("Send was clicked but the message isn't in Sent; check Gmail by hand")
    logger.info("✅ sent to %d recipients", len(to))
    return "sent"


def share_and_mail(web: PhotosWeb, email: EmailConfig, record: dict, album_url: str, title: str,
                   start: date, end: date, send: bool) -> None:
    """Step 3: share link for this week's album, then its Gmail draft (and send, if asked)."""
    link = web.share_link(album_url)
    record["share_url"] = link
    logger.info("✅ share link: %s", link)
    subject, body = email.render(title, link, start, end)
    gmail = GmailWeb(web.page)
    gmail.require_signed_in()
    record["mail"] = deliver(gmail, email.to, subject, body, link, send)


def run(cfg: Config, start: date, end: date, title: str, dry_run: bool, send: bool = False,
        record: Optional[dict] = None) -> int:
    """One run; returns the exit code. A caller that wants to read the run log's record passes a dict."""
    record = {} if record is None else record
    record.update({"at": datetime.now().isoformat(timespec="seconds"), "week_start": start, "week_end": end,
                   "title": title, "dry_run": dry_run})
    try:
        with PhotosWeb(cfg.profile_dir) as web:
            web.require_signed_in()
            existing = web.find_album(title)
            if existing and not dry_run:
                record.update(outcome="exists", url=existing[0], album_count=existing[1])
                logger.info("✅ album already exists (%s items), not creating another: %s", existing[1], existing[0])
                if not cfg.email:
                    return 0
                share_and_mail(web, cfg.email, record, existing[0], title, start, end, send)
                return 0
            if existing:
                logger.info("ℹ️ an album with this title already exists (%s items): %s; a real run "
                            "would refuse to create another", existing[1], existing[0])
            sel = web.collect(start, end)
            record.update(per_day={d.isoformat(): n for d, n in sel.per_day.items()},
                          excluded=sel.excluded, excluded_missing=sel.excluded_missing, selected=sel.selected)
            report(start, end, title, sel)
            if dry_run:
                web.clear_selection()
                record["outcome"] = "dry-run"
                if cfg.email:
                    subject, _ = cfg.email.render(title, "<share link>", start, end)
                    logger.info("ℹ️ a real run would share the album by link and draft \"%s\" to %d recipients%s",
                                subject, len(cfg.email.to), ", then send it" if send else "")
                else:
                    logger.info("ℹ️ no email recipients in config: a real run would not share or draft")
                return 0
            if not sel.selected:
                web.clear_selection()
                record["outcome"] = "empty"
                logger.warning("⚠️ nothing to put in an album for this week")
                return 0
            url = web.create_album(title)
            count = None
            for _ in range(3):  # the album card's count can lag the creation by a few seconds
                found = web.find_album(title)
                count = found[1] if found else None
                if count == sel.selected:
                    break
                web.page.wait_for_timeout(5_000)
            record.update(url=url, album_count=count)
            if count != sel.selected:
                record["outcome"] = "count-mismatch"
                logger.error("❌ album created but it lists %s items, expected %d: %s", count, sel.selected, url)
                return 1
            record["outcome"] = "created"
            logger.info("✅ album created with %d items: %s", count, url)
            print(url)
            if cfg.email:
                share_and_mail(web, cfg.email, record, url, title, start, end, send)
            else:
                logger.info("ℹ️ no email recipients in config: album not shared, no draft")
            return 0
    except (NotSignedIn, GmailNotSignedIn) as exc:
        record.update(outcome="not-signed-in", failure="not-signed-in")
        logger.error("❌ %s. Sign in once with: .venv\\Scripts\\python.exe -m google.photos_weekly_album --login",
                     exc)
        return FAILURE_EXIT["not-signed-in"]
    except Exception as exc:
        failure = failure_kind(exc)
        record.update(error=str(exc), failure=failure)
        if "outcome" in record:  # the album step finished; sharing or mail failed
            record.setdefault("mail", "error")
        else:
            record["outcome"] = "error"
        logger.error("❌ %s", exc)
        return FAILURE_EXIT[failure]
    finally:
        append_run_log(cfg.log_file, record)


def parse_args(argv: Optional[list]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create the weekly Google Photos album, share it and draft the email.")
    parser.add_argument("--dry-run", action="store_true", help="count and report, create nothing")
    parser.add_argument("--send", action="store_true", help="send this week's draft (only that one)")
    parser.add_argument("--login", action="store_true", help="sign the dedicated profile in by hand")
    parser.add_argument("--from", dest="start", type=date.fromisoformat, help="first day (YYYY-MM-DD)")
    parser.add_argument("--to", dest="end", type=date.fromisoformat, help="last day (YYYY-MM-DD)")
    parser.add_argument("--config", type=Path, default=Path(os.environ.get("PHOTOS_ALBUM_CONFIG", DEFAULT_CONFIG)))
    args = parser.parse_args(argv)
    if (args.start is None) != (args.end is None):
        parser.error("--from and --to go together")
    if args.start and not args.start <= args.end <= args.start + timedelta(days=30):
        parser.error("--to must be on or after --from, at most 30 days later")
    if args.send and args.login:
        parser.error("--send doesn't go with --login")
    return args


def main(argv: Optional[list] = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    try:
        cfg = Config.load(args.config)
    except ValueError as exc:
        logger.error("❌ %s", exc)
        return 2
    if args.login:
        logger.info("ℹ️ A Chrome window opens on the dedicated profile; sign in to Google Photos by hand "
                    "(10 min max). It closes by itself once the library loads.")
        with PhotosWeb(cfg.profile_dir) as web:
            if not web.login():
                logger.error("❌ sign-in did not complete in time")
                return 1
        logger.info("✅ profile signed in: %s", cfg.profile_dir)
        return 0
    start, end = (args.start, args.end) if args.start else week_range(date.today())
    try:
        title = build_title(cfg.title_template, cfg.counters, start)
    except ValueError as exc:
        logger.error("❌ %s", exc)
        return 2
    if args.send and not cfg.email:
        logger.error("❌ --send needs email recipients in the config (email.to)")
        return 2
    return run(cfg, start, end, title, args.dry_run, send=args.send or bool(cfg.email and cfg.email.send))


if __name__ == "__main__":
    raise SystemExit(main())
