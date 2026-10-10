"""Outlook on the web driven by Playwright: create, verify and (only when asked) send drafts.

Reusable from any script; ``system/outlook_web_mail.py`` is the command-line tool built on it.

    from system.outlook_web import DraftSpec, OutlookWeb

    spec = DraftSpec(to="someone@example.com", subject="report", attachments=(Path("report.pdf"),))
    with OutlookWeb() as web:
        web.require_signed_in()
        web.draft(spec)
        assert not web.verify([spec.subject])
        # web.send(spec)   # explicit, never implied by drafting

The signed-in Chrome profile lives outside the repo (``OUTLOOK_WEB_PROFILE_DIR``, default
``~/.outlook-web-drafts-profile``); sign in once by hand with ``OutlookWeb.login``.
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Iterator, Optional

logger = logging.getLogger("outlook_web")

OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/"
NO_SUBJECT = "(No subject)"
STEALTH_INIT_SCRIPT = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
# The message list is virtualized (about 8 rows render at once, ~80 px each): a step of ~6 rows keeps
# neighbouring snapshots overlapping, and the cap bounds a list that never reports its end.
SCROLL_STEP = 500
MAX_SCROLLS = 200


class NotSignedIn(RuntimeError):
    """The saved profile is not signed in to the mailbox."""


@dataclass(frozen=True)
class DraftSpec:
    to: str
    subject: str
    attachments: tuple[Path, ...] = ()
    body: str = ""

    def __post_init__(self) -> None:
        if not self.to or not self.subject:
            raise ValueError("a draft needs a recipient and a subject")


def clean_row(text: str) -> str:
    """A row's text without icon glyphs (private-use characters) or blank lines: hovering a row adds
    its action icons, which would make two reads of the same row compare different."""
    lines = (re.sub("[\ue000-\uf8ff]", "", line).strip() for line in text.splitlines())
    return "\n".join(line for line in lines if line)


def merge_snapshots(seen: list[str], new: list[str]) -> list[str]:
    """Append the rows of ``new`` that ``seen`` does not already end with.

    Successive snapshots of a scrolled list overlap; the longest overlap between the tail of ``seen``
    and the head of ``new`` is dropped, so identical rows elsewhere in the list are still kept.
    """
    for size in range(min(len(seen), len(new)), 0, -1):
        if seen[-size:] == new[:size]:
            return seen + new[size:]
    return seen + new


def stealth_launch_kwargs(user_data_dir: str, headless: bool) -> dict:
    return {
        "user_data_dir": user_data_dir,
        "channel": "chrome",
        "headless": headless,
        "locale": "en-US",
        "args": ["--disable-blink-features=AutomationControlled", "--disable-features=Translate",
                 "--no-default-browser-check", "--no-first-run", "--lang=en-US"],
        "ignore_default_args": ["--enable-automation", "--enable-blink-features=IdleDetection"],
        "viewport": {"width": 1280, "height": 900},
        "chromium_sandbox": True,
    }


def profile_dir() -> Path:
    return Path(os.environ.get("OUTLOOK_WEB_PROFILE_DIR") or Path.home() / ".outlook-web-drafts-profile")


class OutlookWeb:
    """One Chrome session on the mailbox. Use as a context manager."""

    def __init__(self, url: str = OUTLOOK_URL, headless: bool = False) -> None:
        self.url = url
        self.headless = headless
        self._pw = None
        self._context = None
        self.page = None

    def __enter__(self) -> "OutlookWeb":
        from playwright.sync_api import sync_playwright
        profile_dir().mkdir(parents=True, exist_ok=True)
        self._pw = sync_playwright().start()
        self._context = self._pw.chromium.launch_persistent_context(
            **stealth_launch_kwargs(str(profile_dir()), self.headless))
        self._context.add_init_script(STEALTH_INIT_SCRIPT)
        self.page = self._context.pages[0] if self._context.pages else self._context.new_page()
        self.page.goto(self.url)
        return self

    def __exit__(self, *exc) -> None:
        try:
            if self._context:
                self._context.close()
        finally:
            if self._pw:
                self._pw.stop()

    # -- locators ---------------------------------------------------------------------------

    def _new_mail_button(self):
        return self.page.get_by_role("button", name=re.compile(r"^New mail", re.I)).first

    def _subject_box(self):
        return self.page.get_by_role("textbox", name="Subject")

    def _rows(self):
        return self.page.get_by_role("option")

    def _open_folder(self, name: str) -> None:
        self.page.get_by_role("treeitem", name=re.compile(rf"^{name}", re.I)).first.click()
        self.page.wait_for_timeout(3_000)
        try:  # the list can still be empty here; an empty first read would end a scan before it began
            self._rows().first.wait_for(state="visible", timeout=10_000)
        except Exception:
            logger.info("ℹ️ no rows appeared in %s (empty folder?)", name)

    def _open_drafts(self) -> None:
        self._open_folder("Drafts")

    def _pause(self, ms: int) -> None:
        self.page.wait_for_timeout(ms)

    def _scroll_list(self, delta: int) -> None:
        """Wheel the open message list by ``delta`` px (negative = up). The pointer goes on the middle
        rendered row: the first one can sit above the visible list, where a wheel does nothing."""
        rows = self._rows()
        count = rows.count()
        if count == 0:
            return
        box = rows.nth(count // 2).bounding_box()
        if box:
            self.page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        self.page.mouse.wheel(0, delta)
        self._pause(600)

    def _settled_rows(self) -> list[str]:
        """The rendered rows once two reads in a row agree: a wheel scrolls smoothly, and a read taken
        mid-animation would overlap its neighbours wrongly."""
        texts = [clean_row(t) for t in self._rows().all_inner_texts()]
        for _ in range(8):
            self._pause(300)
            again = [clean_row(t) for t in self._rows().all_inner_texts()]
            if again == texts:
                break
            texts = again
        return texts

    def _scroll_positions(self) -> Iterator[list[str]]:
        """Yield the rendered row texts at each scroll position of the open list, top to bottom.

        Only a handful of rows are in the DOM at once, so a one-shot read misses the rest of a long
        list. The list has ended when a scroll leaves the rendered rows unchanged, even after a
        pause for Outlook to load more.
        """
        self._scroll_list(-1_000_000)
        previous: Optional[list[str]] = None
        for _ in range(MAX_SCROLLS):
            texts = self._settled_rows()
            if texts == previous:
                self._pause(1_500)
                texts = self._settled_rows()
                if texts == previous:
                    return
            yield texts
            previous = texts
            self._scroll_list(SCROLL_STEP)

    def _all_row_texts(self) -> list[str]:
        """Every row of the open list, in order."""
        rows: list[str] = []
        for texts in self._scroll_positions():
            rows = merge_snapshots(rows, texts)
        return rows

    # -- session ----------------------------------------------------------------------------

    def login(self, wait_seconds: int = 600) -> bool:
        """Wait for the owner to sign in by hand (SSO/MFA); True once the mailbox has loaded."""
        try:
            self._new_mail_button().wait_for(state="visible", timeout=wait_seconds * 1_000)
        except Exception:
            return False
        return True

    def require_signed_in(self, wait_seconds: int = 60) -> None:
        try:
            self._new_mail_button().wait_for(state="visible", timeout=wait_seconds * 1_000)
        except Exception as exc:
            raise NotSignedIn("the profile is not signed in; sign in first") from exc

    def screenshot(self, label: str) -> Optional[Path]:
        folder = profile_dir() / "failures"
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / f"{datetime.now():%Y%m%d-%H%M%S}-{label}.png"
        try:
            self.page.screenshot(path=str(target))
            return target
        except Exception:
            return None

    # -- drafting ---------------------------------------------------------------------------

    def _open_compose(self, attempts: int = 3) -> None:
        """The first click after navigating is sometimes swallowed, so retry until the compose shows.

        "New mail" only exists on the Home ribbon, which an open compose replaces, so go via Drafts first.
        """
        self._open_drafts()
        for _ in range(attempts):
            self._new_mail_button().click()
            try:
                self._subject_box().wait_for(state="visible", timeout=8_000)
                return
            except Exception:
                continue
        raise RuntimeError("the compose pane did not open")

    def _attach(self, path: Path) -> None:
        # The ribbon's first file input is the inline-image one (accept=image/*) and refuses a small .txt.
        self.page.locator("input[type=file]:not([accept])").first.set_input_files(str(path))
        copy_choice = self.page.get_by_text("Attach as a copy")
        try:
            copy_choice.wait_for(state="visible", timeout=15_000)
            copy_choice.click()
        except Exception:
            logger.info("ℹ️ no share-as-link prompt for %s (small file)", path.name)
        self.page.get_by_role("option", name=re.compile(re.escape(path.name))).first.wait_for(
            state="visible", timeout=180_000)

    def draft(self, spec: DraftSpec) -> None:
        """Attachments first, then body, recipient, subject: a subject typed before the upload is lost."""
        page = self.page
        self._open_compose()
        for path in spec.attachments:
            self._attach(path)
        if spec.body:
            body = page.get_by_role("textbox", name="Message body")
            body.click()
            page.keyboard.press("Control+Home")
            page.keyboard.insert_text(spec.body + "\n")
        to_box = page.locator('div[aria-label="To"]').first
        to_box.click()
        to_box.press_sequentially(f"{spec.to};", delay=20)
        page.get_by_text(re.compile(re.escape(spec.to), re.I)).first.wait_for(state="visible", timeout=15_000)
        subject = self._subject_box()
        subject.click()
        subject.fill(spec.subject)
        if subject.input_value() != spec.subject:
            raise RuntimeError(f"subject did not take for {spec.subject}")
        page.keyboard.press("Control+s")
        page.wait_for_timeout(3_000)

    def refresh(self) -> None:
        """Load the mailbox afresh (list rows keep their old text, e.g. "(No subject)", after a save until
        then). Goes to the mail URL rather than reloading: a reload lands on whichever message was open,
        where the Home ribbon with "New mail" is not guaranteed."""
        self.page.goto(self.url)
        self._new_mail_button().wait_for(state="visible", timeout=60_000)

    def drafted_subjects(self, subjects: Iterable[str]) -> set[str]:
        """Which of ``subjects`` are in the Drafts list, scanned through its whole length."""
        self._open_drafts()
        rows = self._all_row_texts()
        return {s for s in subjects if any(s in row for row in rows)}

    def verify(self, subjects: Iterable[str]) -> list[str]:
        """Subjects missing from Drafts as the server returns them after a fresh load."""
        wanted = list(subjects)
        self.refresh()
        present = self.drafted_subjects(wanted)
        return [s for s in wanted if s not in present]

    def restore_subjects(self, expected: dict[str, str]) -> int:
        """Subject-less drafts get their subject back from an attachment name in ``expected``
        (``{file name: subject}``); drafts that carry none of those files are left alone."""
        if not expected:
            return 0
        names = re.compile("|".join(re.escape(n) for n in expected))
        restored = 0
        for _ in range(len(expected) + 5):
            if not self._restore_next(names, expected):
                break
            restored += 1
        return restored

    def _restore_next(self, names: re.Pattern, expected: dict[str, str]) -> bool:
        """Repair the first subject-less draft carrying one of ``expected``'s files, scrolling the
        Drafts list until one is found; False when none is left. A repaired draft moves in the list,
        so the caller starts the next search from the top again."""
        page = self.page
        self.refresh()
        self._open_drafts()
        for _ in self._scroll_positions():
            rows = self._rows().filter(has_text=NO_SUBJECT)
            for index in range(rows.count()):
                rows.nth(index).click()
                page.wait_for_timeout(2_000)
                attachment = page.get_by_role("option", name=names)
                if attachment.count() == 0:
                    continue
                text = attachment.first.inner_text()
                subject_text = next(s for n, s in expected.items() if n in text)
                box = self._subject_box()
                try:
                    box.wait_for(state="visible", timeout=10_000)
                except Exception:
                    logger.warning("⚠️ no Subject box for the draft carrying %s; screenshot %s",
                                   subject_text, self.screenshot("restore-" + subject_text))
                    continue
                box.click()
                box.fill(subject_text)
                page.keyboard.press("Control+s")
                page.wait_for_timeout(3_000)
                return True
        return False

    # -- sending ----------------------------------------------------------------------------

    def _single_recipient_is(self, address: str) -> bool:
        """Exactly one To chip. External chips show ``Name <address>``; a same-tenant contact shows only
        its display name, so a directory-resolved chip is accepted for an address typed in full."""
        chips = self.page.locator('div[aria-label="To"] span[class*="_EType_RECIPIENT_ENTITY"]')
        if chips.count() != 1:
            return False
        chip = chips.first
        if address.lower() in ((chip.get_attribute("aria-label") or "") + chip.inner_text()).lower():
            return True
        return chip.locator('[class*="validPill"]').count() > 0

    def _count_drafts(self, subject: str) -> int:
        """Rows with ``subject`` in the open Drafts list, scanned through its whole length."""
        return sum(subject in row for row in self._all_row_texts())

    def _open_draft(self, subject: str) -> None:
        """Click the row with ``subject``, scrolling the list until it is rendered."""
        for _ in self._scroll_positions():
            rows = self._rows().filter(has_text=subject)
            if rows.count():
                rows.first.click()
                return
        raise RuntimeError(f"no draft row with subject {subject!r} to open")

    def send(self, spec: DraftSpec, timeout_s: int = 240) -> None:
        """Send the one draft matching ``spec``. Refuses unless exactly one draft has the subject,
        its recipient and attachments match, and afterwards it has left Drafts and is in Sent Items."""
        page = self.page
        self._open_drafts()
        found = self._count_drafts(spec.subject)
        if found != 1:
            raise RuntimeError(f"expected exactly one draft with subject {spec.subject!r}, found {found}")
        self._open_draft(spec.subject)
        page.wait_for_timeout(2_000)
        if not self._single_recipient_is(spec.to):
            raise RuntimeError(f"draft {spec.subject!r} is not addressed to exactly {spec.to}")
        if self._subject_box().input_value() != spec.subject:
            raise RuntimeError(f"draft {spec.subject!r} has a different subject than expected")
        for path in spec.attachments:
            page.get_by_role("option", name=re.compile(re.escape(path.name))).first.wait_for(
                state="visible", timeout=15_000)
        page.get_by_role("button", name=re.compile(r"^Send$")).first.click()
        waited = 0
        while waited < timeout_s:
            page.wait_for_timeout(3_000)
            waited += 3
            self._open_drafts()
            if self._count_drafts(spec.subject) == 0:
                break
        else:
            raise RuntimeError(f"{spec.subject!r} is still in Drafts after {timeout_s} s; not confirmed sent")
        self._open_folder("Sent Items")
        if self._rows().filter(has_text=spec.subject).count() == 0:
            raise RuntimeError(f"{spec.subject!r} left Drafts but is not in Sent Items; check the mailbox")
