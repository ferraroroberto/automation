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
from typing import Iterable, Optional

logger = logging.getLogger("outlook_web")

OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/"
NO_SUBJECT = "(No subject)"
STEALTH_INIT_SCRIPT = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"


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

    def _open_drafts(self) -> None:
        self._open_folder("Drafts")

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

    def drafted_subjects(self, subjects: Iterable[str]) -> set[str]:
        """Which of ``subjects`` are in the Drafts list as currently shown."""
        self._open_drafts()
        rows = self._rows().all_inner_texts()
        return {s for s in subjects if any(s in row for row in rows)}

    def verify(self, subjects: Iterable[str]) -> list[str]:
        """Subjects missing from Drafts as the server returns them after a fresh load."""
        wanted = list(subjects)
        self.page.reload()
        self._new_mail_button().wait_for(state="visible", timeout=60_000)
        present = self.drafted_subjects(wanted)
        return [s for s in wanted if s not in present]

    def restore_subjects(self, expected: dict[str, str]) -> int:
        """Subject-less drafts get their subject back from an attachment name in ``expected``
        (``{file name: subject}``); drafts that carry none of those files are left alone."""
        if not expected:
            return 0
        page = self.page
        names = re.compile("|".join(re.escape(n) for n in expected))
        restored = 0
        for _ in range(len(expected) + 5):
            self._open_drafts()
            rows = self._rows().filter(has_text=NO_SUBJECT)
            fixed = False
            for index in range(rows.count()):
                rows.nth(index).click()
                page.wait_for_timeout(2_000)
                attachment = page.get_by_role("option", name=names)
                if attachment.count() == 0:
                    continue
                text = attachment.first.inner_text()
                subject_text = next(s for n, s in expected.items() if n in text)
                box = self._subject_box()
                box.click()
                box.fill(subject_text)
                page.keyboard.press("Control+s")
                page.wait_for_timeout(3_000)
                restored += 1
                fixed = True
                break
            if not fixed:
                break
        return restored

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

    def send(self, spec: DraftSpec, timeout_s: int = 240) -> None:
        """Send the one draft matching ``spec``. Refuses unless exactly one draft has the subject,
        its recipient and attachments match, and afterwards it has left Drafts and is in Sent Items."""
        page = self.page
        self._open_drafts()
        rows = self._rows().filter(has_text=spec.subject)
        if rows.count() != 1:
            raise RuntimeError(f"expected exactly one draft with subject {spec.subject!r}, found {rows.count()}")
        rows.first.click()
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
            if self._rows().filter(has_text=spec.subject).count() == 0:
                break
        else:
            raise RuntimeError(f"{spec.subject!r} is still in Drafts after {timeout_s} s; not confirmed sent")
        self._open_folder("Sent Items")
        if self._rows().filter(has_text=spec.subject).count() == 0:
            raise RuntimeError(f"{spec.subject!r} left Drafts but is not in Sent Items; check the mailbox")
