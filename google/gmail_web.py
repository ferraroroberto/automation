"""Gmail on the web through an already-open Playwright page: find, draft and send by exact subject.

The caller owns the browser (real Chrome on a signed-in persistent profile, via the repo-root
``browser_stealth.py``). Selectors are matched on the English UI; see
``google/photos_weekly_album.md`` for the flow and its failure modes.
"""

from __future__ import annotations

import logging
import re
from typing import Iterable
from urllib.parse import quote

logger = logging.getLogger("gmail_web")

GMAIL_URL = "https://mail.google.com/mail/u/0/"
NO_MATCH = "No messages matched your search"
EMAIL_RE = re.compile(r"[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+")

# Recipient chips in a compose window carry the address in one of these attributes.
RECIPIENTS_JS = r"""(dialog) => [...dialog.querySelectorAll('[email], [data-hovercard-id]')]
  .map(e => e.getAttribute('email') || e.getAttribute('data-hovercard-id') || '')
  .filter(a => a.includes('@'))"""


class GmailError(RuntimeError):
    """Gmail did not behave as the flow expects."""


class GmailNotSignedIn(GmailError):
    pass


def norm(text: str) -> str:
    return " ".join(text.split())


class GmailWeb:
    """Drafts and sends on one Playwright page that is signed in to Google."""

    def __init__(self, page) -> None:
        self.page = page

    def _go(self, fragment: str = "#inbox") -> None:
        self.page.goto(GMAIL_URL + fragment, wait_until="domcontentloaded")
        self.page.bring_to_front()
        self.page.wait_for_timeout(3_000)

    def require_signed_in(self) -> None:
        self._go()
        compose = self.page.get_by_role("button", name=re.compile(r"^Compose$", re.I))
        try:
            compose.first.wait_for(state="visible", timeout=30_000)
        except Exception as exc:
            raise GmailNotSignedIn("the dedicated Chrome profile is not signed in to Gmail") from exc

    def subjects(self, folder: str, subject: str, expect: bool = False) -> list[str]:
        """Subjects in ``folder`` (``draft`` or ``sent``) equal to ``subject`` (Gmail's search is fuzzy).

        ``expect``: the message was just saved or sent, so search again a few times before
        answering "none" (the search can lag a moment behind).
        """
        query = f'in:{folder} subject:"{subject}"'
        for attempt in range(3 if expect else 1):
            if attempt:
                self.page.wait_for_timeout(3_000)
            self._go("#search/" + quote(query, safe=""))
            rows = self.page.locator("[role=main] tr.zA:visible")
            empty = self.page.locator("[role=main]:visible", has_text=NO_MATCH)
            for _ in range(20):
                if rows.count() or empty.count():
                    break
                self.page.wait_for_timeout(500)
            else:
                raise GmailError(f"the {folder} search neither listed messages nor said none matched")
            found = [norm(t) for t in self.page.locator("[role=main] tr.zA:visible span.bog").all_inner_texts()]
            matches = [s for s in found if s == norm(subject)]
            if matches:
                return matches
        return []

    def _row(self, subject: str):
        exact = re.compile(rf"^\s*{re.escape(norm(subject))}\s*$")
        return self.page.locator("[role=main] tr.zA:visible").filter(
            has=self.page.locator("span.bog", has_text=exact)).first

    def _compose_dialog(self):
        dialog = self.page.locator("[role=dialog]").filter(has=self.page.locator('input[name="subjectbox"]')).last
        dialog.wait_for(state="visible", timeout=20_000)
        return dialog

    def _recipients(self, dialog) -> set[str]:
        return {a.lower() for a in dialog.evaluate(RECIPIENTS_JS)}

    def create_draft(self, to: Iterable[str], subject: str, body: str) -> None:
        """Compose a message, check its recipients, then Save & close. Never sends."""
        wanted = {a.lower() for a in to}
        self._go()
        self.page.get_by_role("button", name=re.compile(r"^Compose$", re.I)).first.click()
        dialog = self._compose_dialog()
        field = dialog.get_by_role("combobox", name="To recipients")
        field.click()
        for address in sorted(wanted):
            # A trailing comma turns the typed text into a chip; Enter could pick an autocomplete row.
            field.press_sequentially(address + ",", delay=25)
        dialog.locator('input[name="subjectbox"]').click()
        got = self._recipients(dialog)
        if got != wanted:
            dialog.locator('[aria-label^="Discard draft"]').first.click()
            raise GmailError(f"compose shows {len(got)} recipients, not the {len(wanted)} configured; "
                             "draft discarded")
        dialog.locator('input[name="subjectbox"]').press_sequentially(subject, delay=25)
        text = dialog.get_by_role("textbox", name="Message Body")
        text.click()
        self.page.keyboard.press("Control+Home")  # above any signature
        for number, line in enumerate(body.split("\n")):
            if number:
                self.page.keyboard.press("Enter")
            if line:
                self.page.keyboard.insert_text(line)
        self.page.wait_for_timeout(1_500)
        dialog.get_by_role("button", name="Save & close").click()
        dialog.wait_for(state="detached", timeout=20_000)
        self.page.wait_for_timeout(2_000)

    def send_draft(self, subject: str, to: Iterable[str], must_contain: str) -> None:
        """Open the one draft titled ``subject``; send it only if it carries ``must_contain`` and ``to``."""
        if len(self.subjects("draft", subject)) != 1:
            raise GmailError("expected exactly one draft with this run's subject")
        self._row(subject).click()
        dialog = self._compose_dialog()
        body = dialog.get_by_role("textbox", name="Message Body").inner_text()
        if must_contain not in body:
            raise GmailError("the draft's body doesn't carry this run's share link; not sent")
        if self._recipients(dialog) != {a.lower() for a in to}:
            raise GmailError("the draft's recipients differ from the configured list; not sent")
        dialog.get_by_role("button", name=re.compile(r"^Send\b")).first.click()
        dialog.wait_for(state="detached", timeout=30_000)
        self.page.wait_for_timeout(3_000)
