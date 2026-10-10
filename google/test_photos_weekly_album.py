"""Offline tests for google/photos_weekly_album.py, its Gmail compose step and browser_stealth.py
(no browser, no network)."""

import json
import re
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest import mock

import browser_stealth as bs
from google import photos_weekly_album as pwa
from google import photos_weekly_album_job as job
from google.gmail_web import GmailError, GmailWeb

SAMPLE = Path(pwa.__file__).with_name("photos_weekly_album.json.sample")


class WeekRangeTests(unittest.TestCase):
    def test_saturday_run_takes_the_week_that_just_ended(self):
        self.assertEqual(pwa.week_range(date(2026, 10, 10)), (date(2026, 10, 3), date(2026, 10, 9)))

    def test_friday_is_still_open_so_the_previous_week_is_taken(self):
        self.assertEqual(pwa.week_range(date(2026, 10, 9)), (date(2026, 9, 26), date(2026, 10, 2)))

    def test_every_weekday_gives_a_saturday_to_friday_week_before_today(self):
        for offset in range(14):
            today = date(2026, 10, 10) + timedelta(days=offset)
            with self.subTest(today=today):
                start, end = pwa.week_range(today)
                self.assertEqual((start.weekday(), end.weekday()), (5, 4))
                self.assertEqual(end - start, timedelta(days=6))
                self.assertTrue(1 <= (today - end).days <= 7)

    def test_crosses_month_and_year_boundaries(self):
        self.assertEqual(pwa.week_range(date(2027, 1, 2)), (date(2026, 12, 26), date(2027, 1, 1)))

    def test_days_of_lists_both_ends(self):
        days = pwa.days_of(date(2026, 10, 3), date(2026, 10, 9))
        self.assertEqual((len(days), days[0], days[-1]), (7, date(2026, 10, 3), date(2026, 10, 9)))


class TitleTests(unittest.TestCase):
    COUNTERS = {"a": {"value": 10, "week_start": "2026-09-26"}, "b": {"value": 3, "week_start": "2026-09-26"}}

    def test_counters_go_up_one_per_week(self):
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 9, 26)), 10)
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 10, 3)), 11)
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2027, 9, 25)), 62)

    def test_weeks_before_the_base_count_down(self):
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 9, 12)), 8)

    def test_base_week_must_start_on_saturday(self):
        with self.assertRaisesRegex(ValueError, "not a Saturday"):
            pwa.counter_value(1, date(2026, 9, 27), date(2026, 10, 3))

    def test_template_gets_every_counter(self):
        self.assertEqual(pwa.build_title("W{a}-{b}", self.COUNTERS, date(2026, 10, 10)), "W12-5")

    def test_unknown_placeholder_is_a_config_error(self):
        with self.assertRaisesRegex(ValueError, "placeholder 'c'"):
            pwa.build_title("W{a}-{c}", self.COUNTERS, date(2026, 10, 3))

    def test_committed_sample_builds_a_title(self):
        cfg = pwa.Config.load(SAMPLE)
        self.assertEqual(cfg.profile_dir, pwa.DEFAULT_PROFILE)
        self.assertIn("101", pwa.build_title(cfg.title_template, cfg.counters, date(2026, 1, 10)))


class LabelTests(unittest.TestCase):
    def test_photo_and_video_labels(self):
        self.assertEqual(pwa.label_datetime("Photo - Landscape - Oct 3, 2026, 8:23:55 AM"),
                         datetime(2026, 10, 3, 8, 23, 55))
        self.assertEqual(pwa.label_date("Video - Portrait - Oct 9, 2026, 6:47:52 PM"), date(2026, 10, 9))

    def test_narrow_no_break_space_before_am_pm(self):
        self.assertEqual(pwa.label_date("Photo - Portrait - Dec 31, 2026, 11:59:59\u202fPM"), date(2026, 12, 31))

    def test_unrecognised_labels_give_none(self):
        for label in ("", "Select", "Photo - Landscape - Oct 3, 8:23:55 AM", "Album - Oct 3, 2026, 8:23:55 AM"):
            with self.subTest(label=label):
                self.assertIsNone(pwa.label_date(label))


class ArgsTests(unittest.TestCase):
    def test_from_and_to_go_together_and_in_order(self):
        for argv in (["--from", "2026-10-03"], ["--from", "2026-10-09", "--to", "2026-10-03"],
                     ["--from", "2026-10-01", "--to", "2026-11-15"]):
            with self.subTest(argv=argv), self.assertRaises(SystemExit):
                pwa.parse_args(argv)

    def test_explicit_week(self):
        args = pwa.parse_args(["--from", "2026-10-03", "--to", "2026-10-09", "--dry-run"])
        self.assertEqual((args.start, args.end, args.dry_run, args.send),
                         (date(2026, 10, 3), date(2026, 10, 9), True, False))

    def test_send_is_opt_in(self):
        self.assertTrue(pwa.parse_args(["--send"]).send)
        with self.assertRaises(SystemExit):
            pwa.parse_args(["--send", "--login"])

    def test_send_without_recipients_is_a_config_error(self):
        sample = json.loads(SAMPLE.read_text(encoding="utf-8"))
        del sample["email"]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "c.json"
            path.write_text(json.dumps(sample), encoding="utf-8")
            with self.assertLogs("photos_weekly_album", "ERROR") as logs:
                self.assertEqual(pwa.main(["--send", "--config", str(path)]), 2)
        self.assertIn("email.to", logs.output[0])


EMAIL = {"to": ["a@example.com", "b@example.org"], "subject": "Week {week_start} to {week_end}",
         "body": "Hi,\n{title}\n{link}\n"}


class EmailConfigTests(unittest.TestCase):
    def test_no_recipients_means_no_sharing_or_draft(self):
        for raw in (None, {}, {"to": []}, {"subject": "x {title}", "body": "{link}"}):
            with self.subTest(raw=raw):
                self.assertIsNone(pwa.EmailConfig.load(raw))

    def test_renders_week_title_and_link(self):
        email = pwa.EmailConfig.load(EMAIL)
        subject, body = email.render("T", "https://photos.app.goo.gl/abc", date(2026, 10, 3), date(2026, 10, 9))
        self.assertEqual(subject, "Week 2026-10-03 to 2026-10-09")
        self.assertEqual(body, "Hi,\nT\nhttps://photos.app.goo.gl/abc\n")
        self.assertFalse(email.send)

    def test_rejects_configs_that_could_misfire(self):
        cases = {
            "plain addresses": {**EMAIL, "to": ["Someone <a@example.com>"]},
            "twice": {**EMAIL, "to": ["a@example.com", "A@example.com"]},
            "unknown placeholders": {**EMAIL, "body": "{link} {name}"},
            "own subject": {**EMAIL, "subject": "Weekly photos"},
            "link": {**EMAIL, "body": "Hi"},
            "malformed": {**EMAIL, "body": "{link"},
        }
        for message, raw in cases.items():
            with self.subTest(message), self.assertRaisesRegex(ValueError, message):
                pwa.EmailConfig.load(raw)

    def test_committed_sample_has_a_valid_email_block(self):
        email = pwa.Config.load(SAMPLE).email
        self.assertEqual(len(email.to), 2)
        self.assertTrue(all(a.endswith((".com", ".org")) and "example" in a for a in email.to))
        self.assertFalse(email.send)
        subject, body = email.render("T", "L", date(2026, 10, 3), date(2026, 10, 9))
        self.assertIn("2026-10-03", subject)
        self.assertIn("L", body)


class ShareLinkTests(unittest.TestCase):
    def test_short_and_full_links_are_found(self):
        self.assertEqual(pwa.find_share_link(["x", 'b ["https://photos.app.goo.gl/Ab12Cd",null]']),
                         "https://photos.app.goo.gl/Ab12Cd")
        self.assertEqual(pwa.find_share_link(["https://photos.google.com/share/AF1Q-x_y?key=k3Y-z_ and more"]),
                         "https://photos.google.com/share/AF1Q-x_y?key=k3Y-z_")

    def test_album_and_keyless_urls_are_not_share_links(self):
        self.assertIsNone(pwa.find_share_link(["https://photos.google.com/album/AF1Q",
                                               "https://photos.google.com/share/AF1Q", "", None]))


LINK = "https://photos.app.goo.gl/Ab12Cd"


class _FakeSharePage:
    """An album's Share dialogs as Photos draws them (Oct 2026). "Create link" in the Share dialog
    opens a "Create link to share" confirmation; only its own "Create link" makes the link, which it
    then shows in an input. Once a link exists the Share dialog offers "Copy link" and carries the
    link in its markup only. ``broken`` stands for a dialog that never shows the link."""

    def __init__(self, linked=False, broken=False):
        self.linked, self.broken = linked, broken
        self.url = "https://photos.google.com/album/A"
        self.dialogs, self.clicks = [], []
        self.keyboard = self

    def _share_dialog(self):
        if self.linked:
            return {"heading": "Invite to album", "buttons": ["Close", "Copy link"],
                    "texts": ["Invite to album\nCopy link", "" if self.broken else f'<div data-u="{LINK}">']}
        return {"heading": "Invite to album", "buttons": ["Close", "Create link"], "texts": ["Invite to album", ""]}

    def evaluate(self, script, *args):
        return [dict(d) for d in self.dialogs] if script is pwa.SHARE_DIALOGS_JS else ""

    def get_by_role(self, role, name):
        return self._target("Share")

    def locator(self, selector):
        return self._target(selector.split('"')[1])

    def _target(self, mark):
        page = self

        class Target:
            first = property(lambda self: self)

            def click(self):
                page.clicks.append(mark)
                page._clicked(mark)
        return Target()

    def _clicked(self, mark):
        if mark == "Share":
            self.dialogs = [self._share_dialog()]
        elif mark == "0:Create link":
            self.dialogs.append({"heading": "Create link to share", "buttons": ["Close", "Create link"],
                                 "texts": ["Create link to share", ""]})
        elif mark == "1:Create link":
            self.linked = True
            shown = "" if self.broken else LINK
            self.dialogs[1] = {"heading": "Create link to share", "buttons": ["Close", "Copy"],
                               "texts": [f"Create link to share\n{shown}", shown, ""]}

    def goto(self, url, wait_until):
        self.url, self.dialogs = url, []

    def wait_for_url(self, pattern, timeout):
        raise TimeoutError("stays at /album/")

    def bring_to_front(self):
        pass

    def wait_for_timeout(self, ms):
        pass

    def press(self, key):
        self.dialogs = []

    def on(self, event, handler):
        pass

    remove_listener = on


class ShareFlowTests(unittest.TestCase):
    def _share(self, page):
        web = pwa.PhotosWeb(Path("unused"))
        web.page, web._context = page, mock.Mock()
        return web.share_link(page.url)

    def test_create_link_is_confirmed_in_the_second_dialog_and_read_from_it(self):
        page = _FakeSharePage()
        with self.assertLogs("photos_weekly_album", "INFO") as logs:
            self.assertEqual(self._share(page), LINK)
        self.assertEqual(page.clicks, ["Share", "0:Create link", "1:Create link"])
        self.assertIn("share link created", logs.output[0])

    def test_an_existing_link_is_read_without_creating_another(self):
        page = _FakeSharePage(linked=True)
        self.assertEqual(self._share(page), LINK)
        self.assertEqual(page.clicks, ["Share"])

    def test_an_unreadable_link_says_what_was_tried_and_which_buttons_showed(self):
        page = _FakeSharePage(broken=True)
        with self.assertLogs("photos_weekly_album", "WARNING") as logs, \
                self.assertRaisesRegex(pwa.FlowError, "now shared by link, but the script could not read it"):
            self._share(page)
        self.assertIn("tried: dialogs, responses, page URL, reloaded page URL", logs.output[0])
        self.assertIn("'Create link to share': ['Close', 'Copy']", logs.output[0])


class _FakeAlbumList:
    """The /albums page: ``cards`` are (href, card text); evaluate applies the script's own
    ``a[href*="..."]`` selectors, so the test exercises the real selector."""

    def __init__(self, cards):
        self.cards, self.url = cards, ""

    def evaluate(self, script, *args):
        parts = re.findall(r'a\[href\*="([^"]+)"\]', script)
        return [{"href": href, "lines": text.split("\n")} for href, text in self.cards
                if any(part in href for part in parts)]

    def goto(self, url, wait_until):
        self.url = url

    def bring_to_front(self):
        pass

    def wait_for_timeout(self, ms):
        pass


class FindAlbumTests(unittest.TestCase):
    def _find(self, cards, title="Week 1"):
        web = pwa.PhotosWeb(Path("unused"))
        web.page = _FakeAlbumList(cards)
        web._wheel = lambda: False
        return web.find_album(title)

    def test_a_link_shared_album_is_found_by_its_share_card(self):
        share = "https://photos.google.com/share/AF1Q?key=k"
        self.assertEqual(self._find([("https://photos.google.com/album/X", "Other\n3 items"),
                                     (share, "Week 1\n183 items \u00a0\u00b7\u00a0 Shared")]), (share, 183))

    def test_a_private_album_is_still_found_and_a_missing_one_is_none(self):
        cards = [("https://photos.google.com/album/X", "Week 1\n12 items")]
        self.assertEqual(self._find(cards), ("https://photos.google.com/album/X", 12))
        self.assertIsNone(self._find(cards, title="Week 2"))


class _FakeMail:
    """Gmail by subject: ``drafts``/``sent`` hold subjects; sending moves the draft to Sent."""

    def __init__(self, drafts=(), sent=(), save_fails=False, body_link="L"):
        self.drafts, self.sent = list(drafts), list(sent)
        self.save_fails, self.body_link = save_fails, body_link
        self.created, self.sends = [], []

    def subjects(self, folder, subject, expect=False):
        return [s for s in (self.drafts if folder == "draft" else self.sent) if s == subject]

    def create_draft(self, to, subject, body):
        self.created.append((tuple(to), subject, body))
        if not self.save_fails:
            self.drafts.append(subject)

    def send_draft(self, subject, to, must_contain):
        if must_contain != self.body_link:
            raise RuntimeError("the draft's body doesn't carry this run's share link; not sent")
        self.sends.append(subject)
        self.drafts.remove(subject)
        self.sent.append(subject)


class DeliverTests(unittest.TestCase):
    TO = ["a@example.com"]

    def _deliver(self, mail, send=False, link="L"):
        return pwa.deliver(mail, self.TO, "S", "body L", link, send)

    def test_first_run_drafts_and_never_sends(self):
        mail = _FakeMail(drafts=["other"])
        self.assertEqual(self._deliver(mail), "drafted")
        self.assertEqual((mail.created, mail.sends), ([(("a@example.com",), "S", "body L")], []))

    def test_rerun_reuses_the_draft(self):
        mail = _FakeMail(drafts=["S"])
        self.assertEqual(self._deliver(mail), "draft-exists")
        self.assertEqual((mail.created, mail.drafts), ([], ["S"]))

    def test_already_sent_is_neither_drafted_nor_sent_again(self):
        mail = _FakeMail(sent=["S"])
        self.assertEqual(self._deliver(mail, send=True), "already-sent")
        self.assertEqual((mail.created, mail.sends), ([], []))

    def test_send_sends_only_this_weeks_draft(self):
        mail = _FakeMail(drafts=["older", "S", "other"])
        self.assertEqual(self._deliver(mail, send=True), "sent")
        self.assertEqual((mail.sends, mail.drafts), (["S"], ["older", "other"]))

    def test_send_after_a_fresh_draft(self):
        mail = _FakeMail()
        self.assertEqual(self._deliver(mail, send=True), "sent")
        self.assertEqual((len(mail.created), mail.sends), (1, ["S"]))

    def test_duplicate_drafts_abort_without_sending(self):
        mail = _FakeMail(drafts=["S", "S"])
        with self.assertRaisesRegex(pwa.FlowError, "2 drafts"):
            self._deliver(mail, send=True)
        self.assertEqual((mail.created, mail.sends), ([], []))

    def test_a_draft_that_did_not_save_is_reported(self):
        with self.assertRaisesRegex(pwa.FlowError, "isn't in Drafts"):
            self._deliver(_FakeMail(save_fails=True), send=True)

    def test_a_draft_without_this_runs_link_is_not_sent(self):
        mail = _FakeMail(drafts=["S"], body_link="old link")
        with self.assertRaisesRegex(RuntimeError, "not sent"):
            self._deliver(mail, send=True)
        self.assertEqual(mail.sent, [])


class _FakeCompose:
    """Gmail compose as seen live (#168). Typed text opens the suggestion list over Subject only
    once time passes (a wait, or a click's retries); a comma stays in the field as text. With the
    list open, Tab makes a chip, empties the field and closes the list; before it opens, Tab moves
    focus to Subject and the To field collapses. Escape with no list open closes the window."""

    SUGGESTING = 'input[aria-label="To recipients"][aria-expanded="true"]'

    def __init__(self):
        self.text, self.chips, self.suggesting = "", [], False
        self.to_open, self.open, self.subject, self.saved = True, True, "", False

    # page
    keyboard = last = property(lambda self: self)

    def goto(self, url, wait_until=None):
        pass

    def bring_to_front(self):
        pass

    def wait_for_timeout(self, ms):
        self.settle()

    def insert_text(self, text):
        pass

    def press(self, key):  # page.keyboard
        pass

    def get_by_role(self, role, name=None):
        return _FakeComposeElement(self, (role, name if isinstance(name, str) else "regex"))

    def locator(self, selector, has=None):
        return _FakeComposeElement(self, selector)

    # dialog
    def wait_for(self, state=None, timeout=None):
        if state == "detached" and not self.saved:
            raise TimeoutError("compose window still open")

    def evaluate(self, js):
        return list(self.chips)

    def settle(self):
        self.suggesting = bool(self.text)

    def blur_to(self):  # leaving the field turns its text into one chip
        if self.text:
            self.chips.append(self.text)
        self.text, self.suggesting, self.to_open = "", False, False

    def type_to(self, text):
        if not self.to_open:
            raise TimeoutError("the To field is collapsed")
        self.text += text

    def key_to(self, key):
        if key == "Tab" and self.suggesting:
            self.chips.append(self.text)
            self.text, self.suggesting = "", False
        elif key == "Tab":
            self.blur_to()
        elif key == "Escape" and self.suggesting:
            self.suggesting = False
        elif key == "Escape":
            self.open = False

    def click_subject(self):
        self.settle()
        if not self.open:
            raise TimeoutError("element was detached from the DOM")
        if self.suggesting:
            raise TimeoutError('<div role="option" peoplekit-id=...> intercepts pointer events')
        self.blur_to()


class _FakeComposeElement:
    def __init__(self, compose, key):
        self.compose, self.key = compose, key

    first = last = property(lambda self: self)

    def filter(self, has=None):
        return self.compose  # the compose dialog

    def click(self):
        if self.key == 'input[name="subjectbox"]':
            self.compose.click_subject()
        elif self.key in (("button", "Save & close"), '[aria-label^="Discard draft"]'):
            self.compose.saved = True

    def press_sequentially(self, text, delay=None):
        if self.key == ("combobox", "To recipients"):
            self.compose.type_to(text)
        else:
            self.compose.subject += text

    def press(self, key):
        self.compose.key_to(key)

    def wait_for(self, state=None, timeout=None):
        if self.key != _FakeCompose.SUGGESTING:
            return
        self.compose.settle()
        if self.compose.suggesting != (state == "visible"):
            raise TimeoutError(f"suggestion list not {state}")


class CreateDraftTests(unittest.TestCase):
    TO = ["c@example.com", "a@example.com", "b@example.com"]

    def test_each_address_becomes_a_chip_and_subject_is_reachable(self):
        compose = _FakeCompose()
        GmailWeb(compose).create_draft(self.TO, "S", "body")
        self.assertEqual((compose.chips, compose.subject, compose.saved),
                         (sorted(self.TO), "S", True))

    def test_recipients_that_differ_discard_the_draft(self):
        compose = _FakeCompose()
        compose.chips = ["stray@example.com"]
        with self.assertRaisesRegex(GmailError, "4 recipients, not the 3"):
            GmailWeb(compose).create_draft(self.TO, "S", "body")
        self.assertEqual(compose.subject, "")


class _FakeCheckbox:
    """A checkbox that, like a day header's, has no size until the pointer is over its row."""

    def __init__(self, shown):
        self.shown, self.checked, self.events = shown, False, []

    def is_visible(self):
        return self.shown

    def hover(self):
        self.events.append("hover")
        self.shown = True

    def click(self):
        if not self.shown:
            raise TimeoutError("element is not visible")
        self.events.append("click")
        self.checked = not self.checked

    def wait_for(self, timeout):
        if not self.checked:
            raise TimeoutError("still unchecked")


class _FakePage:
    def __init__(self, box):
        self.box = box

    def locator(self, selector):
        return self

    @property
    def first(self):
        return self.box


class ClickCheckedTests(unittest.TestCase):
    def _click(self, box):
        web = pwa.PhotosWeb(Path("unused"))
        web.page = _FakePage(box)
        web._click_checked("[data-pwa-head=k]", "[data-pwa-head=k] >> xpath=..", True, "day")

    def test_hidden_checkbox_is_revealed_by_hovering_before_the_click(self):
        box = _FakeCheckbox(shown=False)
        self._click(box)
        self.assertEqual(box.events, ["hover", "click"])

    def test_shown_checkbox_is_clicked_without_hovering_what_it_covers(self):
        box = _FakeCheckbox(shown=True)
        self._click(box)
        self.assertEqual(box.events, ["click"])


class _FakeContext:
    def __init__(self):
        self.scripts = []

    def add_init_script(self, script):
        self.scripts.append(script)


class _FakePlaywright:
    def __init__(self, failures):
        self.failures = list(failures)
        self.calls = []
        self.chromium = self

    def launch_persistent_context(self, **kwargs):
        self.calls.append(kwargs)
        if self.failures:
            raise self.failures.pop(0)
        return _FakeContext()


BUSY = RuntimeError("Target page, context or browser has been closed ... exitCode=21, signal=null")


class LaunchPersistentTests(unittest.TestCase):
    def test_busy_profile_is_waited_for_then_launched_with_stealth(self):
        pw, waits = _FakePlaywright([BUSY, BUSY]), []
        context = bs.launch_persistent(pw, "prof", headless=False, disabled_features=("X",),
                                       delays=(1, 2, 3), sleep=waits.append)
        self.assertEqual(waits, [1, 2])
        self.assertEqual(context.scripts, [bs.STEALTH_INIT_SCRIPT])
        kwargs = pw.calls[-1]
        self.assertEqual((kwargs["channel"], kwargs["chromium_sandbox"]), ("chrome", True))
        self.assertIn("--disable-features=Translate,X", kwargs["args"])

    def test_gives_up_after_the_schedule_without_killing_anything(self):
        pw, waits = _FakePlaywright([BUSY] * 3), []
        with self.assertRaises(bs.ProfileBusy):
            bs.launch_persistent(pw, "prof", headless=True, delays=(1, 2), sleep=waits.append)
        self.assertEqual(waits, [1, 2])

    def test_other_launch_errors_are_not_retried(self):
        pw, waits = _FakePlaywright([RuntimeError("Executable doesn't exist")]), []
        with self.assertRaisesRegex(RuntimeError, "Executable"):
            bs.launch_persistent(pw, "prof", headless=True, delays=(1,), sleep=waits.append)
        self.assertEqual(waits, [])


class _PlaywrightTimeout(Exception):
    """Stands in for playwright's TimeoutError, which is recognised by its name and module."""

_PlaywrightTimeout.__name__ = "TimeoutError"
_PlaywrightTimeout.__module__ = "playwright._impl._errors"


class FailureKindTests(unittest.TestCase):
    def test_hidden_window_and_playwright_timeouts_are_told_apart(self):
        self.assertEqual(pwa.failure_kind(pwa.WindowHidden("minimised")), "window-hidden")
        self.assertEqual(pwa.failure_kind(_PlaywrightTimeout("waiting for locator")), "page-timeout")

    def test_everything_else_is_a_plain_error(self):
        for exc in (pwa.FlowError("x"), RuntimeError("y"), TimeoutError("builtin, not playwright")):
            with self.subTest(exc=exc):
                self.assertEqual(pwa.failure_kind(exc), "error")


class _FakeWeb:
    """A PhotosWeb whose steps are scripted: ``steps`` maps a method name to a value or an exception."""

    steps: dict = {}

    def __init__(self, profile_dir):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        pass

    def _step(self, name):
        value = self.steps.get(name)
        if isinstance(value, Exception):
            raise value
        return value

    def require_signed_in(self):
        self._step("require_signed_in")

    def find_album(self, title):
        return self._step("find_album")


def _config(tmp: Path, **email) -> Path:
    raw = json.loads(SAMPLE.read_text(encoding="utf-8"))
    raw.update(profile_dir=str(tmp / "profile"), log_file=str(tmp / "run.log"))
    if email:
        raw["email"].update(email)
    else:
        del raw["email"]
    path = tmp / "config.json"
    path.write_text(json.dumps(raw), encoding="utf-8")
    return path


class RunRecordTests(unittest.TestCase):
    WEEK = (date(2026, 10, 3), date(2026, 10, 9))

    def _run(self, steps, **email):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = pwa.Config.load(_config(Path(tmp), **email))
            record: dict = {}
            _FakeWeb.steps = steps
            with mock.patch.object(pwa, "PhotosWeb", _FakeWeb), self.assertLogs("photos_weekly_album"):
                code = pwa.run(cfg, *self.WEEK, "T", False, record=record)
            logged = json.loads(cfg.log_file.read_text(encoding="utf-8").splitlines()[-1])
        return code, record, logged

    def test_album_already_there_is_exit_0_and_records_the_url(self):
        code, record, logged = self._run({"find_album": ("https://photos.google.com/album/X", 12)})
        self.assertEqual((code, record["outcome"], record["url"]), (0, "exists", "https://photos.google.com/album/X"))
        self.assertNotIn("failure", record)
        self.assertEqual(logged["outcome"], "exists")

    def test_each_failure_has_its_own_exit_code_and_kind(self):
        cases = {
            "not-signed-in": (pwa.NotSignedIn("signed out"), 3),
            "window-hidden": (pwa.WindowHidden("minimised"), 4),
            "page-timeout": (_PlaywrightTimeout("locator timed out"), 5),
            "error": (RuntimeError("boom"), 1),
        }
        for kind, (exc, expected) in cases.items():
            with self.subTest(kind):
                code, record, logged = self._run({"find_album": exc})
                self.assertEqual((code, record["failure"], logged["failure"]), (expected, kind, kind))


class _Notifier:
    sent: list = []

    def __init__(self, env):
        self.env = env

    def send(self, text):
        self.sent.append((self.env["NOTIFY_CATEGORY"], text))
        return True


ENV = {"NOTIFY_PYTHON": "py", "NOTIFY_SCRIPT": "notify.py"}
RECORD = {"title": "Week 101", "week_start": date(2026, 10, 3), "week_end": date(2026, 10, 9),
          "url": "https://photos.google.com/album/X", "share_url": "https://photos.app.goo.gl/abc"}


class DescribeTests(unittest.TestCase):
    def test_created_album_reports_link_and_draft_status(self):
        for mail, draft in ((None, "not shared"), ("drafted", "saved, not sent"), ("sent", "Draft: sent")):
            with self.subTest(mail=mail):
                out = job.describe(0, {**RECORD, "outcome": "created", "mail": mail})
                self.assertEqual((out.kind, out.exit_code), ("done", 0))
                self.assertIn("https://photos.app.goo.gl/abc", out.text)
                self.assertIn(draft, out.text)

    def test_album_already_there_is_a_no_op_unless_it_made_progress(self):
        for mail in (None, "draft-exists", "already-sent"):
            self.assertEqual(job.describe(0, {**RECORD, "outcome": "exists", "mail": mail}).kind, "no-op")
        self.assertEqual(job.describe(0, {**RECORD, "outcome": "exists", "mail": "drafted"}).kind, "done")

    def test_empty_week_is_its_own_outcome(self):
        out = job.describe(0, {**RECORD, "outcome": "empty"})
        self.assertEqual((out.kind, out.exit_code), ("empty", 0))
        self.assertIn("nothing to add", out.text)

    def test_every_failure_reads_differently_and_says_what_to_do(self):
        texts = {}
        for kind in ("not-signed-in", "desktop-locked", "window-hidden", "page-timeout"):
            out = job.describe(job.EXIT[kind], {"title": "T", "failure": kind})
            self.assertTrue(out.failed)
            self.assertEqual(out.kind, kind)
            texts[kind] = out.text
        self.assertEqual(len(set(texts.values())), 4)
        self.assertIn("--login", texts["not-signed-in"])
        self.assertIn("Unlock", texts["desktop-locked"])

    def test_count_mismatch_and_plain_errors(self):
        self.assertEqual(job.describe(1, {"title": "T", "outcome": "count-mismatch", "url": "U"}).kind, "count-mismatch")
        out = job.describe(1, {"title": "T", "outcome": "error", "failure": "error", "error": "x" * 500})
        self.assertEqual((out.kind, out.exit_code), ("error", 1))
        self.assertLess(len(out.text), 300)

    def test_a_broken_share_step_still_points_at_the_album(self):
        out = job.describe(1, {**RECORD, "outcome": "created", "mail": "error", "failure": "error", "error": "boom"})
        self.assertIn("album is made", out.text)
        self.assertIn(RECORD["url"], out.text)


class ExecuteTests(unittest.TestCase):
    TODAY = date(2026, 10, 10)

    def _execute(self, locked=False, run_code=0, run_record=None, **email):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append({"dates": (start, end), "dry_run": dry_run, "send": send})
            record.update(run_record or {"outcome": "created", "url": "U", "mail": None})
            return run_code

        _Notifier.sent = []
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(job, "FleetNotifier", _Notifier), \
                self.assertLogs("photos_weekly_album_job"):
            path = _config(Path(tmp), **email)
            code = job.execute(path, self.TODAY, ENV, locked=lambda: locked, run=fake_run)
            log = Path(tmp, "run.log")
            lines = log.read_text(encoding="utf-8").splitlines() if log.exists() else []
        return code, calls, list(_Notifier.sent), lines

    def test_saturday_run_takes_the_week_that_just_ended_and_never_asks_to_send(self):
        code, calls, sent, _ = self._execute()
        self.assertEqual(code, 0)
        self.assertEqual(calls, [{"dates": (date(2026, 10, 3), date(2026, 10, 9)), "dry_run": False, "send": False}])
        self.assertEqual([category for category, _ in sent], ["log"])

    def test_sending_follows_the_config_alone(self):
        _, calls, _, _ = self._execute(send=False, to=["a@example.com"])
        self.assertFalse(calls[0]["send"])
        _, calls, _, _ = self._execute(send=True, to=["a@example.com"])
        self.assertTrue(calls[0]["send"])

    def test_locked_desktop_starts_nothing_and_alerts(self):
        code, calls, sent, lines = self._execute(locked=True)
        self.assertEqual((code, calls), (4, []))
        self.assertEqual(sent[0][0], "attention")
        self.assertIn("locked", sent[0][1])
        self.assertEqual(json.loads(lines[-1])["failure"], "desktop-locked")

    def test_a_failed_run_goes_to_the_attention_chat_with_its_exit_code(self):
        code, _, sent, _ = self._execute(run_code=4, run_record={"outcome": "error", "failure": "window-hidden"})
        self.assertEqual((code, sent[0][0]), (4, "attention"))
        self.assertIn("minimised", sent[0][1])

    def test_a_bad_config_is_notified_not_silent(self):
        _Notifier.sent = []
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(job, "FleetNotifier", _Notifier), \
                self.assertLogs("photos_weekly_album_job", "ERROR"):
            code = job.execute(Path(tmp) / "missing.json", self.TODAY, ENV)
        self.assertEqual((code, _Notifier.sent[0][0]), (2, "attention"))

    def test_the_desktop_check_runs_without_error(self):
        self.assertIn(job.desktop_locked(), (True, False, None))

    def test_the_notifier_gets_the_category_for_the_outcome(self):
        with mock.patch.object(job, "FleetNotifier", _Notifier):
            _Notifier.sent = []
            job.notify(ENV, job.Outcome("done", "ok", 0))
            job.notify(ENV, job.Outcome("error", "bad", 1))
        self.assertEqual([c for c, _ in _Notifier.sent], ["log", "attention"])


if __name__ == "__main__":
    unittest.main()
