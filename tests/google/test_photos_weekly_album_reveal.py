"""Offline tests for PhotosWeb._click_checked / _reveal_box in google/photos_weekly_album.py
(no browser, no network)."""

import unittest
from pathlib import Path

from google import photos_weekly_album as pwa


class _FakeCheckbox:
    """A checkbox that, like a day header's, has no size until the pointer is over its row."""

    def __init__(self, shown):
        self.shown, self.checked, self.events = shown, False, []

    def is_visible(self):
        return self.shown

    def scroll_into_view_if_needed(self, timeout=None):
        pass

    def hover(self, timeout=None):
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

    def evaluate(self, script, *args):
        pass

    def wait_for_timeout(self, ms):
        pass


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


class _FakeTilePage:
    """One screenshot tile in the virtualised grid. Its checkbox only shows once the pointer is
    over the tile, and the tile's anchor exists only while the tile is scrolled into view, so a
    hover without a prior scroll waits for an anchor that is not there."""

    def __init__(self, in_view=False, stale_until_rescan=False, gone=False):
        self.in_view, self.stale, self.gone = in_view, stale_until_rescan, gone
        self.hovered = self.clicked = False
        self.rescans = 0
        self.timeouts = []

    def locator(self, selector):
        return _FakeTileLocator(self, "anchor" if "/photo/" in selector else
                                "state" if "aria-checked" in selector else "box")

    def evaluate(self, script, *args):
        if script == pwa.SCAN_JS:
            self.rescans += 1
            self.stale = False

    def wait_for_timeout(self, ms):
        pass


class _FakeTileLocator:
    first = property(lambda self: self)

    def __init__(self, page, kind):
        self.page, self.kind = page, kind

    def scroll_into_view_if_needed(self, timeout=None):
        self.page.timeouts.append(timeout)
        if self.page.gone or self.page.stale:
            raise TimeoutError("element is not attached to the DOM")
        self.page.in_view = True

    def is_visible(self):
        return self.page.hovered

    def hover(self, timeout=30_000):
        self.page.timeouts.append(timeout)
        if not self.page.in_view:
            raise TimeoutError("waiting for locator(a[href*=/photo/...])")
        self.page.hovered = True

    def click(self):
        self.page.clicked = self.page.hovered

    def wait_for(self, timeout=None):
        if not self.page.clicked:
            raise TimeoutError("aria-checked never changed")


class OffScreenTileTests(unittest.TestCase):
    def _click(self, page):
        web = pwa.PhotosWeb(Path("unused"))
        web.page = page
        web._click_checked('[data-pwa-tile="T"]', 'a[href*="/photo/T"]', False, "screenshot tile")

    def test_an_off_screen_tile_is_scrolled_into_view_before_its_checkbox_is_revealed(self):
        page = _FakeTilePage(in_view=False)
        self._click(page)
        self.assertTrue(page.clicked)

    def test_a_tile_re_rendered_by_the_scroll_is_re_marked_and_retried(self):
        page = _FakeTilePage(stale_until_rescan=True)
        with self.assertLogs("photos_weekly_album", "INFO") as logs:
            self._click(page)
        self.assertEqual(page.rescans, 1)
        self.assertTrue(page.clicked)
        self.assertIn("rescanning", logs.output[0])

    def test_a_tile_that_never_comes_back_fails_fast_naming_the_step(self):
        page = _FakeTilePage(gone=True)
        with self.assertRaises(pwa.FlowError) as ctx:
            self._click(page)
        self.assertIn("screenshot tile", str(ctx.exception))
        self.assertEqual(page.rescans, 1)
        self.assertTrue(all(t == pwa.REVEAL_TIMEOUT_MS for t in page.timeouts))
