"""The shared normalization pass used by normalize_names.py and normalize_url.py."""

import sys
import unittest
from pathlib import Path
from unittest import mock

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "notion"))

import utils as notion_utils  # noqa: E402


def _extract(page):
    return (page["id"], "t", page["value"]) if page.get("value") else None


def _run(pages, update_page, dry_run=False, annotate=None):
    return notion_utils.normalize_pages(
        pages,
        extract_page_info=_extract,
        transform=str.lower,
        update_page=update_page,
        make_result=lambda page_id, _t, original, new: {"id": page_id, "from": original, "to": new},
        dry_run=dry_run,
        verb="change",
        unchanged_label="Already normalized",
        annotate=annotate,
    )


class NormalizePagesTests(unittest.TestCase):
    PAGES = [{"id": "a", "value": "Foo"}, {"id": "b", "value": "bar"}, {"id": "c"}]

    def test_updates_only_changed_pages_and_skips_unextractable_ones(self):
        update = mock.Mock(return_value=True)
        results = _run(self.PAGES, update)
        update.assert_called_once_with("a", "foo")
        self.assertEqual([r["id"] for r in results], ["a", "b"])

    def test_dry_run_writes_nothing(self):
        update = mock.Mock(return_value=True)
        results = _run(self.PAGES, update, dry_run=True)
        update.assert_not_called()
        self.assertEqual(results[0]["to"], "foo")

    def test_annotation_is_appended_to_the_log_line(self):
        with self.assertLogs(notion_utils.log, level="INFO") as logs:
            _run(self.PAGES, mock.Mock(return_value=True), annotate=lambda value: f" [{value}]")
        self.assertTrue(any('"foo" [foo]' in line for line in logs.output))


class PatchPagePropertiesTests(unittest.TestCase):
    def test_error_response_status_is_logged_even_though_it_is_falsy(self):
        # requests.Response.__bool__ is .ok, so a 4xx response is falsy.
        response = requests.Response()
        response.status_code = 400
        response._content = b'{"message": "bad"}'
        error = requests.exceptions.HTTPError("400", response=response)
        self.assertFalse(response)
        with mock.patch.object(notion_utils.requests, "patch", side_effect=error):
            with self.assertLogs(notion_utils.log, level="ERROR") as logs:
                self.assertFalse(notion_utils.patch_page_properties({}, "p" * 32, {}))
        self.assertTrue(any("Status: 400" in line for line in logs.output))


if __name__ == "__main__":
    unittest.main()
