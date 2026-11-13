"""A direct-URL resume only appends when the server honours the Range request."""

import contextlib
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "video"))

import video_download  # noqa: E402


class _FakeResponse:
    def __init__(self, status_code, body):
        self.status_code = status_code
        self._body = body
        self.headers = {"content-length": str(len(body))}

    def raise_for_status(self):
        pass

    def iter_content(self, block_size):
        yield self._body


def _download(response, existing: bytes):
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / "clip.mp4"
        target.write_bytes(existing)
        with mock.patch.object(
            video_download, "pooled_request", return_value=contextlib.nullcontext(response)
        ):
            video_download.download_direct_url("https://host/clip.mp4", tmp, "clip.mp4")
        return target.read_bytes()


class ResumeTests(unittest.TestCase):
    def test_206_appends_the_remaining_bytes(self):
        self.assertEqual(_download(_FakeResponse(206, b"world"), b"hello "), b"hello world")

    def test_200_after_ignored_range_replaces_the_partial_file(self):
        self.assertEqual(_download(_FakeResponse(200, b"hello world"), b"hello "), b"hello world")

    def test_fresh_download_writes_the_body(self):
        self.assertEqual(_download(_FakeResponse(200, b"hello"), b""), b"hello")


if __name__ == "__main__":
    unittest.main()
