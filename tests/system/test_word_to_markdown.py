"""Offline tests for system/word_to_markdown.py (temp dirs, generated .docx only)."""

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from docx import Document

from system import word_to_markdown as w2m


class ConversionFailureTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.folder = Path(self._tmp.name)

    def _corrupt_docx(self) -> Path:
        path = self.folder / "broken.docx"
        path.write_bytes(b"this is not a zip archive")
        return path

    def _valid_docx(self) -> Path:
        path = self.folder / "fine.docx"
        document = Document()
        document.add_paragraph("Hello from a synthetic document")
        document.save(path)
        return path

    def test_corrupt_document_is_reported_as_a_failure(self):
        # convert_document used to turn the exception into an error-markdown
        # string, which was written to the .md and reported as success.
        source = self._corrupt_docx()
        converter = w2m.WordToMarkdownConverter()

        self.assertFalse(converter.convert_single_document(str(source)))
        self.assertFalse(source.with_suffix(".md").exists())

    def test_command_line_exits_non_zero_for_a_corrupt_document(self):
        source = self._corrupt_docx()
        with mock.patch("sys.argv", ["word_to_markdown.py", str(source)]):
            with self.assertRaises(SystemExit) as raised:
                w2m.main()
        self.assertEqual(raised.exception.code, 1)

    def test_valid_document_still_converts(self):
        source = self._valid_docx()
        converter = w2m.WordToMarkdownConverter()

        self.assertTrue(converter.convert_single_document(str(source)))
        self.assertIn("Hello from a synthetic document", source.with_suffix(".md").read_text(encoding="utf-8"))

    def test_batch_conversion_does_not_count_a_corrupt_document_as_processed(self):
        source = self._corrupt_docx()
        converter = w2m.WordToMarkdownConverter()
        converter.test_folder = self.folder / "converted"
        converter.test_folder.mkdir()

        with mock.patch.object(converter, "_create_test_folder"), \
                mock.patch.object(converter, "_find_word_documents", return_value=[source]), \
                mock.patch.object(converter, "_copy_to_test_folder", return_value=source), \
                self.assertLogs(w2m.logger, level="INFO") as logs:
            converter.run_test_conversion()

        self.assertTrue(any("0/1 documents processed" in line for line in logs.output), logs.output)
        self.assertEqual(list(converter.test_folder.iterdir()), [])


if __name__ == "__main__":
    unittest.main()
