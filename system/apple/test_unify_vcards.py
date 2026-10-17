"""Regression tests for unify_vcards field preservation (synthetic contacts only)."""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from unify_vcards import ContactMerger, VCardParser, VCardWriter  # noqa: E402

ALEX = "\n".join([
    "BEGIN:VCARD",
    "VERSION:3.0",
    "FN:Alex Example",
    "N:Example;Alex;;;",
    "TEL;TYPE=CELL:+1 555 0100",
    "BDAY:1990-04-02",
    "URL:https://example.test/alex",
    "CATEGORIES:friends,work",
    "X-CUSTOM-FIELD:alpha",
    "PHOTO;ENCODING=b;TYPE=JPEG:AAAA",
    " BBBB",
    " CCCC",
    "END:VCARD",
])

ALEX_TWIN = "\n".join([
    "BEGIN:VCARD",
    "VERSION:3.0",
    "FN:Alex Example",
    "N:Example;Alex;;;",
    "TEL;TYPE=HOME:+1 555 0101",
    "BDAY:1990-04-02",
    "NICKNAME:Lex",
    "X-CUSTOM-FIELD:beta",
    "END:VCARD",
])

PRESERVED = [
    "BDAY:1990-04-02",
    "URL:https://example.test/alex",
    "CATEGORIES:friends,work",
    "X-CUSTOM-FIELD:alpha",
    "PHOTO;ENCODING=b;TYPE=JPEG:AAAA",
    " BBBB",
    " CCCC",
]


def _written_lines(contacts) -> list:
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "out.vcf")
        VCardWriter().write_contacts(contacts, path)
        with open(path, "r", encoding="utf-8", newline="") as handle:
            return handle.read().replace("\r\n", "\n").split("\n")


class UntouchedContactTests(unittest.TestCase):
    def test_untouched_contact_is_written_unchanged(self):
        contacts = VCardParser().parse_content(ALEX + "\n")
        written = _written_lines(contacts)
        for line in PRESERVED:
            self.assertIn(line, written)
        # Folded PHOTO lines keep their order.
        start = written.index(PRESERVED[4])
        self.assertEqual(written[start:start + 3], PRESERVED[4:])


class MergedContactTests(unittest.TestCase):
    def test_merge_carries_unrecognised_lines_from_every_contact(self):
        a, b = VCardParser().parse_content(ALEX + "\n" + ALEX_TWIN + "\n")
        merged = ContactMerger().merge_contacts([a, b], "Alex Example")
        written = _written_lines([merged])

        for line in PRESERVED:
            self.assertIn(line, written)
        self.assertIn("NICKNAME:Lex", written)
        self.assertIn("X-CUSTOM-FIELD:beta", written)
        # The shared BDAY line is carried once, not twice.
        self.assertEqual(written.count("BDAY:1990-04-02"), 1)
        # Recognised fields are still merged by the existing logic.
        self.assertIn("TEL;TYPE=CELL:+1 555 0100", written)
        self.assertIn("TEL;TYPE=HOME:+1 555 0101", written)

    def test_recognised_lines_are_not_duplicated_as_extras(self):
        a, b = VCardParser().parse_content(ALEX + "\n" + ALEX_TWIN + "\n")
        merged = ContactMerger().merge_contacts([a, b], "Alex Example")
        written = _written_lines([merged])
        self.assertEqual(sum(1 for line in written if line.startswith("FN:")), 1)
        self.assertEqual(sum(1 for line in written if line.startswith("VERSION")), 1)
        self.assertEqual(sum(1 for line in written if line.startswith("BEGIN:VCARD")), 1)


if __name__ == "__main__":
    unittest.main()
