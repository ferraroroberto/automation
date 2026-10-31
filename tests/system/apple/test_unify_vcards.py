"""Regression tests for unify_vcards field preservation (synthetic contacts only)."""

import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "system/apple"))

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


class LowercaseTypeParameterTests(unittest.TestCase):
    """iCloud exports write `type=`; the parameter name is case-insensitive."""

    def test_label_is_read_from_a_lowercase_parameter(self):
        from _vcard_fields import extract_type_label

        self.assertEqual(extract_type_label("TEL;type=CELL;type=pref"), "CELL")
        self.assertEqual(extract_type_label("EMAIL;type=work", legacy_positional_fallback=False), "WORK")

    def test_convert_to_apple_keeps_a_lowercase_typed_line(self):
        from convert_to_apple import AppleContactConverter

        converter = AppleContactConverter.__new__(AppleContactConverter)
        line = "TEL;type=CELL;type=pref:+1 555 0100"
        converted = converter._format_phone_line(line)
        self.assertNotIn("TYPE=TYPE=", converted)
        self.assertTrue(converted.lower().startswith("tel;type=cell"))

    def test_writer_does_not_double_the_prefix_for_a_lowercase_typed_phone(self):
        raw = "\n".join([
            "BEGIN:VCARD", "VERSION:3.0", "FN:Sam Example", "N:Example;Sam;;;",
            "TEL;type=CELL;type=pref:+1 555 0100", "END:VCARD",
        ])
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "in.vcf")
            Path(path).write_text(raw, encoding="utf-8")
            contacts = VCardParser().parse_file(path)
            out = os.path.join(tmp, "out.vcf")
            VCardWriter().write_contacts(contacts, out)
            written = Path(out).read_text(encoding="utf-8")
        self.assertNotIn("TYPE=TYPE=", written)


if __name__ == "__main__":
    unittest.main()
