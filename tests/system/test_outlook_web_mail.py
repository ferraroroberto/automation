"""Offline tests for system/outlook_web.py and system/outlook_web_mail.py (no browser, no network)."""

import ast
import re
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from system import outlook_web as ow
from system import outlook_web_mail as owm


def _make_parts(folder: Path, numbers, total, stem="copilot.zip.b64") -> None:
    for n in numbers:
        (folder / f"{stem}.part{n:02d}of{total:02d}.txt").write_text("x")


def _args(argv):
    parser_args = []
    with mock.patch.object(owm, "execute", lambda *a, **k: parser_args.append((a, k)) or 0):
        code = owm.main(argv)
    return code, parser_args


class ParsePartSelectionTests(unittest.TestCase):
    def test_range_and_list(self):
        self.assertEqual(owm.parse_part_selection("8-10"), {8, 9, 10})
        self.assertEqual(owm.parse_part_selection("3, 5,9-11"), {3, 5, 9, 10, 11})
        self.assertEqual(owm.parse_part_selection("0"), {0})

    def test_rejects_garbage(self):
        for bad in ("", "a", "5-3", "4-", "-4", "1,,x", "1-2-3"):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                owm.parse_part_selection(bad)


class DiscoverPartsTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_complete_set_gets_two_digit_subjects_in_order(self):
        _make_parts(self.folder, range(1, 16), 15)
        parts = owm.discover_parts(self.folder, "*.txt", "part_")
        self.assertEqual([p.subject for p in parts], [f"part_{n:02d}" for n in range(1, 16)])

    def test_selection_returns_only_those_parts(self):
        _make_parts(self.folder, range(1, 16), 15)
        parts = owm.discover_parts(self.folder, "*.txt", "part_", {8, 9, 15})
        self.assertEqual([p.subject for p in parts], ["part_08", "part_09", "part_15"])

    def test_gap_is_refused_and_named(self):
        _make_parts(self.folder, [1, 2, 3, 4, 6], 6)
        with self.assertRaisesRegex(ValueError, r"missing parts: \[5\]"):
            owm.discover_parts(self.folder, "*.txt", "part_")

    def test_duplicate_number_is_refused(self):
        _make_parts(self.folder, [1, 2], 2)
        _make_parts(self.folder, [2], 2, stem="dup")
        with self.assertRaisesRegex(ValueError, "two files for part 2"):
            owm.discover_parts(self.folder, "*.txt", "part_")

    def test_mixed_totals_are_refused(self):
        _make_parts(self.folder, [1], 3, stem="a")
        _make_parts(self.folder, [2], 4, stem="b")
        with self.assertRaisesRegex(ValueError, "disagree on the total"):
            owm.discover_parts(self.folder, "*.txt", "part_")

    def test_unnumbered_file_is_refused(self):
        (self.folder / "notes.txt").write_text("x")
        with self.assertRaisesRegex(ValueError, "no part number"):
            owm.discover_parts(self.folder, "*.txt", "part_")

    def test_encode_and_split_naming_is_accepted_and_gaps_still_caught(self):
        for n in (1, 2, 3):
            (self.folder / f"copilot_{n:03d}.txt").write_text("x")
        parts = owm.discover_parts(self.folder, "*.txt", "part_")
        self.assertEqual([p.subject for p in parts], ["part_01", "part_02", "part_03"])
        (self.folder / "copilot_005.txt").write_text("x")
        with self.assertRaisesRegex(ValueError, r"missing parts: \[4\]"):
            owm.discover_parts(self.folder, "*.txt", "part_")

    def test_part_zero_helper_file_is_selectable(self):
        _make_parts(self.folder, [1, 2], 2)
        (self.folder / "join-and-decode_000.txt").write_text("x")
        parts = owm.discover_parts(self.folder, "*.txt", "part_", {0, 1, 2})
        self.assertEqual([p.subject for p in parts], ["part_00", "part_01", "part_02"])
        self.assertEqual([p.subject for p in owm.discover_parts(self.folder, "*.txt", "part_")],
                         ["part_01", "part_02"])


class BuildSpecsTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)
        self.a = self.dir / "report.pdf"
        self.b = self.dir / "notes.txt"
        self.a.write_text("a")
        self.b.write_text("b")

    def _specs(self, *argv):
        captured = []
        with mock.patch.object(owm, "execute", lambda specs, *a, **k: captured.extend(specs) or 0):
            code = owm.main(list(argv))
        return code, captured

    def test_single_file_defaults_to_its_name_as_subject(self):
        code, specs = self._specs(str(self.a), "--to", "x@y.z")
        self.assertEqual((code, [s.subject for s in specs]), (0, ["report.pdf"]))
        self.assertEqual(specs[0].attachments, (self.a,))

    def test_subject_template_and_body(self):
        _, specs = self._specs(str(self.a), str(self.b), "--to", "x@y.z", "--subject", "Q3 {stem}", "--body", "hi")
        self.assertEqual([s.subject for s in specs], ["Q3 report", "Q3 notes"])
        self.assertEqual(specs[0].body, "hi")

    def test_together_makes_one_draft_with_every_file(self):
        _, specs = self._specs(str(self.a), str(self.b), "--to", "x@y.z", "--together", "--subject", "Files")
        self.assertEqual(len(specs), 1)
        self.assertEqual(specs[0].attachments, (self.a, self.b))

    def test_ambiguous_requests_exit_two_without_running(self):
        cases = [
            [str(self.a), str(self.b), "--to", "x@y.z", "--subject", "same"],
            [str(self.a), str(self.b), "--to", "x@y.z", "--together"],
            [str(self.dir / "nope.bin"), "--to", "x@y.z"],
            [str(self.a), "--to", "x@y.z", "--subject", "{bad}"],
            [str(self.a), "--to", "x@y.z", "--parts", "1-2"],
            [str(self.dir), "--to", "x@y.z", "--subject", "x"],
        ]
        for argv in cases:
            with self.subTest(argv=argv):
                code, specs = self._specs(*argv)
                self.assertEqual((code, specs), (2, []))

    def test_body_file(self):
        body = self.dir / "body.txt"
        body.write_text("from file", encoding="utf-8")
        _, specs = self._specs(str(self.a), "--to", "x@y.z", "--body-file", str(body))
        self.assertEqual(specs[0].body, "from file")

    def test_dry_run_exits_zero_without_a_browser(self):
        parts = self.dir / "parts"
        parts.mkdir()
        _make_parts(parts, range(1, 4), 3)
        with mock.patch.object(owm, "OutlookWeb", side_effect=AssertionError("browser started")):
            self.assertEqual(owm.main([str(parts), "--to", "a@b.c", "--parts", "1-3", "--dry-run"]), 0)

    def test_yes_without_send_is_a_usage_error(self):
        with self.assertRaises(SystemExit):
            owm.main([str(self.a), "--to", "x@y.z", "--yes"])


class FakeWeb:
    """Stands in for OutlookWeb; records calls so the orchestration can be checked offline."""
    instance = None

    def __init__(self, url, headless, missing=(), signed_in=True, sent=(), lacking=None, repair_error=None):
        FakeWeb.instance = self
        self.calls, self.missing, self.signed_in = [], list(missing), signed_in
        self.sent = set(sent)
        self.lacking, self.repair_error = lacking or {}, repair_error

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def require_signed_in(self):
        if not self.signed_in:
            raise ow.NotSignedIn("no")

    def restore_subjects(self, expected):
        return 0

    def repair_attachments(self, specs):
        self.calls.append(("repair", [s.subject for s in specs]))
        if self.repair_error:
            raise self.repair_error
        return dict(self.lacking)

    def drafted_subjects(self, subjects):
        return set()

    def sent_subjects(self, subjects):
        return {s for s in subjects if s in self.sent}

    def draft(self, spec):
        self.calls.append(("draft", spec.subject))

    def verify(self, subjects):
        return list(self.missing)

    def send(self, spec):
        self.calls.append(("send", spec.subject))

    def screenshot(self, label):
        return None


class ExecuteTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.f = Path(self._tmp.name) / "a.txt"
        self.f.write_text("x")
        self.specs = [ow.DraftSpec("x@y.z", "a.txt", (self.f,))]

    def _run(self, send, yes, tty=False, **fake):
        factory = lambda url, headless: FakeWeb(url, headless, **fake)
        with mock.patch.object(owm, "OutlookWeb", factory), \
                mock.patch.object(owm.sys.stdin, "isatty", return_value=tty):
            return owm.execute(self.specs, "u", True, send, yes), FakeWeb.instance.calls

    def test_drafting_never_sends(self):
        self.assertEqual(self._run(False, False), (0, [("draft", "a.txt"), ("repair", ["a.txt"])]))

    def test_send_with_yes_sends_what_it_drafted(self):
        self.assertEqual(self._run(True, True),
                         (0, [("draft", "a.txt"), ("repair", ["a.txt"]), ("send", "a.txt")]))

    def test_a_draft_still_without_its_file_fails_the_run_and_is_never_sent(self):
        with self.assertLogs("outlook_web_mail", "ERROR") as logs:
            code, calls = self._run(True, True, lacking={"a.txt": ["a.txt"]})
        self.assertEqual(code, 1)
        self.assertNotIn(("send", "a.txt"), calls)
        self.assertIn("a.txt", "\n".join(logs.output))

    def test_a_count_mismatch_stops_the_run_without_sending(self):
        error = ow.DraftCountMismatch("scan 5, counter 3")
        with self.assertLogs("outlook_web_mail", "ERROR") as logs:
            code, calls = self._run(True, True, repair_error=error)
        self.assertEqual(code, 1)
        self.assertNotIn(("send", "a.txt"), calls)
        self.assertIn("scan 5, counter 3", "\n".join(logs.output))

    def test_send_without_confirmation_leaves_drafts_unsent(self):
        self.assertEqual(self._run(True, False, tty=False), (2, [("draft", "a.txt"), ("repair", ["a.txt"])]))

    def test_typed_confirmation(self):
        with mock.patch("builtins.input", return_value="SEND"):
            self.assertEqual(self._run(True, False, tty=True)[0], 0)
        with mock.patch("builtins.input", return_value="yes"):
            self.assertEqual(self._run(True, False, tty=True), (2, [("draft", "a.txt"), ("repair", ["a.txt"])]))

    def test_unverified_drafts_are_never_sent(self):
        code, calls = self._run(True, True, missing=["a.txt"])
        self.assertEqual(code, 1)
        self.assertNotIn(("send", "a.txt"), calls)

    def test_send_rerun_skips_subjects_already_in_sent_items(self):
        second = self.f.with_name("b.txt")
        second.write_text("y")
        self.specs = [ow.DraftSpec("x@y.z", "a.txt", (self.f,)), ow.DraftSpec("x@y.z", "b.txt", (second,))]
        code, calls = self._run(True, True, sent=["a.txt"])
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("draft", "b.txt"), ("repair", ["b.txt"]), ("send", "b.txt")])

    def test_send_rerun_with_everything_sent_does_nothing(self):
        self.assertEqual(self._run(True, True, sent=["a.txt"]), (0, []))

    def test_draft_only_run_does_not_consult_sent_items(self):
        self.assertEqual(self._run(False, False, sent=["a.txt"]),
                         (0, [("draft", "a.txt"), ("repair", ["a.txt"])]))

    def test_not_signed_in_exits_two(self):
        self.assertEqual(self._run(True, True, signed_in=False), (2, []))


class MergeSnapshotsTests(unittest.TestCase):
    def test_overlap_is_dropped(self):
        self.assertEqual(ow.merge_snapshots(["a", "b", "c"], ["b", "c", "d"]), ["a", "b", "c", "d"])

    def test_no_overlap_appends_everything(self):
        self.assertEqual(ow.merge_snapshots(["a", "b"], ["c", "d"]), ["a", "b", "c", "d"])

    def test_same_snapshot_adds_nothing(self):
        self.assertEqual(ow.merge_snapshots(["a", "b"], ["a", "b"]), ["a", "b"])

    def test_empty_sides(self):
        self.assertEqual(ow.merge_snapshots([], ["a"]), ["a"])
        self.assertEqual(ow.merge_snapshots(["a"], []), ["a"])

    def test_identical_rows_elsewhere_in_the_list_are_kept(self):
        # "x" appears twice, far apart: only the overlapping tail is dropped.
        self.assertEqual(ow.merge_snapshots(["x", "a", "b"], ["a", "b", "x"]), ["x", "a", "b", "x"])


class CleanRowTests(unittest.TestCase):
    def test_hover_icons_and_blank_lines_do_not_change_a_row(self):
        plain = "[Draft] ROBERTO FERRARO\nobs_05\n16:07\nGracias! Roberto"
        hovered = "\n[Draft] ROBERTO FERRARO\n\n\nobs_05\n\n16:07\nGracias! Roberto"
        self.assertEqual(ow.clean_row(hovered), plain)
        self.assertEqual(ow.clean_row(plain), plain)


class AvatarInitialsTests(unittest.TestCase):
    def test_avatar_initials_do_not_change_a_row(self):
        plain = "[Draft] SOME PERSON\npart_05\n16:07"
        for initials in ("RF", "R", "XY"):
            with self.subTest(initials=initials):
                self.assertEqual(ow.clean_row(f"{initials}\n{plain}"), plain)

    def test_only_whole_lines_of_one_or_two_capitals_are_dropped(self):
        row = "[Draft] SOME PERSON\nREPORT\nRe: Q3\nok\nA1"
        self.assertEqual(ow.clean_row(row), row)

    def test_snapshots_differing_only_by_avatar_merge_without_duplicates(self):
        rows = [f"[Draft] SOME PERSON\nobs_{n:02d}\n15:{n:02d}" for n in range(12)]
        first = [f"RF\n{r}" for r in rows[:8]]
        second = [f"R\n{r}" if n % 2 else r for n, r in enumerate(rows[4:])]
        merged = ow.merge_snapshots([ow.clean_row(r) for r in first], [ow.clean_row(r) for r in second])
        self.assertEqual(merged, rows)


class _Rows:
    def __init__(self, texts):
        self._texts = texts

    def all_inner_texts(self):
        return list(self._texts)


class VirtualListWeb(ow.OutlookWeb):
    """An OutlookWeb whose message list is virtualized like the real one: only ``window`` rows are
    visible at a time and wheeling moves the window. No browser."""
    ROW_PX = 80

    def __init__(self, rows, window=8):
        super().__init__()
        self.rows, self.window, self.top = rows, window, 0

    def _open_drafts(self):
        pass

    def _pause(self, ms):
        pass

    def _rows(self):
        return _Rows(self.rows[self.top:self.top + self.window])

    def _scroll_list(self, delta):
        limit = max(0, len(self.rows) - self.window)
        self.top = max(0, min(limit, self.top + delta // self.ROW_PX))


class ScanLongListTests(unittest.TestCase):
    @staticmethod
    def _rows(count):
        return [f"[Draft] ROBERTO FERRARO\nobs_{n:02d}\n15:{n:02d}" for n in range(count)]

    def test_every_row_of_a_long_list_is_read(self):
        rows = self._rows(29)
        self.assertEqual(VirtualListWeb(rows)._all_row_texts(), rows)

    def test_a_short_list_and_an_empty_list(self):
        rows = self._rows(3)
        self.assertEqual(VirtualListWeb(rows)._all_row_texts(), rows)
        self.assertEqual(VirtualListWeb([])._all_row_texts(), [])

    def test_scan_starts_from_the_top_even_when_scrolled_down(self):
        web = VirtualListWeb(self._rows(29))
        web.top = 21
        self.assertEqual(web._all_row_texts(), self._rows(29))

    def test_drafted_subjects_sees_rows_beyond_the_rendered_window(self):
        web = VirtualListWeb(self._rows(29))
        wanted = [f"obs_{n:02d}" for n in range(29)]
        self.assertEqual(web.drafted_subjects(wanted + ["obs_99"]), set(wanted))

    def test_count_drafts_finds_a_duplicate_outside_the_window(self):
        rows = self._rows(29) + ["[Draft] ROBERTO FERRARO\nobs_03\n16:40"]
        web = VirtualListWeb(rows)
        self.assertEqual(web._count_drafts("obs_03"), 2)
        self.assertEqual(web._count_drafts("obs_04"), 1)
        self.assertEqual(web._count_drafts("obs_50"), 0)


class FolderCountTests(unittest.TestCase):
    def test_counter_text_variants(self):
        for label, expected in (("Drafts 12", 12), ("Drafts [12]", 12), ("Drafts\n12", 12),
                                ("Drafts\n3", 3), ("drafts, 7 items", 7), ("Drafts", None),
                                ("Sent Items 4", None), ("", None)):
            with self.subTest(label=label):
                self.assertEqual(ow.parse_folder_count(label), expected)

    def test_count_state(self):
        self.assertEqual(ow.count_state(5, 5), "agree")
        self.assertEqual(ow.count_state(5, 3), "disagree")
        self.assertEqual(ow.count_state(0, None), "agree")
        self.assertEqual(ow.count_state(4, None), "unknown")


class SimMailbox(ow.OutlookWeb):
    """An OutlookWeb over an invented mailbox: ``drafts`` maps subject -> file names the saved draft
    carries. ``lose`` lists subjects whose next ``n`` saves drop their files. No browser."""

    def __init__(self, drafts, lose=None, counter="live"):
        super().__init__()
        self.drafts = {s: set(f) for s, f in drafts.items()}
        self.lose = dict(lose or {})
        self.counter, self.opened = counter, None
        self.log = []

    # the page, the pauses and the loads are not under test
    def refresh(self):
        pass

    def _open_drafts(self):
        pass

    def _pause(self, ms):
        pass

    def _all_row_texts(self):
        return [f"[Draft] SOME PERSON\n{s}\n10:00" for s in self.drafts]

    def _draft_counter(self):
        return len(self.drafts) if self.counter == "live" else self.counter

    def _open_draft_exact(self, subject):
        if subject not in self.drafts:
            raise RuntimeError(f"no draft with exactly the subject {subject!r} to open")
        self.opened = subject

    def _attachment_shown(self, name):
        return name in self.drafts[self.opened]

    def _click_discard(self):
        self.log.append(("discard", self.opened))
        del self.drafts[self.opened]

    def draft(self, spec):
        self.log.append(("draft", spec.subject))
        dropped = self.lose.get(spec.subject, 0) > 0
        if dropped:
            self.lose[spec.subject] -= 1
        self.drafts[spec.subject] = set() if dropped else {p.name for p in spec.attachments}


def _spec(subject, *files):
    return ow.DraftSpec("x@y.z", subject, tuple(Path(f) for f in files))


class AttachmentCheckTests(unittest.TestCase):
    def test_a_draft_without_its_file_is_named_with_the_missing_file(self):
        web = SimMailbox({"one": {"one.zip"}, "two": set(), "three": {"three.zip"}})
        specs = [_spec("one", "one.zip"), _spec("two", "two.zip"), _spec("three", "three.zip")]
        self.assertEqual(web.verify_attachments(specs), {"two": ["two.zip"]})

    def test_every_file_of_a_multi_file_draft_is_checked(self):
        web = SimMailbox({"both": {"a.zip"}})
        self.assertEqual(web.verify_attachments([_spec("both", "a.zip", "b.zip")]), {"both": ["b.zip"]})

    def test_complete_drafts_and_attachment_free_specs_pass(self):
        web = SimMailbox({"one": {"one.zip"}, "note": set()})
        self.assertEqual(web.verify_attachments([_spec("one", "one.zip"), _spec("note")]), {})

    def test_a_subject_that_is_not_exactly_one_row_is_an_error_not_a_pass(self):
        for drafts in ({}, {"dup": {"d.zip"}, "dup ": {"d.zip"}}):
            with self.subTest(drafts=drafts):
                web = SimMailbox(drafts)
                with self.assertRaisesRegex(RuntimeError, "expected exactly one draft"):
                    web.verify_attachments([_spec("dup", "d.zip")])


class CountCrossCheckTests(unittest.TestCase):
    def test_disagreement_raises_and_names_both_numbers(self):
        web = SimMailbox({"one": {"one.zip"}, "two": set()}, counter=5)
        with self.assertRaisesRegex(ow.DraftCountMismatch, "2 rows.*counter says 5"):
            web.verify_attachments([_spec("one", "one.zip")])

    def test_a_disagreement_never_discards_or_redrafts(self):
        web = SimMailbox({"two": set()}, counter=9)
        with self.assertRaises(ow.DraftCountMismatch):
            web.repair_attachments([_spec("two", "two.zip")])
        self.assertEqual(web.log, [])
        self.assertIn("two", web.drafts)

    def test_a_lagging_counter_that_catches_up_on_the_re_read_is_accepted(self):
        web = SimMailbox({"one": {"one.zip"}, "two": {"two.zip"}})
        reads = iter([1, 2])
        web._draft_counter = lambda: next(reads)
        self.assertEqual(web.verify_attachments([_spec("one", "one.zip")]), {})

    def test_an_unreadable_counter_blocks_a_discard_but_not_a_read_only_check(self):
        web = SimMailbox({"two": set()}, counter=None)
        with self.assertLogs("outlook_web", "WARNING"):
            self.assertEqual(web.verify_attachments([_spec("two", "two.zip")]), {"two": ["two.zip"]})
        with self.assertRaisesRegex(ow.DraftCountMismatch, "unreadable"), self.assertLogs("outlook_web", "WARNING"):
            web.discard_draft("two")
        self.assertEqual(web.log, [])

    def test_an_empty_folder_with_no_counter_is_consistent(self):
        web = SimMailbox({}, counter=None)
        self.assertEqual(web._check_counts(strict=True), [])


class RepairAttachmentsTests(unittest.TestCase):
    def test_only_the_draft_lacking_its_file_is_redone(self):
        web = SimMailbox({"one": {"one.zip"}, "two": set(), "three": {"three.zip"}})
        specs = [_spec("one", "one.zip"), _spec("two", "two.zip"), _spec("three", "three.zip")]
        self.assertEqual(web.repair_attachments(specs), {})
        self.assertEqual(web.log, [("discard", "two"), ("draft", "two")])
        self.assertEqual(web.drafts["two"], {"two.zip"})
        self.assertEqual(set(web.drafts), {"one", "two", "three"})

    def test_a_healthy_set_is_left_untouched(self):
        web = SimMailbox({"one": {"one.zip"}})
        self.assertEqual(web.repair_attachments([_spec("one", "one.zip")]), {})
        self.assertEqual(web.log, [])

    def test_the_redo_is_capped_and_the_draft_is_reported(self):
        web = SimMailbox({"two": set()}, lose={"two": 99})
        with self.assertLogs("outlook_web", "WARNING"):
            result = web.repair_attachments([_spec("two", "two.zip")], redo_cap=2)
        self.assertEqual(result, {"two": ["two.zip"]})
        self.assertEqual(web.log, [("discard", "two"), ("draft", "two")] * 2)

    def test_a_second_attempt_can_succeed(self):
        web = SimMailbox({"two": set()}, lose={"two": 1})
        with self.assertLogs("outlook_web", "WARNING"):
            self.assertEqual(web.repair_attachments([_spec("two", "two.zip")], redo_cap=3), {})
        self.assertEqual(web.log, [("discard", "two"), ("draft", "two")] * 2)

    def test_a_discard_that_removes_more_than_that_draft_aborts_the_redo(self):
        web = SimMailbox({"one": {"one.zip"}, "two": set()})

        def discard_both():
            web.log.append(("discard", "both"))
            web.drafts.clear()
        web._click_discard = discard_both
        with self.assertLogs("outlook_web", "WARNING"), self.assertRaisesRegex(RuntimeError, "did not remove exactly"):
            web.repair_attachments([_spec("one", "one.zip"), _spec("two", "two.zip")])
        self.assertEqual(web.log, [("discard", "both")])

    def test_a_discard_that_removes_nothing_aborts_the_redo(self):
        web = SimMailbox({"two": set()})
        web._click_discard = lambda: web.log.append(("discard", "nothing"))
        with self.assertLogs("outlook_web", "WARNING"), self.assertRaisesRegex(RuntimeError, "did not remove exactly"):
            web.repair_attachments([_spec("two", "two.zip")])
        self.assertEqual(web.log, [("discard", "nothing")])


class AttachmentLocatorTests(unittest.TestCase):
    def test_list_rows_are_filtered_out_of_the_attachment_options(self):
        web = ow.OutlookWeb()
        filtered = web.page = mock.MagicMock()
        options = filtered.get_by_role.return_value
        self.assertTrue(web._attachment_shown("report.pdf"))
        options.filter.assert_called_once_with(has_not_text="[Draft]")
        options.filter.return_value.first.wait_for.side_effect = TimeoutError
        self.assertFalse(web._attachment_shown("report.pdf"))


class _Button:
    def __init__(self, page, name):
        self.page, self.name = page, name

    @property
    def first(self):
        return self

    def is_visible(self):
        return self.page.home_selected

    def click(self, timeout=30_000):
        self.page.clicks.append(self.name)
        if self.name == "Home":
            self.page.home_selected = True


class _Page:
    """Ribbon stand-in: "New mail" is visible only while the Home tab is selected."""

    def __init__(self, home_selected, has_home_tab=True):
        self.home_selected, self.has_home_tab, self.clicks = home_selected, has_home_tab, []

    def get_by_role(self, role, name=None, **kwargs):
        if role == "tab" and not self.has_home_tab:
            raise RuntimeError("no Home tab")
        return _Button(self, name if role == "tab" else "New mail")


class OpenComposeTests(unittest.TestCase):
    def _web(self, **page):
        web = ow.OutlookWeb()
        web.page = _Page(**page)
        web._open_drafts = lambda: None
        web._subject_box = lambda: mock.Mock()
        return web

    def test_message_ribbon_is_switched_to_home_before_new_mail(self):
        web = self._web(home_selected=False)
        web._open_compose()
        self.assertEqual(web.page.clicks, ["Home", "New mail"])

    def test_home_ribbon_is_left_alone(self):
        web = self._web(home_selected=True)
        web._open_compose()
        self.assertEqual(web.page.clicks, ["New mail"])

    def test_missing_home_tab_still_tries_new_mail(self):
        web = self._web(home_selected=False, has_home_tab=False)
        web._open_compose(attempts=1)
        self.assertEqual(web.page.clicks, ["New mail"])


class StructureTests(unittest.TestCase):
    def test_draftspec_requires_recipient_and_subject(self):
        for to, subject in (("", "s"), ("a@b.c", "")):
            with self.assertRaises(ValueError):
                ow.DraftSpec(to, subject)

    def test_only_the_send_method_can_click_send(self):
        tree = ast.parse(Path(ow.__file__).read_text(encoding="utf-8"))
        holders = set()
        for func in (n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)):
            for const in (n for n in ast.walk(func) if isinstance(n, ast.Constant) and isinstance(n.value, str)):
                if re.search(r"\^Send\$|Control\+(Enter|Return)", const.value):
                    holders.add(func.name)
        self.assertEqual(holders, {"send"})

    def test_cli_never_sends_unless_asked(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "a.txt"
            f.write_text("x")
            _, calls = _args([str(f), "--to", "a@b.c"])
            self.assertIs(calls[0][0][3], False)
            _, calls = _args([str(f), "--to", "a@b.c", "--send", "--yes"])
            self.assertIs(calls[0][0][3], True)


if __name__ == "__main__":
    unittest.main()
