"""Draft (and, only when asked, send) Outlook-on-the-web mail with files attached.

    python -m system.outlook_web_mail --login                        # once: sign in by hand
    python -m system.outlook_web_mail report.pdf --to ADDR [--subject "Q3 report"] [--body "Hi"]
    python -m system.outlook_web_mail a.zip b.pdf --to ADDR          # one draft per file, subject = file name
    python -m system.outlook_web_mail a.zip b.pdf --to ADDR --together --subject "Files"   # one draft, two files
    python -m system.outlook_web_mail PARTS_DIR --to ADDR [--parts 8-15]   # numbered parts: subject part_NN
    ... --dry-run                                                    # list the drafts, no browser
    ... --send [--yes]                                               # after verifying, send exactly those drafts

Drafting never sends. ``--send`` sends only the drafts this run verified, after a typed confirmation
(or ``--yes``). Run from the repo root with the repo's .venv; the library is ``system/outlook_web.py``.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from system.outlook_web import (OUTLOOK_URL, DraftSpec, NotSignedIn, OutlookWeb)

logger = logging.getLogger("outlook_web_mail")

PART_RE = re.compile(r"part(\d+)of(\d+)", re.IGNORECASE)
SPLIT_RE = re.compile(r"_(\d{3})\.txt", re.IGNORECASE)  # base64_encode_decode.py's "Encode & Split"


@dataclass(frozen=True)
class Part:
    number: int
    path: Path
    subject: str


def parse_part_selection(text: str) -> set[int]:
    """'8-15' or '3,5,9-11' -> the set of part numbers."""
    numbers: set[int] = set()
    for chunk in text.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        low, dash, high = chunk.partition("-")
        if not low.strip().isdigit() or (dash and not high.strip().isdigit()):
            raise ValueError(f"bad part selection: {chunk!r}")
        start, end = int(low), int(high) if dash else int(low)
        if end < start:
            raise ValueError(f"bad part range: {chunk!r}")
        numbers.update(range(start, end + 1))
    if not numbers:
        raise ValueError("empty part selection")
    return numbers


def _number_and_total(name: str) -> tuple[int, Optional[int]]:
    """Part number (and total when the name states it) from ``...partNNofMM...`` or ``..._NNN.txt``."""
    match = PART_RE.search(name)
    if match:
        return int(match.group(1)), int(match.group(2))
    match = SPLIT_RE.search(name)
    if match:
        return int(match.group(1)), None
    raise ValueError(f"no part number (partNNofMM or _NNN.txt) in file name: {name}")


def discover_parts(folder: Path, pattern: str, prefix: str,
                   selection: Optional[set[int]] = None) -> list[Part]:
    """Numbered parts in ``folder``, sorted; refuses gaps, duplicate numbers and mixed totals."""
    if not folder.is_dir():
        raise ValueError(f"not a folder: {folder}")
    found: dict[int, Part] = {}
    totals: set[int] = set()
    for path in sorted(p for p in folder.glob(pattern) if p.is_file()):
        number, total = _number_and_total(path.name)
        if number in found:
            raise ValueError(f"two files for part {number}: {found[number].path.name}, {path.name}")
        if total is not None:
            totals.add(total)
        found[number] = Part(number, path, f"{prefix}{number:02d}")
    if not found:
        raise ValueError(f"no files matching {pattern!r} in {folder}")
    if len(totals) > 1:
        raise ValueError(f"parts disagree on the total: {sorted(totals)}")
    total = next(iter(totals)) if totals else max(found)
    wanted = selection if selection is not None else set(range(1, total + 1))
    missing = sorted(wanted - found.keys())
    if missing:
        raise ValueError(f"missing parts: {missing}")
    return [found[n] for n in sorted(wanted)]


def build_specs(args: argparse.Namespace) -> list[DraftSpec]:
    """The drafts a command line asks for; raises ValueError on anything ambiguous or missing."""
    body = args.body or ""
    if args.body_file:
        body = Path(args.body_file).read_text(encoding="utf-8")
    sources = [Path(s) for s in args.sources]
    if len(sources) == 1 and sources[0].is_dir():
        if args.subject or args.together:
            raise ValueError("--subject and --together do not apply to a folder of numbered parts")
        selection = parse_part_selection(args.parts) if args.parts else None
        parts = discover_parts(sources[0], args.glob, args.subject_prefix, selection)
        return [DraftSpec(args.to, p.subject, (p.path,), body) for p in parts]
    if args.parts:
        raise ValueError("--parts only applies to a folder of numbered parts")
    for path in sources:
        if not path.is_file():
            raise ValueError(f"not a file: {path}")
    if args.together:
        if not args.subject:
            raise ValueError("--together needs --subject")
        specs = [DraftSpec(args.to, args.subject, tuple(sources), body)]
    else:
        template = args.subject or "{name}"
        try:
            specs = [DraftSpec(args.to, template.format(name=p.name, stem=p.stem), (p,), body) for p in sources]
        except (KeyError, IndexError) as exc:
            raise ValueError(f"--subject may only use {{name}} and {{stem}}: {exc}") from exc
    subjects = [s.subject for s in specs]
    if len(set(subjects)) != len(subjects):
        raise ValueError("two drafts would share a subject; use a {name} or {stem} template in --subject")
    return specs


def confirm_send(count: int, to: str, assume_yes: bool, interactive: bool) -> bool:
    """Sending needs a typed SEND, or --yes; with no terminal and no --yes the answer is no."""
    if assume_yes:
        return True
    if not interactive:
        return False
    return input(f"Type SEND to send {count} message(s) to {to}: ").strip() == "SEND"


def execute(specs: list[DraftSpec], url: str, headless: bool, send: bool, assume_yes: bool) -> int:
    subjects = [s.subject for s in specs]
    expected = {a.name: s.subject for s in specs for a in s.attachments}
    with OutlookWeb(url, headless) as web:
        try:
            web.require_signed_in()
        except NotSignedIn:
            logger.error("❌ not signed in; run with --login first")
            return 2
        if web.restore_subjects(expected):
            logger.info("ℹ️ restored a subject on a draft left by an earlier run")
        already = web.drafted_subjects(subjects)
        if already:
            logger.info("ℹ️ already in Drafts, skipped: %s", sorted(already))
        for spec in (s for s in specs if s.subject not in already):
            try:
                web.draft(spec)
            except Exception as exc:
                logger.error("❌ %s failed: %s", spec.subject, exc)
                logger.error("❌ screenshot: %s", web.screenshot(spec.subject))
                return 1
            logger.info("✅ %s drafted", spec.subject)
        missing = web.verify(subjects)
        if missing:
            logger.warning("⚠️ no subject saved for %s; repairing", missing)
            web.restore_subjects(expected)
            missing = web.verify(subjects)
        if missing:
            logger.error("❌ still missing from Drafts: %s", missing)
            logger.error("❌ screenshot: %s", web.screenshot("verify"))
            return 1
        logger.info("✅ all %d drafts are in Drafts with their subjects", len(specs))
        if not send:
            logger.info("ℹ️ nothing was sent")
            return 0
        if not confirm_send(len(specs), specs[0].to, assume_yes, sys.stdin.isatty()):
            logger.error("❌ send not confirmed (type SEND at the prompt, or pass --yes); drafts left unsent")
            return 2
        for spec in specs:
            try:
                web.send(spec)
            except Exception as exc:
                logger.error("❌ sending %s failed: %s", spec.subject, exc)
                logger.error("❌ screenshot: %s", web.screenshot("send-" + spec.subject))
                return 1
            logger.info("✅ %s sent", spec.subject)
        return 0


def run_login(url: str) -> int:
    with OutlookWeb(url, headless=False) as web:
        logger.info("ℹ️ Sign in by hand (SSO/MFA). The window closes once the mailbox loads.")
        if not web.login(600):
            logger.error("❌ the mailbox did not load within 600 s")
            return 1
    logger.info("✅ signed in; profile saved")
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Draft (and optionally send) Outlook-on-the-web mail with attachments.")
    parser.add_argument("sources", nargs="*", help="a folder of numbered parts, or one or more files")
    parser.add_argument("--to", help="recipient address")
    parser.add_argument("--subject", help="subject; may use {name} / {stem}; default {name}")
    parser.add_argument("--body", help="text placed above the signature")
    parser.add_argument("--body-file", help="read the body text from this file")
    parser.add_argument("--together", action="store_true", help="one draft carrying all the files")
    parser.add_argument("--parts", help="folder mode: only these parts, e.g. 8-15 or 3,5,9-11")
    parser.add_argument("--glob", default="*.txt", help="folder mode: file pattern (default *.txt)")
    parser.add_argument("--subject-prefix", default="part_", help="folder mode: prefix of the NN subject")
    parser.add_argument("--url", default=OUTLOOK_URL, help="Outlook on the web mail URL")
    parser.add_argument("--login", action="store_true", help="open a visible Chrome to sign in once")
    parser.add_argument("--dry-run", action="store_true", help="list the drafts; no browser")
    parser.add_argument("--headless", action="store_true", help="run without a visible window")
    parser.add_argument("--send", action="store_true", help="after verifying, send exactly these drafts")
    parser.add_argument("--yes", action="store_true", help="with --send: skip the typed confirmation")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if args.login:
        return run_login(args.url)
    if not args.sources or not args.to:
        parser.error("a source (folder or files) and --to are required (or use --login)")
    if args.yes and not args.send:
        parser.error("--yes only applies with --send")
    try:
        specs = build_specs(args)
    except ValueError as exc:
        logger.error("❌ %s", exc)
        return 2
    for spec in specs:
        names = ", ".join(a.name for a in spec.attachments)
        logger.info("ℹ️ %s <- %s to %s", spec.subject, names, spec.to)
    if args.dry_run:
        return 0
    return execute(specs, args.url, args.headless, args.send, args.yes)


if __name__ == "__main__":
    raise SystemExit(main())
