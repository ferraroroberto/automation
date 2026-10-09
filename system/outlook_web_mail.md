# Outlook Web Mail

Drafts Outlook-on-the-web messages with any files attached and, only when asked, sends them. Two pieces:

- `system/outlook_web.py` is the reusable library (a `DraftSpec` plus an `OutlookWeb` browser session). Import it from any script.
- `system/outlook_web_mail.py` is the command-line tool built on it.

It started as a way to relay a large file past a firewall by email: split it with `base64_encode_decode.py` ("Encode & Split File"), draft the parts, send them, then join and decode on the other side. It now handles one file, several files, or a folder of numbered parts.

## One-time login

```powershell
& .\.venv\Scripts\python.exe -m system.outlook_web_mail --login
```

A visible Chrome opens on the mailbox. Sign in by hand (SSO and MFA included); the window closes once the mailbox loads (10 minutes max). The script never sees or stores your password. The signed-in profile lives outside the repo, in `OUTLOOK_WEB_PROFILE_DIR` (default `~/.outlook-web-drafts-profile`).

## Command line

```powershell
# One file, one draft (subject defaults to the file name)
& .\.venv\Scripts\python.exe -m system.outlook_web_mail report.pdf --to someone@example.com --subject "Q3 report" --body "Hi, attached."

# Several files, one draft each (subject template; default {name})
& .\.venv\Scripts\python.exe -m system.outlook_web_mail a.zip b.pdf --to someone@example.com --subject "Delivery {stem}"

# Several files in ONE message
& .\.venv\Scripts\python.exe -m system.outlook_web_mail a.zip b.pdf --to someone@example.com --together --subject "The files"

# A folder of numbered parts: subject part_NN, one draft per part
& .\.venv\Scripts\python.exe -m system.outlook_web_mail E:\path\to\parts --to someone@example.com --parts 8-15

# Look first; no browser is started
... --dry-run
```

| Option | Meaning |
|---|---|
| `sources` | A folder of numbered parts, or one or more files. |
| `--to` | Recipient address (required; never stored). |
| `--subject` | Subject; may use `{name}` and `{stem}`. Needed with `--together`. |
| `--body`, `--body-file` | Text placed above the signature. |
| `--together` | One draft carrying every file. |
| `--parts` | Folder mode: only these parts, e.g. `8-15` or `3,5,9-11`. They must all exist. |
| `--glob` | Folder mode: file pattern (default `*.txt`). |
| `--subject-prefix` | Folder mode: subject is the prefix plus the two-digit part number (default `part_`). |
| `--send` | After verifying the drafts, send exactly those drafts (see below). |
| `--yes` | With `--send`: skip the typed confirmation. |
| `--login`, `--dry-run`, `--headless`, `--url` | Sign in once; list only; no window; another mail URL. |

Folder mode takes part numbers from `…partNNofMM…` names (the total is then checked) or `…_NNN.txt` names (gaps below the highest number are caught). A helper file named `…_000.txt` is part 0: draft it with `--parts 0`, or `--parts 0-15` with the rest. Anything ambiguous (a gap, a duplicate number, two drafts with the same subject, a missing file) is refused with exit 2 before the browser starts.

## Sending

Drafting never sends. `--send` is the only way mail leaves, and it goes through the same checks every time:

1. The drafts are created (or found already in Drafts) and verified after a fresh reload.
2. You must type `SEND` at the prompt, or pass `--yes`. With no terminal and no `--yes` it refuses (exit 2) and leaves the drafts unsent.
3. For each message, `send` refuses unless exactly one draft has that subject, it is addressed to `--to`, and its attachments are present. After clicking Send it waits for the draft to leave Drafts and for the subject to appear in Sent Items; if it cannot confirm that, it fails and tells you to check the mailbox.

A re-run only sends drafts that are still in Drafts, so already-sent messages are not sent twice.

## Library use

```python
from pathlib import Path
from system.outlook_web import DraftSpec, OutlookWeb

spec = DraftSpec(to="someone@example.com", subject="report", attachments=(Path("report.pdf"),), body="Hi")
with OutlookWeb() as web:
    web.require_signed_in()
    web.draft(spec)
    assert not web.verify([spec.subject])
    # web.send(spec)  # explicit; drafting never implies it
```

`DraftSpec(to, subject, attachments=(), body="")` is the whole input. `OutlookWeb` also offers `drafted_subjects`, `verify`, `restore_subjects` and `screenshot`.

## What a run does

1. Opens the mailbox with the saved profile; exits 2 if it is not signed in. It restores the subject on any subject-less draft carrying one of this run's files (left by an earlier run) and skips drafts already present, so a re-run never duplicates. Other drafts of yours are left alone.
2. For each remaining draft, always from the Drafts list (where "New mail" is): attaches each file as a copy (choosing it when Outlook offers a OneDrive link), then body, recipient and subject, then saves. The subject goes in last on purpose: one typed before the upload finishes was lost in manual runs.
3. Reloads and checks every subject is in Drafts, repairing a subject-less draft from its attachment name. Exit 0 only when all are verified, otherwise 1 with a screenshot under `<profile>/failures/`.

Exit codes: `0` ok, `1` a draft or send failed or could not be verified, `2` usage error, refused request, not signed in, or send not confirmed.

## Caveats

- Outlook on the web is not under our control. If its page changes, a selector can stop matching; the run then fails with an error and a screenshot rather than skipping silently.
- Outlook often saves a new draft without its subject on the first pass; the repair step exists for that.
- Each 9 MB attachment uploads before the next draft starts, so a long run takes a while.

## Tests

`python -m unittest system.test_outlook_web_mail` covers part discovery, selection parsing, request building, the send logic against a fake browser (never sends unless asked, needs confirmation, never sends unverified drafts), and a guard that only the library's `send` method can click Send. The real browser steps are checked by live runs.
