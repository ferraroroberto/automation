# Git Email Rewrite — 2026-05-10

## Problem

`C:\Users\rober\.gitconfig` had the wrong email:

```
user.email = roberto.ferraro-at-gmail.com
```

The `-at-` was an anti-spam obfuscation style accidentally applied to the git identity.
GitHub only counts contributions from commits whose author email matches a verified account email,
so ~700 commits across 25 repos were invisible on the contribution graph.

## What was done

### 1. Fixed the global git identity

```
C:\Users\rober\.gitconfig
  user.email = 35553560+ferraroroberto@users.noreply.github.com
```

This is GitHub's privacy-preserving noreply format. It counts toward contributions without
exposing the real email address in public commit history.

### 2. Added a global pre-commit hook

Location: `C:\Users\rober\.githooks\pre-commit`
Activated via: `git config --global core.hooksPath C:/Users/rober/.githooks`

The hook blocks any commit whose author email isn't the noreply address, with a clear error
message and the fix command. Bypassable with `git commit --no-verify` when intentional.

### 3. Rewrote commit history across 25 repos

Tool: `git filter-repo` (installed via `pip install git-filter-repo`)

Emails remapped in every commit (author + committer):
- `roberto.ferraro-at-gmail.com`  → noreply
- `roberto.ferraro@gmail.com`     → noreply
- `rferraro@caixabank.com`        → noreply

Target noreply: `35553560+ferraroroberto@users.noreply.github.com`

Repos rewritten and force-pushed:

| Repo | Notes |
|---|---|
| accounting-quarterly | |
| arboldelossuenos | |
| automation | |
| closed-company-accounting | |
| copilot-studio-transcripts | |
| email-archiver | |
| externalrisk | |
| facilitation-shuffle | |
| family-accounting | |
| grocery-shopping-automation | |
| illustration-color-edit | |
| inspiration-system | |
| local-llm-hub | |
| mass-html-to-markdown | |
| mathgamesforkids | |
| mcp-personal-onedrive | |
| old\email-automation | same remote as `automation` |
| pdf-to-markdown | |
| project-scaffolding | |
| reporting | |
| social-media-analytics | |
| vibe-coding-workshop | |
| voice-transcriber | smoke-test repo, done first |
| website | |
| whisper-iphone-keyboard | |
| work | |

Skipped: `suna\suna` — third-party fork with hundreds of upstream authors.

Emails not remapped (left as-is):
- `noreply@anthropic.com` — Claude Code agent commits
- `cursoragent@cursor.com` — Cursor agent commits
- `*+Copilot@users.noreply.github.com` — GitHub Copilot commits
- `35553560+ferraroroberto@users.noreply.github.com` — already correct

### 4. Verification

Final scan across all 26 repos confirmed zero remaining bad-email commits.

## How to redo this on a single repo (if needed)

```powershell
# Prerequisites: pip install git-filter-repo

$remote = git remote get-url origin
git filter-repo --force --email-callback "bad = [b'roberto.ferraro-at-gmail.com', b'roberto.ferraro@gmail.com', b'rferraro@caixabank.com']
return b'35553560+ferraroroberto@users.noreply.github.com' if email in bad else email"
git remote add origin $remote
git push --force origin --all
git push --force origin --tags
```

Key gotcha: if filter-repo was previously run (killed mid-run or re-run), it leaves a
`.git\filter-repo\already_ran` marker that causes an interactive Y/N prompt in non-TTY
sessions. Delete it before re-running:

```powershell
Remove-Item .git\filter-repo\already_ran -Force -ErrorAction SilentlyContinue
```

## Prevention

Two layers:

**Layer 1 — correct global config** (already applied):
```
git config --global user.email "35553560+ferraroroberto@users.noreply.github.com"
```

**Layer 2 — pre-commit hook** at `C:\Users\rober\.githooks\pre-commit`:
```sh
#!/bin/sh
ALLOWED="35553560+ferraroroberto@users.noreply.github.com"
AUTHOR_EMAIL=$(git var GIT_AUTHOR_IDENT | sed 's/.*<\(.*\)>.*/\1/')
if [ "$AUTHOR_EMAIL" != "$ALLOWED" ]; then
  echo "COMMIT BLOCKED: author email '$AUTHOR_EMAIL' is not allowed."
  echo "Fix: git config user.email \"$ALLOWED\""
  exit 1
fi
```

Bypass for intentional one-offs: `git commit --no-verify`

## Setting up a new or additional machine

The rewrite fixed the history on GitHub (server-side). Each machine you commit from also needs
the same local configuration, otherwise new commits from that machine will use the wrong email
and be invisible on the contribution graph again.

Run this on every machine:

```powershell
# 1. Fix the global email
git config --global user.email "35553560+ferraroroberto@users.noreply.github.com"
git config --global user.name "Roberto Ferraro"

# 2. Create the hooks directory and pre-commit hook
New-Item -ItemType Directory -Force -Path "$HOME\.githooks" | Out-Null
@'
#!/bin/sh
ALLOWED="35553560+ferraroroberto@users.noreply.github.com"
AUTHOR_EMAIL=$(git var GIT_AUTHOR_IDENT | sed 's/.*<\(.*\)>.*/\1/')
if [ "$AUTHOR_EMAIL" != "$ALLOWED" ]; then
  echo ""
  echo "  COMMIT BLOCKED: author email '$AUTHOR_EMAIL' is not on the allowlist."
  echo "  Expected: $ALLOWED"
  echo ""
  echo "  Fix with:"
  echo "    git config user.email \"$ALLOWED\""
  echo ""
  echo "  Or bypass once (if intentional):"
  echo "    git commit --no-verify"
  echo ""
  exit 1
fi
'@ | Set-Content "$HOME\.githooks\pre-commit" -Encoding utf8

# 3. Point git to the hooks directory
git config --global core.hooksPath "$HOME/.githooks"

# 4. Scan for any repo-level overrides that would shadow the global config
Get-ChildItem E:\automation -Directory | Where-Object { Test-Path "$($_.FullName)\.git" } | ForEach-Object {
  $email = git -C $_.FullName config --local user.email 2>$null
  if ($email) { Write-Output "LOCAL OVERRIDE in $($_.Name): $email — unset with: git -C '$($_.FullName)' config --local --unset user.email" }
}
```

Note: the history rewrite (force-push) already happened from this machine and is on GitHub.
Other machines just need the config above — they do **not** need to re-run `git filter-repo`.
After pulling/cloning on the new machine, `git log` will show the rewritten history automatically.
