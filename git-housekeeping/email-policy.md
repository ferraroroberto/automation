# Git Email Identity Policy

This repo (and all sibling repos under `E:\automation\`) use GitHub's privacy-preserving noreply address as the canonical git author email:

```
35553560+ferraroroberto@users.noreply.github.com
```

GitHub counts contributions from commits whose author email matches a verified account email. The noreply format satisfies that requirement without exposing the real address in public commit history.

## Prevention: two layers

**Layer 1 — correct global config** (apply once per machine):

```powershell
git config --global user.email "35553560+ferraroroberto@users.noreply.github.com"
git config --global user.name "Roberto Ferraro"
```

**Layer 2 — pre-commit hook** at `C:\Users\rober\.githooks\pre-commit`, activated via:

```powershell
git config --global core.hooksPath "$HOME/.githooks"
```

The hook compares `git var GIT_AUTHOR_IDENT` against the allowlisted noreply address and exits 1 with a fix hint when they differ, so a mis-configured machine can never land a commit under the wrong identity. Its body is written by step 2 of [Setting up a new machine](#setting-up-a-new-machine) below — that script is the single copy in this document; don't transcribe a second one here.

Bypass for intentional one-offs: `git commit --no-verify`

## Setting up a new machine

Run this on every machine you commit from:

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

Note: new machines only need the config above — they do **not** need to re-run `git filter-repo`. After pulling or cloning, `git log` shows the already-rewritten history automatically.

## Rewriting commit history on a single repo

Use this if a repo accumulates bad-email commits (e.g. after a new machine was set up without Layer 1/2):

Fill the `bad` list in locally before running — the whole point of the rewrite is to get those addresses *out* of public history, so this repo does not carry a literal copy of them. Get the actual set from the repo you are about to rewrite:

```powershell
# Which author/committer emails does this repo's history actually contain?
git log --all --format='%ae%n%ce' | Sort-Object -Unique
```

Then paste the ones to remap into the callback (leave the exclusions below as-is):

```powershell
# Prerequisites: pip install git-filter-repo

$remote = git remote get-url origin
git filter-repo --force --email-callback "bad = [b'OLD_PERSONAL_ADDRESS', b'OLD_EMPLOYER_ADDRESS']
return b'35553560+ferraroroberto@users.noreply.github.com' if email in bad else email"
git remote add origin $remote
git push --force origin --all
git push --force origin --tags
```

Key gotcha: if `filter-repo` was previously run (killed mid-run or re-run), it leaves a `.git\filter-repo\already_ran` marker that causes an interactive Y/N prompt in non-TTY sessions. Delete it before re-running:

```powershell
Remove-Item .git\filter-repo\already_ran -Force -ErrorAction SilentlyContinue
```

Emails to exclude from remapping (leave as-is):
- `noreply@anthropic.com` — Claude Code agent commits
- `cursoragent@cursor.com` — Cursor agent commits
- `*+Copilot@users.noreply.github.com` — GitHub Copilot commits
- `35553560+ferraroroberto@users.noreply.github.com` — already correct
