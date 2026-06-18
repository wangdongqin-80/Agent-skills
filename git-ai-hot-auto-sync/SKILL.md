---
name: git-ai-hot-auto-sync
description: Use when Codex must create or run automatic Git collaboration sync for wangdongqin-80/AI-HOT, including scheduled checks, file-change watching, pulling remote updates, committing and pushing local changes, cloning the configured repository, and notifying users about merge conflicts, divergence, authentication failures, wrong remotes, or unsafe repository states.
---

# Git AI-HOT Auto Sync

## Overview

Synchronize `https://github.com/wangdongqin-80/AI-HOT` with a remote collaborator by running a conservative Git loop: commit local changes, fetch, pull clean remote updates, push local commits, and stop with a notification when human judgment is required.

Use `scripts/sync_ai_hot.ps1` as the automation entrypoint. Never promise that conflicts can be resolved automatically.

## Quick Start

Resolve the installed skill root, then run one of:

```powershell
# Dry-run: report planned actions only.
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-root>\scripts\sync_ai_hot.ps1 -Mode once

# One real sync cycle.
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-root>\scripts\sync_ai_hot.ps1 -Mode once -EnableWrites -Toast

# Scheduled polling, default every 5 minutes unless changed.
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-root>\scripts\sync_ai_hot.ps1 -Mode timer -IntervalSeconds 300 -EnableWrites -Toast

# File-change monitoring with debounce.
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-root>\scripts\sync_ai_hot.ps1 -Mode watch -EnableWrites -Toast
```

Defaults:

```text
RepoUrl = https://github.com/wangdongqin-80/AI-HOT
Mode = once
IntervalSeconds = 300
```

Pass `-RepoPath <path>` when the repository is not the current directory or `.\AI-HOT`.

## Sync Algorithm

1. Locate or clone the configured repository.
2. Refuse to operate unless `origin` contains `wangdongqin-80/AI-HOT`.
3. If local changes exist, stage them and create an auto-sync commit.
4. Run `git fetch --prune`.
5. Compare `HEAD` to upstream.
6. If only behind, run `git pull --no-rebase --ff-only`.
7. If only ahead, run `git push` unless `-NoPush` is set.
8. If both ahead and behind, stop and notify that manual merge or rebase is required.
9. Report success or the exact stop reason.

## Trigger Choice

| User intent | Use |
|---|---|
| "sync now", "pull latest", "push my changes" | `-Mode once` |
| "keep it synced", "check every few minutes" | `-Mode timer -IntervalSeconds 300` |
| "sync whenever files change" | `-Mode watch` |
| "show me what would happen" | omit `-EnableWrites` |

Prefer scheduled polling for pair collaboration. Use file watching only when frequent small auto-sync commits are acceptable.

## Conflict And Safety Rules

- Stop on merge conflicts, branch divergence, non-fast-forward pull failures, authentication failures, detached HEAD, missing Git, and wrong remotes.
- Do not auto-stash, reset, force-push, switch branches, edit conflict markers, or choose a merge strategy on behalf of the user.
- Use console notifications by default. Add `-Toast` when desktop popup notification is helpful.
- After any stop condition, tell the user the failing command or condition and the manual decision needed.
- If `git push` fails, keep the local commit and report that it still needs publishing.

## Commit Message Rule

Use the built-in auto-sync commit message for unattended runs. When running interactively and there is enough context, replace it with a concise message describing the files or behavior changed. Never use vague messages like `update`, `fix`, or `changes`.

## Resource

- `scripts/sync_ai_hot.ps1`: deterministic PowerShell sync runner for one-shot, timer, and watch modes.
