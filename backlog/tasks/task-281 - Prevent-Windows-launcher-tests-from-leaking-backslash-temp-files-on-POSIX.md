---
id: TASK-281
title: Prevent Windows launcher tests from leaking backslash temp files on POSIX
status: Done
assignee:
  - codex
created_date: '2026-04-05 21:13'
updated_date: '2026-04-05 21:20'
labels:
  - tests
  - launcher
  - bug
dependencies: []
documentation:
  - /home/sudacode/github/SubMiner2/AGENTS.md
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Windows-specific launcher tests in launcher/mpv.test.ts currently create real filesystem entries using path.win32.join(...) with a POSIX mkdtemp base. On Linux/macOS this produces literal backslash-named paths like \\tmp\\subminer-test-win-dir-* inside the repo/worktree, and the existing cleanup only removes the POSIX /tmp base directory. Fix the tests so they still cover Windows path resolution behavior without leaking stray files into the working tree.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Running the Windows findAppBinary tests on a POSIX host does not create new untracked \\tmp\\subminer-test-win-* files in the repository root.
- [x] #2 The Windows launcher tests still validate PATH and install-directory resolution behavior.
- [x] #3 Relevant launcher tests pass after the change.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test in launcher/mpv.test.ts that exercises the Windows findAppBinary cases on a POSIX host and asserts they do not leave new backslash-named temp artifacts in the repository root.
2. Refactor the Windows launcher tests to avoid creating real filesystem paths from path.win32.join(...) on POSIX; keep Windows path assertions via stubs and only create real files with native POSIX paths where needed.
3. Run the targeted launcher tests and confirm no new \\tmp\\subminer-test-win-* artifacts appear in git status.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Investigation: reproduced the leak locally. The source is launcher/mpv.test.ts Windows findAppBinary tests that combine a POSIX mkdtemp base with path.win32.join(...), creating literal backslash-named entries like \\tmp\\subminer-test-win-dir-* in the repo root. Existing cleanup only removes the POSIX /tmp base directory.

User approved implementation plan on 2026-04-05.

Implemented in launcher/mpv.test.ts by replacing the leaky Windows PATH/install-directory helpers with pure fs stubs (access/exists/stat) and fixed Windows path strings instead of creating real path.win32 filesystem entries on POSIX. Added a regression test that snapshots repo-root \\tmp\\subminer-test-win-* artifacts before/after running the Windows cases and asserts no new entries are created.

Verification: `bun test launcher/mpv.test.ts --test-name-pattern 'findAppBinary Windows cases do not leak backslash temp artifacts on POSIX|findAppBinary resolves SubMiner.exe on PATH on Windows|findAppBinary resolves a Windows install directory to SubMiner.exe'` passed (3/3). `bun test launcher/mpv.test.ts` still has one unrelated pre-existing sandbox failure in `launchAppCommandDetached handles child process spawn errors` because the test opens `~/.config/SubMiner/logs/app-2026-04-05.log` and hits `EROFS` in this environment.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Reworked the Windows `findAppBinary` tests in `launcher/mpv.test.ts` so they no longer create real backslash-named temp files on POSIX hosts. The PATH and install-directory cases now use synthetic Windows path strings plus `fs.accessSync` / `fs.existsSync` / `fs.statSync` stubs to exercise the same resolver behavior without writing `\\tmp\\subminer-test-win-*` entries into the repository root.

Added a POSIX regression test that snapshots existing repo-root `\\tmp\\subminer-test-win-*` artifacts, runs the Windows path-resolution cases, and asserts the artifact set is unchanged. This catches future regressions where a Windows-path test accidentally writes literal backslash paths on Linux/macOS.

Tests run:
- `bun test launcher/mpv.test.ts --test-name-pattern 'findAppBinary Windows cases do not leak backslash temp artifacts on POSIX|findAppBinary resolves SubMiner.exe on PATH on Windows|findAppBinary resolves a Windows install directory to SubMiner.exe'`
- `bun test launcher/mpv.test.ts` (all relevant `findAppBinary` tests passed; one unrelated existing sandbox failure remains in `launchAppCommandDetached handles child process spawn errors` due `EROFS` opening `~/.config/SubMiner/logs/...`)
<!-- SECTION:FINAL_SUMMARY:END -->
