---
id: TASK-286.1
title: 'Assess and address PR #49 subsequent CodeRabbit review round'
status: Done
assignee: []
created_date: '2026-04-11 23:14'
updated_date: '2026-04-11 23:16'
labels:
  - bug
  - code-review
  - windows
  - release
dependencies: []
references:
  - .github/workflows/prerelease.yml
  - src/window-trackers/mpv-socket-match.ts
  - src/window-trackers/win32.ts
  - src/core/services/overlay-shortcut.ts
parent_task_id: TASK-286
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the next unresolved CodeRabbit review threads on PR #49 after commit 9ce5de2f and resolve the still-valid follow-up issues without reopening already-assessed stale comments.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All still-actionable CodeRabbit comments in the latest PR #49 round are fixed or explicitly shown stale with evidence.
- [x] #2 Regression coverage is added or updated for any behavior-sensitive fix in workflow or Windows socket matching.
- [x] #3 Relevant verification passes for the touched workflow, tracker, and shared matcher changes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Verify the five unresolved CodeRabbit threads against current branch state and separate still-valid bugs from stale comments.
2. Add or extend the narrowest failing tests for exact socket-path matching and prerelease workflow invariants before changing production code.
3. Implement minimal fixes in the prerelease workflow and Windows socket matching/cache path, leaving stale comments documented with evidence instead of forcing no-op changes.
4. Run targeted verification, then record the fixed-vs-stale assessment and close the subtask.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Assessed five unresolved PR #49 threads after 9ce5de2f. Fixed prerelease workflow cache keys to include `runner.arch`, changed prerelease publishing to validate artifacts before release creation/edit and only undraft after uploads complete, tightened Windows socket matching to require exact argument boundaries, and stopped memoizing null command-line lookup misses in the Win32 cache path.

Stale assessment: the `src/core/services/overlay-shortcut.ts` thread is still obsolete. Current code at `registerOverlayShortcuts()` returns `hasConfiguredOverlayShortcuts(shortcuts)`, not `false`, and the overlay-local handling remains intentionally driven by local fallback dispatch rather than global registration in this runtime path.

Verification: `bun test src/prerelease-workflow.test.ts src/window-trackers/mpv-socket-match.test.ts`, `bun test src/window-trackers/windows-tracker.test.ts src/prerelease-workflow.test.ts src/window-trackers/mpv-socket-match.test.ts`, `bun run typecheck`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Handled the next CodeRabbit round on PR #49 by fixing the still-valid prerelease workflow and Windows socket-matching issues while documenting the stale overlay-shortcut comment instead of forcing a no-op code change. The prerelease workflow now scopes all dependency caches by `runner.arch`, validates the final artifact set before touching the GitHub release, creates/edits the prerelease as a draft during uploads, and only flips `--draft=false` after all assets succeed. On Windows, socket matching now requires an exact `--input-ipc-server` argument boundary so `subminer-1` no longer matches `subminer-10`, and transient PowerShell/CIM misses no longer get cached forever as null command lines.

Regression coverage was added for the workflow invariants and exact socket matching. Verification passed with targeted prerelease workflow tests, Windows tracker tests, socket-matcher tests, and `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
