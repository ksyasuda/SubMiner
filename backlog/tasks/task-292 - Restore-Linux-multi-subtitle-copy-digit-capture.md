---
id: TASK-292
title: Restore Linux multi-subtitle copy digit capture
status: Done
assignee:
  - '@codex'
created_date: '2026-04-25 21:31'
updated_date: '2026-04-25 21:36'
labels:
  - bug
  - linux
  - shortcuts
  - clipboard
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
On Linux, the copy-subtitle-multiple shortcut opens the numeric prompt but the follow-up digit is not captured, so the flow times out. User confirmed `wl-copy` itself is installed and working, so investigate the shortcut/digit capture path and restore multi-line subtitle copy without regressing existing session action behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Linux copy-subtitle-multiple shortcut accepts a follow-up digit and copies that number of recent subtitle lines instead of timing out.
- [x] #2 The fix avoids depending on Linux Electron global shortcut digit registration for the follow-up numeric selection when a renderer-visible session can handle it.
- [x] #3 Regression tests cover the Linux multi-copy shortcut/digit flow and existing non-Linux/global shortcut behavior remains intact.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression coverage for Linux copy-subtitle-multiple local shortcut fallback starting renderer/session numeric selection instead of main-process digit globalShortcut capture.
2. Patch the overlay shortcut fallback/runtime path so Linux visible-overlay multi-copy and mine-sentence-multiple can dispatch session-action numeric selection when renderer handling is available, while preserving main-process numeric sessions for CLI/non-renderer paths.
3. Run targeted tests for shortcut fallback, overlay runtime, and renderer keyboard numeric selection; then run typecheck or a wider focused gate if needed.
4. Update task acceptance criteria/final notes after verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented the approved path by keeping multi-step numeric overlay shortcuts out of the main-process local fallback. The visible overlay now receives the original keydown and uses the existing renderer/session-action numeric selection flow for follow-up digits, avoiding Linux Electron globalShortcut digit capture for multi-copy and mine-sentence-multiple. Verification: targeted shortcut/renderer tests and changelog lint pass. `bun run typecheck` is currently blocked by unrelated existing errors in CLI/AniList dictionary-candidate work and `src/main/dependencies.ts` manual-selection API shape.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored Linux multi-line subtitle copy by preventing main-process overlay shortcut fallback from consuming multi-step numeric shortcuts (`copySubtitleMultiple` and `mineSentenceMultiple`). Those shortcuts now fall through to the visible overlay renderer, where the existing session binding flow prompts for a digit and dispatches the counted session action locally instead of relying on Electron globalShortcut digit registration. Added regression coverage for the fallback behavior and renderer follow-up digit dispatch, plus a changelog fragment.

Verification: `bun test src/core/services/overlay-shortcut-handler.test.ts src/renderer/handlers/keyboard.test.ts`; `bun run changelog:lint`. Full `bun run typecheck` was attempted but is blocked by unrelated current worktree errors in CLI/AniList dictionary-candidate tests/types and `src/main/dependencies.ts`.
<!-- SECTION:FINAL_SUMMARY:END -->
