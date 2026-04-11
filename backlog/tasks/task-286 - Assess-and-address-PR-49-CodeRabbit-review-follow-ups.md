---
id: TASK-286
title: 'Assess and address PR #49 CodeRabbit review follow-ups'
status: Done
assignee:
  - codex
created_date: '2026-04-11 18:55'
updated_date: '2026-04-11 22:40'
labels:
  - bug
  - code-review
  - windows
  - overlay
dependencies: []
references:
  - src/main/runtime/config-hot-reload-handlers.ts
  - src/renderer/handlers/keyboard.ts
  - src/renderer/handlers/mouse.ts
  - vendor/subminer-yomitan
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the current PR #49 review round and resolve the actionable CodeRabbit findings on the Windows update branch.

Focus areas include the renderer mouse interaction fix, config hot-reload keyboard state, and any other review items that still apply after verifying the current branch state.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All actionable CodeRabbit comments on PR #49 are either fixed or shown to be obsolete with evidence.
- [x] #2 Regression tests are added or updated for any behavior change that could regress.
- [x] #3 The branch passes the repo's relevant verification checks for the touched areas.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Pull the current unresolved CodeRabbit review threads for PR #49 and cluster them into still-actionable fixes versus obsolete/nit-only items.
2. For each still-actionable behavior bug, add or extend the narrowest failing test first in the touched suite before changing production code.
3. Implement the minimal fixes across the affected runtime, renderer, plugin, IPC, and Windows tracker files, keeping each change traceable to the review thread.
4. Run targeted verification for the touched areas, update task notes with assessment results, and capture which review comments were fixed versus assessed as obsolete or deferred nitpicks.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Assessed PR #49 CodeRabbit threads. Fixed the real regressions in first-run CLI gating, IPC session-action validation, renderer controller-modal lifecycle notifications, async subtitle-sidebar toggle guarding, plugin config-dir resolution priority, prerelease artifact upload failure handling, immersion tracker lazy startup, win32 z-order error handling, and Windows socket-aware mpv matching.

Review assessment: the overlay-shortcut lifecycle comment is obsolete for the current architecture because overlay-local shortcuts are intentionally handled through the local fallback path and the runtime only tracks configured-state/cleanup. Refactor-only nit comments for splitting `scripts/build-changelog.ts` and `src/core/services/session-bindings.ts` were left as follow-up quality work, not behavior bugs in this PR round.

Verification: `bun test src/main/runtime/first-run-setup-service.test.ts src/core/services/session-bindings.test.ts src/core/services/app-ready.test.ts src/core/services/ipc.test.ts src/renderer/handlers/keyboard.test.ts src/main/overlay-runtime.test.ts src/window-trackers/mpv-socket-match.test.ts`, `bun test src/window-trackers/windows-tracker.test.ts`, `bun run typecheck`, `lua scripts/test-plugin-lua-compat.lua`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the current CodeRabbit round on PR #49 and addressed the still-valid behavior issues rather than blanket-applying every bot suggestion. The branch now treats the new session/stats CLI flags as explicit startup commands during first-run setup, validates the new session actions through IPC, points session-binding command diagnostics at the correct config field, keeps immersion tracker startup lazy until later runtime triggers, and notifies overlay modal lifecycle state when controller-select/debug are opened from local keyboard bindings. I also switched the subtitle-sidebar IPC callback to the async guarded path so promise rejections feed renderer recovery instead of being dropped.

On the Windows/plugin side, the mpv plugin now prefers config-file matches before falling back to an existing config directory, prerelease workflow uploads fail if expected Linux/macOS artifacts are missing, the Win32 z-order bind path now validates the `GetWindowLongW` call for the window above mpv, and the Windows tracker now passes the target socket path into native polling and filters mpv instances by command line so multiple sockets can be distinguished on Windows. Added/updated regression coverage for first-run gating, IPC validation, session-binding diagnostics, controller modal lifecycle notifications, modal ready-listener dispatch, and socket-path matching. Verification run: `bun run typecheck`, the targeted Bun test suites for the touched areas, `bun test src/window-trackers/windows-tracker.test.ts`, and `lua scripts/test-plugin-lua-compat.lua`.
<!-- SECTION:FINAL_SUMMARY:END -->
