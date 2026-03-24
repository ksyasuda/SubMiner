---
id: TASK-201
title: Suppress repeated macOS overlay loading OSD during fullscreen tracker flaps
status: Done
assignee:
  - '@codex'
created_date: '2026-03-19 18:47'
updated_date: '2026-03-23 03:22'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-visibility.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/overlay-visibility-runtime.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/state.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-visibility.test.ts
priority: high
ordinal: 131500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce macOS fullscreen annoyance where the visible overlay briefly loses tracking and re-shows the `Overlay loading...` OSD even though the overlay runtime is already initialized and no new instance is launching. Keep the first startup/loading feedback, but suppress repeat loading notifications caused by subsequent tracker churn during fullscreen enter/leave or focus flaps.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The first macOS visible-overlay load still shows the existing `Overlay loading...` OSD when tracker data is not yet ready.
- [x] #2 Repeated macOS tracker flaps after the overlay has already recovered do not immediately re-show `Overlay loading...` on every loss/recovery cycle.
- [x] #3 Focused regression tests cover the repeated tracker-loss/recovery path and preserve the initial-load notification behavior.
- [x] #4 The change does not alter overlay runtime bootstrap or single-instance behavior; only notification suppression behavior changes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused failing regressions in `src/core/services/overlay-visibility.test.ts` that preserve the first macOS `Overlay loading...` OSD and suppress an immediate second OSD after tracker recovery/loss churn.
2. Extend the overlay-visibility state/runtime plumbing with a small macOS loading-OSD suppression state so tracker flap retries can be rate-limited without touching overlay bootstrap or single-instance logic.
3. Reset the suppression when the user explicitly hides the visible overlay so intentional hide/show retries can still surface first-load feedback.
4. Run focused verification for the touched overlay visibility/runtime tests and update the task with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added optional loading-OSD suppression hooks to `src/core/services/overlay-visibility.ts` so macOS can rate-limit repeated `Overlay loading...` notifications without changing overlay bootstrap behavior.

Implemented service-local suppression state in `src/main/overlay-visibility-runtime.ts` with a 30s cooldown and explicit reset when the visible overlay is manually hidden, so fullscreen tracker flaps stay quiet but intentional hide/show retries can still show loading feedback.

Added focused regressions in `src/core/services/overlay-visibility.test.ts` for `loss -> recover -> immediate loss` suppression and for manual hide resetting suppression.

Verification: `bun test src/core/services/overlay-visibility.test.ts`; `bun test src/main/runtime/overlay-visibility-runtime-main-deps.test.ts src/main/runtime/overlay-visibility-runtime.test.ts`; `bun run typecheck`; `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane runtime-compat src/core/services/overlay-visibility.ts src/main/overlay-visibility-runtime.ts src/core/services/overlay-visibility.test.ts` -> passed. Real-runtime lane skipped: change is notification suppression logic and cheap/runtime-compat coverage was sufficient for this scoped behavior change; no live mpv/macOS fullscreen session was run in this turn.

Docs update required: no. Changelog fragment required: yes; added `changes/2026-03-19-overlay-loading-osd-fullscreen-flaps.md`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Reduced repeated macOS `Overlay loading...` popups caused by fullscreen tracker flap churn without touching overlay bootstrap or single-instance behavior. `src/core/services/overlay-visibility.ts` now accepts optional suppression hooks around the loading OSD path, and `src/main/overlay-visibility-runtime.ts` uses service-local state to rate-limit that OSD for 30 seconds while resetting the suppression when the visible overlay is explicitly hidden. Added focused regressions in `src/core/services/overlay-visibility.test.ts` to preserve the first-load notification, suppress immediate repeat notifications after tracker recovery/loss churn, and keep manual hide/show retries able to surface the loading OSD again. Added changelog fragment `changes/2026-03-19-overlay-loading-osd-fullscreen-flaps.md`. Verification passed with targeted overlay tests, typecheck, and the `runtime-compat` verifier lane; live macOS/mpv fullscreen runtime validation was not run in this turn.
<!-- SECTION:FINAL_SUMMARY:END -->
