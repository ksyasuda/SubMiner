---
id: TASK-323
title: Fix macOS overlay hiding while mpv remains active
status: Done
assignee:
  - '@codex'
created_date: '2026-05-03 07:41'
updated_date: '2026-05-03 07:48'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - src/core/services/overlay-visibility.ts
  - src/window-trackers/macos-tracker.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
macOS visible overlay can hide/reload during normal playback even while mpv, or the overlay over mpv, remains the active viewing surface. The fix should preserve overlay visibility and subtitle continuity during transient macOS focus/tracker flaps, while still hiding the overlay when the tracked mpv window is genuinely unavailable or another app is brought forward.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When the macOS tracker has recent valid mpv geometry, transient focus/helper misses do not hide the visible overlay or force a reload.
- [x] #2 The overlay still hides when the tracked mpv window is genuinely lost beyond the existing tracking grace behavior.
- [x] #3 A regression test covers the macOS active-playback case where mpv/overlay focus is preserved despite a transient non-tracking state.
- [x] #4 Relevant docs or task notes are updated if behavior or verification guidance changes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression in `src/core/services/overlay-visibility.test.ts`: on macOS, after the overlay is visible/tracked, a transient tracker state with `isTracking() === false` but non-null `getGeometry()` keeps the overlay visible, updates bounds, and does not call `hide()` or loading OSD.
2. Implement the minimal macOS preserve path in `src/core/services/overlay-visibility.ts`, mirroring the existing Windows transient non-minimized branch but without Windows z-order binding.
3. Preserve existing startup/lost-window behavior: `windowTracker: null` and `isTracking() === false` with `getGeometry() === null` still hide and show the first loading OSD.
4. Run focused tests for `src/core/services/overlay-visibility.test.ts`; then typecheck or the repo runtime verification lane if the focused patch passes.
5. Update TASK-323 notes/acceptance criteria with verification results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a macOS overlay visibility regression for transient tracker loss with retained geometry. The test failed first because the old path marked tracker-not-ready and hid the overlay. Implemented a scoped preserve path in `src/core/services/overlay-visibility.ts`: macOS now keeps the visible overlay alive only when the tracker still has retained geometry; true loss with null geometry still hides and emits the existing loading OSD behavior. Added changelog fragment `changes/323-macos-overlay-tracker-flaps.md`.

Verification: `bun test src/core/services/overlay-visibility.test.ts` passed after the fix; `bun test src/window-trackers/macos-tracker.test.ts src/core/services/overlay-visibility.test.ts` passed; `bun run typecheck` passed; `bun run test:env` passed; isolated `bun test src/core/services/subsync.test.ts` passed; `bun run build` passed; `bun run test:smoke:dist` passed; `bun run changelog:lint` passed. `bun run test:fast` failed twice in an unrelated broad-suite interaction where `src/renderer/handlers/keyboard.ts` tried to use missing `window.electronAPI` while `src/core/services/subsync.test.ts` was running, followed by Bun node:test nested-test cascade errors.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the macOS visible-overlay hide/reload path during normal playback by preserving the overlay when the tracker briefly reports non-tracking but still has retained mpv geometry. The overlay visibility service now treats that macOS state like a transient tracker flap: it keeps bounds/layer/order refreshed and leaves the overlay click-through instead of hiding or showing the loading OSD. True macOS loss remains unchanged: no tracker or null geometry still hides the overlay and uses the existing loading behavior.

Added regression coverage in `src/core/services/overlay-visibility.test.ts` for the active-playback case and added changelog fragment `changes/323-macos-overlay-tracker-flaps.md`.

Verification passed: focused overlay tests, macOS tracker + overlay tests, typecheck, `test:env`, isolated `subsync.test.ts`, build, dist smoke, and changelog lint. Full `test:fast` remains blocked by an unrelated broad-suite interaction where renderer keyboard state fires without `window.electronAPI` during `subsync.test.ts`, then Bun reports node:test cascade errors.
<!-- SECTION:FINAL_SUMMARY:END -->
