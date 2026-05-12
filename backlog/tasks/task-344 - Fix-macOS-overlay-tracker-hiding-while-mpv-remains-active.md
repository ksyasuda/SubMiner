---
id: TASK-344
title: Fix macOS overlay tracker hiding while mpv remains active
status: Done
assignee:
  - codex
created_date: '2026-05-11 08:27'
updated_date: '2026-05-11 08:41'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - src/main/runtime/overlay-visibility-runtime.ts
  - src/window-trackers
modified_files:
  - src/core/services/overlay-visibility.ts
  - src/core/services/overlay-visibility.test.ts
  - changes/344-macos-overlay-active-mpv.md
priority: high
ordinal: 182500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
macOS playback overlay should match Windows behavior: the tracker may only hide or alter overlay layering when mpv is no longer the active playback window. When mpv remains topmost or fullscreen, the visible overlay must stay present and interactive unless the user manually hides it or minimizes mpv.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 With mpv active on macOS, window-tracker updates do not hide the visible overlay or make it click-through.
- [x] #2 With mpv fullscreen on macOS, tracker geometry/layering refreshes preserve overlay interactivity.
- [x] #3 Overlay visibility still changes when mpv is no longer active, and manual hide/minimize behavior remains intact.
- [x] #4 A regression test covers the macOS active-mpv path that previously produced an overlay loading OSD and non-interactive overlay.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression in `src/core/services/overlay-visibility.test.ts` for macOS with mpv tracked/focused and retained geometry: visibility refresh must not hide the overlay, must not emit loading OSD, and must leave the overlay interactive (`setIgnoreMouseEvents(false)`) unless forced passthrough is active.
2. Update `src/core/services/overlay-visibility.ts` so macOS tracker refreshes preserve the visible overlay while mpv is active/fullscreen, and only hide/re-layer for explicit manual hide, minimized/untracked target, or non-active mpv cases.
3. Run the focused overlay visibility tests, then a broader fast gate if the focused fix is green.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented in the shared overlay visibility service. macOS now keeps tracked/focused mpv overlays interactive instead of defaulting to mouse passthrough, and preserves an already visible active-mpv overlay during temporary tracker-not-ready refreshes without showing the loading OSD. Forced passthrough, modal hide, manual hide, Windows minimized handling, and initial macOS tracker-not-ready startup behavior remain covered by tests.

Verification: `bun test src/core/services/overlay-visibility.test.ts` passed; affected overlay/mouse/runtime group passed; `bun run typecheck`, `bun run changelog:lint`, `bun run build`, `bun run test:env`, and `bun run test:smoke:dist` passed. `bun run test:fast` is blocked by existing cross-file test pollution: `src/core/services/subsync.test.ts` passes alone, but fails when run after `src/renderer/handlers/keyboard.test.ts` because `window.electronAPI` is undefined in a lingering keyboard handler; Bun then reports nested node:test errors for later files. `bun run format:check:src` is blocked by pre-existing formatting drift in `src/core/services/stats-window.ts`; touched files pass direct Prettier check.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Updated macOS overlay visibility logic so a tracked/focused mpv window keeps the visible overlay interactive instead of click-through.
- Preserved an already visible active-mpv overlay during temporary macOS tracker-not-ready refreshes, avoiding the loading OSD/hide path for that active playback case.
- Added regression coverage for active mpv tracker refreshes and transient tracker-not-ready refreshes, plus updated old macOS expectations to the new active-mpv contract.
- Added a changelog fragment for the user-visible overlay fix.

Verification:
- Passed: `bun test src/core/services/overlay-visibility.test.ts`
- Passed: `bun test src/core/services/overlay-visibility.test.ts src/core/services/overlay-runtime-init.test.ts src/core/services/overlay-shortcut-handler.test.ts src/renderer/handlers/mouse.test.ts`
- Passed: `bun run typecheck`
- Passed: `bun run changelog:lint`
- Passed: `bun run build`
- Passed: `bun run test:env`
- Passed: `bun run test:smoke:dist`
- Blocked: `bun run test:fast` by existing keyboard/subsync cross-file global pollution; isolated `bun test src/core/services/subsync.test.ts` passes.
- Blocked: `bun run format:check:src` by pre-existing formatting drift in `src/core/services/stats-window.ts`; touched files pass direct Prettier check.
<!-- SECTION:FINAL_SUMMARY:END -->
