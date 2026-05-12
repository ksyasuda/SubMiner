---
id: TASK-349
title: Fix macOS overlay window ordering behind foreground apps
status: Done
assignee: []
created_date: '2026-05-12 08:50'
updated_date: '2026-05-12 08:58'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - src/core/services/overlay-visibility.ts
  - src/window-trackers
  - TASK-344
modified_files:
  - src/core/services/overlay-visibility.ts
  - src/core/services/overlay-visibility.test.ts
  - src/core/services/overlay-window-input.ts
  - src/core/services/overlay-window.test.ts
  - changes/349-macos-overlay-z-order.md
priority: high
ordinal: 183500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
macOS overlay should stay visually above mpv, but remain grouped with mpv in normal desktop stacking. When another app/window is brought in front of mpv, that window must also appear in front of the overlay, matching Windows behavior. This follows the earlier active-mpv fix that stopped the overlay from hiding while mpv remained foremost.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When mpv is the foreground playback window on macOS, the overlay remains visible above mpv.
- [x] #2 When another application or window is brought in front of mpv on macOS, that foreground window appears above both mpv and the overlay.
- [x] #3 Restoring mpv to the foreground restores the overlay above mpv without requiring a restart.
- [x] #4 Regression coverage documents the macOS stacking relationship and does not regress the prior active-mpv overlay preservation behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression coverage for macOS mpv focus loss: the overlay must release its topmost level, remain visible/click-through, and stop enforcing layer order while mpv is behind another window.
2. Add focused blur-handler coverage so the macOS visible overlay does not restack itself when it loses focus.
3. Update overlay visibility and blur handling to use tracker focus as the macOS stacking boundary: focused mpv raises overlay; unfocused mpv releases topmost and skips restack.
4. Run focused overlay tests, formatting, typecheck, changelog lint, env/build/smoke checks; document any blocked broad gate separately.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented macOS stacking boundary using tracker focus. When tracked mpv is unfocused and the overlay itself is not focused, the visible overlay now releases Electron always-on-top, remains visible/click-through, and skips layer-order enforcement. Visible overlay blur restacking is also skipped on macOS, matching the Windows no-restack path for focus loss. `test:fast` remains blocked by existing cross-file pollution: `keyboard.test.ts` leaves `window.electronAPI` undefined for a later `subsync.test.ts`, causing Bun nested `node:test` errors in subsequent files.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Updated macOS overlay visibility so tracked mpv focus controls stacking: focused mpv keeps the overlay raised; unfocused mpv releases topmost while keeping the overlay visible and click-through.
- Stopped macOS visible overlay blur handling from immediately restacking the overlay above unrelated foreground windows.
- Added regression tests for macOS mpv focus loss and macOS blur restacking behavior.
- Added a changelog fragment for the user-visible overlay z-order fix.

Verification:
- Passed: `bun test src/core/services/overlay-visibility.test.ts src/core/services/overlay-window.test.ts`
- Passed: `bunx prettier --check src/core/services/overlay-visibility.ts src/core/services/overlay-visibility.test.ts src/core/services/overlay-window-input.ts src/core/services/overlay-window.test.ts changes/349-macos-overlay-z-order.md`
- Passed: `bun run typecheck`
- Passed: `bun run changelog:lint`
- Passed: `bun run test:env`
- Passed: `bun run build`
- Passed: `bun run test:smoke:dist`
- Blocked: `bun run test:fast` by existing keyboard/subsync cross-file global pollution; focused and environment tests pass.
<!-- SECTION:FINAL_SUMMARY:END -->
