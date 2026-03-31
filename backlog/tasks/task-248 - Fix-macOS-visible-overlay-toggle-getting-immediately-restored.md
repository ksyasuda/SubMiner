---
id: TASK-248
title: Fix macOS visible overlay toggle getting immediately restored
status: Done
assignee: []
created_date: '2026-03-29 10:03'
updated_date: '2026-03-31 19:37'
labels: []
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/plugin/subminer/process.lua
  - /Users/sudacode/projects/japanese/SubMiner/plugin/subminer/ui.lua
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/cli-command.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/overlay-visibility-runtime.ts
ordinal: 165500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix the visible overlay toggle path on macOS so the user can reliably hide the overlay after it has been shown. The current behavior can ignore the toggle or hide the overlay briefly before it is restored immediately.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Pressing the visible-overlay toggle hides the overlay when it is currently shown on macOS.
- [x] #2 A manual hide is not immediately undone by startup or readiness flows.
- [x] #3 The mpv/plugin toggle path matches the intended visible-overlay toggle behavior.
- [x] #4 Regression tests cover the failing toggle path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the toggle/re-show logic from code paths around mpv plugin control commands and auto-play readiness.
2. Add regression coverage for manual toggle-off staying hidden through readiness completion.
3. Patch the plugin/control path so manual visible-overlay toggles are not undone by readiness auto-show.
4. Run targeted tests, then the relevant verification lane.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: the mpv plugin readiness callback (`subminer-autoplay-ready`) could re-issue `--show-visible-overlay` after a manual toggle/hide. Initial fix only suppressed the next readiness restore, but repeated readiness callbacks in the same media session could still re-show the overlay. The plugin toggle path also still used legacy `--toggle` instead of the explicit visible-overlay command.

Implemented a session-scoped suppression flag in the Lua plugin so a manual hide/toggle during the pause-until-ready window blocks readiness auto-show for the rest of the current auto-start session, then resets on the next auto-start session.

Added Lua regression coverage for both behaviors: manual toggle-off stays hidden through readiness completion, repeated readiness callbacks in the same session stay suppressed, and `subminer-toggle` emits `--toggle-visible-overlay` rather than legacy `--toggle`.

Follow-up investigation found a second issue in `src/core/services/cli-command.ts`: pure visible-overlay toggle commands still ran the MPV connect/start path (`connectMpvClient`) because `--toggle` and `--toggle-visible-overlay` were classified as start-like commands. That side effect could retrigger startup visibility work even after the plugin-side fix.

Updated CLI command handling so only `--start` reconnects MPV. Pure toggle/show/hide overlay commands still initialize overlay runtime when needed, but they no longer restart/reconnect the MPV control path.

Renderer/modal follow-ups: restored focused-overlay mpv y-chord proxy in `src/renderer/handlers/keyboard.ts`, added a modal-close guard in `src/main/overlay-runtime.ts` so modal teardown does not re-show a manually hidden overlay, and added a duplicate-toggle debounce in `src/main/runtime/overlay-visibility-actions.ts` to ignore near-simultaneous toggle requests inside the main process.

2026-03-29: added regression for repeated subminer-autoplay-ready signals after manual y-t hide. Root cause: Lua plugin suppression only blocked the first ready-time restore, so later ready callbacks in the same media session could re-show the visible overlay. Updated plugin suppression to remain active for the full current auto-start session and reset on the next auto-start trigger.

2026-03-29: live mpv log showed repeated `subminer-autoplay-ready` script messages from Electron during paused startup, each triggering plugin `--show-visible-overlay` and immediate re-show. Fixed `src/main/runtime/autoplay-ready-gate.ts` so plugin readiness is signaled once per media while paused retry loops only re-issue `pause=false` instead of re-signaling readiness.

2026-03-29: Added window-level guard for stray visible-overlay re-show on macOS. `src/core/services/overlay-window.ts` now immediately re-hides the visible overlay window on `show` if overlay state is false, covering native/Electron re-show paths that bypass normal visibility actions. Regression: `src/core/services/overlay-window.test.ts`. Verified with full gate and rebuilt unsigned mac bundle.

2026-03-29: added a blur-path guard for the visible overlay window. `src/core/services/overlay-window.ts` now skips topmost restacking when a visible-overlay blur fires after overlay state already flipped off, covering a macOS hide-in-flight path that could immediately reassert the window. Regression coverage added in `src/core/services/overlay-window.test.ts`; verified with targeted overlay tests, full gate, and rebuilt unsigned mac bundle.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Confirmed with user that macOS `y-t` now works. Cleaned the patch set down to the remaining justified fixes: explicit visible-overlay plugin toggle/suppression, pure-toggle CLI no longer reconnects MPV, autoplay-ready signaling only fires once per media, and the final visible-overlay blur guard that stops macOS restacking after a manual hide. Full gate passed again before commit `c939c580` (`fix: stabilize macOS visible overlay toggle`).
<!-- SECTION:FINAL_SUMMARY:END -->
