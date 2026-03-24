---
id: TASK-220
title: Restore YouTube overlay mpv keybindings after picker routing
status: Done
assignee:
  - codex
created_date: '2026-03-22 00:00'
updated_date: '2026-03-22 23:49'
labels:
  - bug
  - overlay
  - youtube
  - keyboard
dependencies: []
references:
  - src/renderer/handlers/keyboard.ts
  - src/renderer/modals/youtube-track-picker.ts
  - src/renderer/handlers/keyboard.test.ts
  - src/renderer/modals/youtube-track-picker.test.ts
documentation:
  - docs/workflow/verification.md
priority: high
ordinal: 118800
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Regression: after adding the YouTube subtitle picker modal path, visible-overlay keydown handling can stop before reaching the shared mpv keybinding dispatch path. Result: default overlay mpv bindings like `Space` pause/play and `q` quit stop working while the overlay owns focus during YouTube playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Unhandled keys while the YouTube track picker state is active still fall through to the shared overlay mpv keybinding dispatcher.
- [x] #2 The YouTube picker continues to consume `Enter` and `Escape` for its own actions.
- [x] #3 Renderer regression tests cover both the picker modal key contract and the shared keyboard dispatch fallback.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing renderer keyboard regression test covering YouTube picker state plus shared mpv keybinding fallback.
2. Update the global keyboard handler to return early only when the YouTube picker actually handles the key event.
3. Update the picker modal handler to return false for unhandled keys while preserving `Enter`/`Escape`.
4. Run the cheap renderer verification lane and record results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Fixed the regression by making the global renderer keyboard handler stop early for the YouTube picker only when the picker actually consumes the key. The picker modal now returns `false` for unrelated keys, so shared overlay mpv bindings like `Space` and `KeyQ` still dispatch while the visible overlay has focus.

Added regression coverage in the keyboard handler suite for mpv keybinding fallback during YouTube picker state, plus a picker-modal contract test that keeps `Escape` handled but leaves unrelated keys unclaimed.

Verification:
- `bun test src/renderer/handlers/keyboard.test.ts src/renderer/modals/youtube-track-picker.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/renderer/handlers/keyboard.ts src/renderer/handlers/keyboard.test.ts src/renderer/modals/youtube-track-picker.ts src/renderer/modals/youtube-track-picker.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/renderer/handlers/keyboard.ts src/renderer/handlers/keyboard.test.ts src/renderer/modals/youtube-track-picker.ts src/renderer/modals/youtube-track-picker.test.ts`
- verifier artifact: `.tmp/skill-verification/subminer-verify-20260322-234831-b2m6nJ`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored YouTube-session overlay mpv keybindings by removing an unconditional early return added to the renderer keyboard path for the YouTube subtitle picker modal. Unhandled keys now fall through to the shared mpv keybinding dispatcher, while handled picker keys (`Enter`, `Escape`) still stay local to the picker. Added renderer regression tests for both the keyboard fallback path and the picker modal key-consumption contract.
<!-- SECTION:FINAL_SUMMARY:END -->
