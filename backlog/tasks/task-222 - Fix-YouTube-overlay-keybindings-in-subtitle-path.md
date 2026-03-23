---
id: TASK-222
title: Fix YouTube overlay keybindings in subtitle path
status: Done
assignee:
  - codex
created_date: '2026-03-23 08:32'
updated_date: '2026-03-23 08:38'
labels:
  - bug
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Users watching video through the YouTube subtitle path cannot use some overlay keyboard controls such as quit and pause/play. Restore expected overlay keybinding behavior for that playback path without regressing other overlay input handling.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Overlay quit and pause/play keybindings work while using the YouTube subtitle path.
- [x] #2 Existing overlay keybinding behavior for non-YouTube playback remains unchanged.
- [x] #3 Regression coverage exercises the YouTube subtitle path keyboard handling.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test around YouTube track-picker close to verify it requests main-process main-window focus restoration before returning overlay focus locally.
2. Update the YouTube track-picker close flow to call `window.electronAPI.focusMainWindow()` alongside the existing `window.focus()` and `overlay.focus()` restoration.
3. Run targeted tests for the picker/keyboard paths to verify YouTube playback regains overlay keybindings without regressing existing overlay behavior.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Investigated overlay input path. Renderer already maps Space/KeyQ to mpv commands, but YouTube track-picker close only restores DOM focus (`window.focus` + `overlay.focus`) and does not invoke main-process window focus recovery, unlike the keyboard-mode focus reclaim path. Suspected root cause: overlay BrowserWindow focus is not restored after the YouTube picker closes, so playback keybindings stop reaching renderer keydown handlers.

User approved implementation plan on 2026-03-23. Proceeding with TDD: add failing regression first, then minimal fix, then targeted verification.

Implemented fix in the YouTube track-picker close path: request main-process `focusMainWindow()` before restoring renderer window/overlay focus so overlay keydown handlers regain input after YouTube subtitle selection.

Verification: `bun test src/renderer/modals/youtube-track-picker.test.ts` and `bun test src/renderer/handlers/keyboard.test.ts` both pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored overlay keyboard focus after closing the YouTube subtitle picker by invoking the main-process `focusMainWindow()` recovery path before local window/overlay focus restoration. Added regression coverage to the YouTube picker modal test and verified existing keyboard handler coverage for YouTube picker passthrough keys (`Space`, `KeyQ`) remains green.
<!-- SECTION:FINAL_SUMMARY:END -->
