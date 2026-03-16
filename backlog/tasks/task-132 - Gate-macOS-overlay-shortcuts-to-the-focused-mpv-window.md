---
id: TASK-132
title: Gate macOS overlay shortcuts to the focused mpv window
status: Done
assignee:
  - codex
created_date: '2026-03-08 18:24'
updated_date: '2026-03-16 05:13'
labels:
  - bug
  - macos
  - shortcuts
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-shortcut.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/window-trackers/macos-tracker.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/scripts/get-mpv-window-macos.swift
priority: high
ordinal: 52500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the macOS shortcut handling so SubMiner overlay keybinds do not intercept system or other-app shortcuts while SubMiner is in the background. Overlay shortcuts should only be active while the tracked mpv window is present and focused, and should stop grabbing keyboard input when mpv is not the frontmost window.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 On macOS, overlay shortcuts do not trigger while mpv is not the focused/frontmost window.
- [x] #2 On macOS, overlay shortcuts remain available while the tracked mpv window is open and focused.
- [x] #3 Existing non-macOS shortcut behavior is unchanged.
- [x] #4 Automated tests cover the macOS focus-gating behavior and guard against background shortcut interception.
- [x] #5 Any user-facing docs/config notes affected by the behavior change are updated in the same task if needed.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing macOS-focused shortcut lifecycle test that proves overlay shortcuts stay inactive when the tracked mpv window exists but is not frontmost, and activate when that tracked window becomes frontmost.
2. Add a failing tracker/helper test that covers the focused/frontmost signal parsed from the macOS helper output.
3. Extend the macOS helper/tracker contract to surface both geometry and focused/frontmost state for the tracked mpv window.
4. Wire overlay shortcut activation to require both overlay runtime initialization and tracked-mpv focus on macOS, while leaving non-macOS behavior unchanged.
5. Re-run the targeted shortcut/tracker tests, then the broader related shortcut/runtime suite, and update task notes/acceptance criteria based on results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a macOS-specific shortcut activation predicate so global overlay shortcuts now require both overlay runtime readiness and a focused tracked mpv window; non-macOS behavior still keys off runtime readiness only.

Extended the base window tracker with optional focus-state callbacks/getters and wired initializeOverlayRuntime to re-sync overlay shortcuts whenever tracker focus changes.

Updated the macOS helper/tracker contract to return geometry plus frontmost/focused state for the tracked mpv process and added parser coverage for focused and unfocused output.

Verified with `bun x tsc -p tsconfig.json --noEmit`, targeted shortcut/tracker tests, and `bun run test:core:src` (439 passing).

No user-facing config or documentation surface changed, so no docs update was required for this fix.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the macOS background shortcut interception bug by gating SubMiner's global overlay shortcuts on tracked mpv focus instead of overlay-runtime initialization alone. The macOS window helper now reports whether the tracked mpv process is frontmost, the tracker exposes focus change callbacks, and overlay shortcut synchronization re-runs when that focus state flips so `Ctrl+C`/`Ctrl+V` and similar shortcuts are no longer captured while mpv is in the background.

The change keeps existing non-macOS shortcut behavior unchanged. Added regression coverage for the activation decision, tracker focus-change re-sync, and macOS helper output parsing. Verification: `bun x tsc -p tsconfig.json --noEmit`, targeted shortcut/tracker tests, and `bun run test:core:src` (439 passing).
<!-- SECTION:FINAL_SUMMARY:END -->
