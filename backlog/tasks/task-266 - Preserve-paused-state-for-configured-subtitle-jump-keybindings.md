---
id: TASK-266
title: Preserve paused state for configured subtitle-jump keybindings
status: Done
assignee: []
created_date: '2026-04-01 03:19'
updated_date: '2026-04-01 03:19'
labels:
  - renderer
  - mpv
  - keybindings
  - regression
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Regression: configured overlay keybindings that forward raw mpv subtitle-jump commands (for example previous-subtitle on H) can resume playback when invoked while paused. Keyboard-driven edge jumps already preserve paused state; configured keybindings should match that behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Configured subtitle-jump keybindings preserve paused playback state after backward seek
- [x] #2 Existing keyboard-driven subtitle navigation behavior remains unchanged
- [x] #3 Regression test covers paused configured subtitle-jump keybinding handling
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Configured overlay keybindings that forward `sub-seek` commands now re-check paused state and reapply pause after the seek when playback was already paused. This aligns raw configured subtitle-jump keybindings with the existing keyboard-driven edge-jump behavior and adds regression coverage for the paused backward-seek case.
<!-- SECTION:FINAL_SUMMARY:END -->
