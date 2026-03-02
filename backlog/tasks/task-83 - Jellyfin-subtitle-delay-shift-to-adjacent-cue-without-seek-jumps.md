---
id: TASK-83
title: 'Jellyfin subtitle delay: shift to adjacent cue without seek jumps'
status: Done
assignee: []
created_date: '2026-03-02 00:06'
updated_date: '2026-03-02 00:06'
labels: []
dependencies: []
priority: high
ordinal: 9003
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add keybinding-friendly special commands that shift `sub-delay` to align current subtitle start with next/previous cue start, without `sub-seek` probing (avoid playback jump).

Scope:

- add special commands for next/previous line alignment;
- compute delta from active subtitle cue timeline (external subtitle file/URL, including Jellyfin-delivered URLs);
- apply `add sub-delay <delta>` and show OSD value;
- keep existing proxy OSD behavior for direct `sub-delay` keybinding commands.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 New special commands exist for subtitle-delay shift to next/previous cue boundary.
- [x] #2 Shift logic parses active external subtitle source timings (SRT/VTT/ASS) and computes delta from current `sub-start`.
- [x] #3 Runtime applies delay shift without `sub-seek` and shows OSD feedback.
- [x] #4 Direct `sub-delay` proxy commands also show OSD current value.
- [x] #5 Tests added for cue parsing/shift behavior and IPC dispatch wiring.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented no-jump subtitle-delay alignment commands:

- added `__sub-delay-next-line` and `__sub-delay-prev-line` special commands;
- added `createShiftSubtitleDelayToAdjacentCueHandler` to parse cue start times from active external subtitle source and apply `add sub-delay` delta from current `sub-start`;
- wired command handling through IPC runtime deps into main runtime;
- retained/extended OSD proxy feedback for `sub-delay` keybindings;
- updated configuration docs and added regression tests for subtitle-delay shift and IPC command routing.

<!-- SECTION:FINAL_SUMMARY:END -->
