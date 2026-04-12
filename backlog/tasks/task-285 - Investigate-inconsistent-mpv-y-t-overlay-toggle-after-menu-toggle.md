---
id: TASK-285
title: Investigate inconsistent mpv y-t overlay toggle after menu toggle
status: To Do
assignee: []
created_date: '2026-04-07 22:55'
updated_date: '2026-04-07 22:55'
labels:
  - bug
  - overlay
  - keyboard
  - mpv
dependencies: []
references:
  - plugin/subminer/process.lua
  - plugin/subminer/ui.lua
  - src/renderer/handlers/keyboard.ts
  - src/main/runtime/autoplay-ready-gate.ts
  - src/core/services/overlay-window-input.ts
  - backlog/tasks/task-248 - Fix-macOS-visible-overlay-toggle-getting-immediately-restored.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User report: toggling the visible overlay with mpv `y-t` is inconsistent. After manually toggling through the `y-y` menu, `y-t` may allow one hide, but after toggling back on it can stop hiding the overlay again, forcing the user back into the menu path.

Initial assessment:

- no active backlog item currently tracks this exact report
- nearest prior work is `TASK-248`, which fixed a macOS-specific visible-overlay restore bug and is marked done
- current targeted regressions for the old fix surface pass, including plugin ready-signal suppression, focused-overlay `y-t` proxy dispatch, autoplay-ready gate deduplication, and blur-path restacking guards

This should be treated as a fresh investigation unless reproduction proves it is the same closed macOS issue resurfacing on the current build.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Reproduce the reported `y-t` / `y-y` inconsistency on the affected platform and identify the exact event sequence
- [ ] #2 Determine whether the failure is in mpv plugin command dispatch, focused-overlay key forwarding, or main-process visible-overlay state transitions
- [ ] #3 Fix the inconsistency so repeated hide/show/hide cycles work from `y-t` without requiring menu recovery
- [ ] #4 Add regression coverage for the reproduced failing sequence
- [ ] #5 Record whether this is a regression of `TASK-248` or a distinct bug
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the report with platform/build details and capture whether the failing `y-t` press originates in raw mpv or the focused overlay y-chord proxy path.
2. Trace visible-overlay state mutations across plugin toggle commands, autoplay-ready callbacks, and main-process visibility/window blur handling.
3. Patch the narrowest failing path and add regression coverage for the exact hide/show/hide sequence.
4. Re-run targeted plugin, overlay visibility, overlay window, and renderer keyboard suites before broader verification.
<!-- SECTION:PLAN:END -->
