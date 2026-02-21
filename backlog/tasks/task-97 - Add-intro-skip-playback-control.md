---
id: TASK-97
title: Add intro skip playback control
status: To Do
assignee: []
created_date: '2026-02-21 04:41'
updated_date: '2026-02-21 04:41'
labels:
  - playback
  - ux
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add an intro skip control so users can jump past opening sequences quickly during playback. Start with a reliable manual control (shortcut/action) and clear user feedback after seek.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Add a configurable skip duration (for example 60/75/90 seconds).
- Expose skip intro via keybinding and optional UI action in overlay/help.
- Show transient confirmation (OSD/overlay message) after skip action.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define config and keybinding surface for intro skip duration and trigger.
2. Implement intro skip command that performs bounded seek in active playback session.
3. Wire command to user trigger path (keyboard + optional on-screen action if present).
4. Emit user feedback after successful skip (current time + skipped duration).
5. Add tests for command dispatch, seek bounds, and config fallback behavior.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 User can trigger intro skip during playback with configured shortcut/action.
- [ ] #2 Skip performs bounded seek and never seeks before start or beyond stream duration.
- [ ] #3 Skip duration is configurable with sane default.
- [ ] #4 User receives visible confirmation after skip.
- [ ] #5 Automated tests cover config + seek behavior.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Playback control tests pass
- [ ] #2 User-facing config/docs updated for intro skip control
<!-- DOD:END -->
