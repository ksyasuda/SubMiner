---
id: TASK-77
title: 'Subtitle hover: auto-pause playback with config toggle'
status: Done
assignee: []
created_date: '2026-02-28 22:43'
updated_date: '2026-02-28 22:43'
labels: []
dependencies: []
priority: medium
ordinal: 8000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add a user-facing subtitle config option to pause mpv playback when the cursor hovers subtitle text and resume playback when the cursor leaves.

Scope:
- New config key: `subtitleStyle.autoPauseVideoOnHover`.
- Default should be enabled.
- Hover pause/resume must not unpause if playback was already paused before hover.
- Docs/examples/tests updated.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 `subtitleStyle.autoPauseVideoOnHover` exists and defaults to `true`.
- [x] #2 Overlay pauses playback on subtitle hover and resumes on leave only when hover-triggered pause occurred.
- [x] #3 Main/renderer IPC exposes pause-state query for safe hover behavior.
- [x] #4 Config docs/examples and user docs/readme mention the new behavior and toggle.
- [x] #5 Regression tests cover config parsing/validation and hover behavior edge cases.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented `subtitleStyle.autoPauseVideoOnHover` with default `true`, wired through config defaults/resolution/types, renderer state/style, and mouse hover handlers. Added playback pause-state IPC (`getPlaybackPaused`) to avoid false resume when media was already paused. Added renderer hover behavior tests (including race/cancel case) and config/resolve tests. Updated config examples and docs (`README`, usage, shortcuts, mining workflow, configuration) to document default hover pause/resume behavior and disable path.

<!-- SECTION:FINAL_SUMMARY:END -->
