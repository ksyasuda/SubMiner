---
id: TASK-30.5
title: >-
  Wire mpv script-message/shortcut trigger into streaming modal and playback
  path
status: To Do
assignee: []
created_date: '2026-02-13 18:33'
updated_date: '2026-02-13 18:34'
labels: []
dependencies:
  - TASK-30.3
  - TASK-30.4
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Allow users to open streaming selection from within mpv via keybind/menu and route that intent into renderer modal and playback flow without requiring separate window focus changes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Add/extend mpv Lua plugin or existing command registry to emit a custom action for opening the streaming picker.
- [ ] #2 Handle this action in main IPC/mpv-command pipeline and forward to renderer modal state.
- [ ] #3 Add at least one default keybinding/menu entry documented in config/plugin notes.
- [ ] #4 Ensure playback launched from the in-player flow uses the same command path and error messages as renderer-initiated flow.
- [ ] #5 Add graceful handling when feature is disabled or no providers are configured.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Phase 5 — In-player entry: mpv trigger/menu integration
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 No duplicate mpv command parsing between picker and legacy commands.
- [ ] #2 Feature can be used in overlay and mpv-only mode where applicable.
- [ ] #3 No dependency on modal open state when launched by mpv trigger.
- [ ] #4 Manual and keybind invocations behave consistently.
<!-- DOD:END -->
