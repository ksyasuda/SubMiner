---
id: TASK-255
title: Add overlay playlist browser modal for sibling video files and mpv queue
status: To Do
assignee: []
created_date: '2026-03-30 05:46'
labels:
  - feature
  - overlay
  - mpv
  - launcher
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add an in-session overlay modal that opens from a keybinding during active playback and lets the user browse video files from the current file's parent directory alongside the active mpv playlist. The modal should sort local files in best-effort episode order, highlight the current item, and allow keyboard/mouse interaction to add files into the mpv queue, remove queued items, and reorder queued items without leaving playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 An overlay modal can be opened during active playback from a dedicated keybinding and closed without disrupting existing modal behavior.
- [ ] #2 The modal shows video files from the current media file's parent directory in best-effort episode order and highlights the current file when present.
- [ ] #3 The modal shows the active mpv playlist/queue with enough metadata to identify the current item and queued order.
- [ ] #4 The user can add a directory file to the mpv playlist, remove playlist items, and reorder playlist items from the modal using both mouse and keyboard interactions.
- [ ] #5 Modal state stays in sync after playlist mutations so the rendered queue reflects mpv's current playlist order.
- [ ] #6 Feature coverage includes automated tests for ordering/playlist behavior and docs or shortcut/help updates for the new modal.
<!-- AC:END -->
