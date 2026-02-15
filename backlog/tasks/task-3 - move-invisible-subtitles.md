---
id: TASK-3
title: move invisible subtitles
status: Done
assignee:
  - codex
created_date: '2026-02-11 03:34'
updated_date: '2026-02-11 04:28'
labels: []
dependencies: []
ordinal: 1000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add keybinding that will toggle edit mode on the invisible subtitles allowing for fine-grained control over positioning. use arrow keys and vim hjkl for motion and enter/ctrl+s to save and esc to cancel
<!-- SECTION:DESCRIPTION:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
- Implemented invisible subtitle position edit mode toggle with movement/save/cancel controls.
- Added persistence for invisible subtitle offsets (`invisibleOffsetXPx`, `invisibleOffsetYPx`) alongside existing `yPercent` subtitle position state.
- Updated edit mode visuals to highlight invisible subtitle text using the same styling as debug visualization.
- Removed the edit-mode dashed bounding box.
- Updated top HUD instruction text to reference arrow keys only (while keeping `hjkl` movement support).
<!-- SECTION:NOTES:END -->
