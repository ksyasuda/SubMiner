---
id: TASK-20.3
title: >-
  Implement content-bounded sizing algorithm for visible and invisible overlay
  windows
status: To Do
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-12 09:42'
labels: []
dependencies: []
parent_task_id: TASK-20
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Implement content-bounded sizing for visible/invisible windows using measured rects plus tracker origin, with robust clamping and jitter resistance.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Bounds algorithm applies configurable padding, minimum size, display-workarea clamp, and integer snap.
- [ ] #2 Main-process bounds updates are thresholded/debounced to reduce jitter and unnecessary `setBounds` churn.
- [ ] #3 When no valid measurement exists, layer falls back to safe tracker/display bounds without breaking interaction.
- [ ] #4 Visible+invisible overlays can coexist without full-window overlap/input conflicts caused by shared fullscreen bounds.
<!-- AC:END -->
