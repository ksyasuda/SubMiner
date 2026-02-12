---
id: TASK-20.2
title: Add renderer-to-main IPC contract for measured overlay content bounds
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
Add renderer-to-main IPC for content measurement reporting, so main process can size each overlay window from post-layout DOM bounds.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Preload exposes a typed API for reporting overlay content bounds with layer metadata.
- [ ] #2 Main-process IPC handler validates payload shape/range and stores latest measurement per layer.
- [ ] #3 Renderer emits measurement updates on subtitle, mode, style, and render-metric changes with throttling/debounce.
- [ ] #4 No crashes or unbounded logging when measurements are missing/empty/invalid; fallback behavior is explicit.
<!-- AC:END -->
