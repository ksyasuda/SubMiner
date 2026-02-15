---
id: TASK-20.2
title: Add renderer-to-main IPC contract for measured overlay content bounds
status: Done
assignee: []
created_date: '2026-02-12 08:47'
updated_date: '2026-02-13 08:05'
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
- [x] #1 Preload exposes a typed API for reporting overlay content bounds with layer metadata.
- [x] #2 Main-process IPC handler validates payload shape/range and stores latest measurement per layer.
- [x] #3 Renderer emits measurement updates on subtitle, mode, style, and render-metric changes with throttling/debounce.
- [x] #4 No crashes or unbounded logging when measurements are missing/empty/invalid; fallback behavior is explicit.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a typed `OverlayContentMeasurement` IPC contract exposed in preload and Electron API typings. Implemented a main-process measurement store with strict payload validation and rate-limited warning logs for invalid reports. Added renderer-side debounced measurement reporting that emits updates on subtitle content/mode/style/render-metric and resize changes, explicitly sending `contentRect: null` when no measured content exists to signal fallback behavior.

Added unit coverage for measurement validation and store behavior.

Closed per user request to delete parent task and subtasks.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented renderer-to-main measurement reporting for overlay content bounds with per-layer metadata. Main now validates and stores latest measurements per layer safely, renderer emits debounced updates on relevant state changes, and invalid/missing payload handling is explicit and non-spammy.
<!-- SECTION:FINAL_SUMMARY:END -->
