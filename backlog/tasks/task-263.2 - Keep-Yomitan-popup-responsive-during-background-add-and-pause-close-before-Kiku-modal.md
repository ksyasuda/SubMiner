---
id: TASK-263.2
title: >-
  Keep Yomitan popup responsive during background add and pause/close before
  Kiku modal
status: Done
assignee: []
created_date: '2026-04-01 00:42'
updated_date: '2026-04-01 02:35'
labels:
  - anki
  - yomitan
  - kiku
  - ux
dependencies: []
parent_task_id: TASK-263
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Make Yomitan popup add run in background without blocking popup responsiveness. Before opening Kiku field-grouping modal, pause MPV and close the Yomitan popup/parser window if open.
<!-- SECTION:DESCRIPTION:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Popup save path returns immediately and prevents duplicate submits
- [x] #2 Field-grouping modal request pauses MPV and closes Yomitan popup window first
- [x] #3 Regression tests cover async save dispatch and main-side pause/close hook
<!-- DOD:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-31: Removed the custom pending label/gray save-button presentation from vendored Yomitan. Background add still runs asynchronously with the internal pending-save guard, so duplicate clicks are ignored while the button keeps its stock appearance.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Yomitan popup save dispatches note creation/add in the background with an internal pending-save guard so repeated clicks are ignored without blocking the popup. Before opening the Kiku field-grouping modal, the renderer now closes the visible lookup popup and pauses MPV. Follow-up UX polish removed the custom pending label/gray styling so the save button keeps Yomitan’s stock presentation while the background action runs.
<!-- SECTION:FINAL_SUMMARY:END -->
