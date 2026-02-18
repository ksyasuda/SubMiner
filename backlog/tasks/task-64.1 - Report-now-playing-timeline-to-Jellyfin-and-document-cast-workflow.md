---
id: TASK-64.1
title: Report now-playing timeline to Jellyfin and document cast workflow
status: To Do
assignee:
  - '@sudacode'
created_date: '2026-02-17 21:25'
labels:
  - jellyfin
  - docs
  - telemetry
dependencies: []
parent_task_id: TASK-64
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Send playback start/progress/stop updates from SubMiner to Jellyfin during cast sessions and document configuration/usage/troubleshooting for the new mode.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 SubMiner posts playing/progress/stopped updates for casted sessions at a reasonable interval.
- [ ] #2 Timeline reporting failures do not crash playback and are logged at debug/warn levels.
- [ ] #3 Jellyfin integration docs include cast-to-device setup, expected behavior, and troubleshooting.
- [ ] #4 Regression tests for reporting payload construction and error handling are added.
<!-- AC:END -->
