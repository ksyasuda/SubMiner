---
id: TASK-30.8
title: >-
  Add observability and tuning metrics for Aniyomi subtitle-source fallback
  decisions
status: To Do
assignee: []
created_date: '2026-02-13 18:41'
labels: []
dependencies:
  - TASK-30.7
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add lightweight telemetry/analytics hooks (local logs + optional structured counters) to measure how Aniyomi/anime streaming source selection behaves, including soft-sub preference, hard-sub fallback usage, and ffmpeg+Jimaku post-processing outcomes, to support source ranking tuning.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Track per-playback decision metadata including chosen source, language match score, subtitle mode (soft/hard), and reason for source preference ordering.
- [ ] #2 Emit success/failure counters for hard-sub stripping attempts (started/succeeded/failed/unsupported codec) with reason codes.
- [ ] #3 Log whether Jimaku subtitle attachment was available and successfully loaded for ffmpeg-assisted flows.
- [ ] #4 Capture user-visible fallback reasons when preferred English/soft-sub sources are absent and hard-sub path is used.
- [ ] #5 Add a debug/report view or log artifact with counters that can be reviewed in-app or via config/log files.
- [ ] #6 Document metrics definitions so developers can tune source scorer and fallback policy without code changes.
- [ ] #7 Ensure instrumentation has low overhead and is opt-out-safe with existing config flags.
<!-- AC:END -->
