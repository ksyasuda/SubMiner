---
id: TASK-75
title: Move mpv OSD log writes to buffered async path
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - logging
  - performance
  - main-process
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`appendToMpvLog` in `src/main.ts` uses synchronous fs operations on OSD writes. Frequent OSD activity can block the event loop and adds avoidable latency in main process runtime paths.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Replace sync file writes with buffered async logger flow.
2. Ensure log directory creation and write failures remain best-effort (non-fatal).
3. Add flush behavior for shutdown to avoid dropping trailing log lines.
4. Keep log format/path unchanged unless explicitly intended.
5. Add tests for non-blocking behavior and graceful error handling.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 OSD log path no longer uses synchronous fs calls
- [ ] #2 Logging remains best-effort and does not break runtime behavior on fs failure
- [ ] #3 Shutdown flush behavior verified by tests
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Core tests pass with new logging path
- [ ] #2 No user-visible regression in MPV OSD/log output behavior
<!-- DOD:END -->

