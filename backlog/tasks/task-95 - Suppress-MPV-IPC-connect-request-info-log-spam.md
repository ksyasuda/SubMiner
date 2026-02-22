---
id: TASK-95
title: Suppress MPV IPC connect-request info log spam
status: Done
assignee: []
created_date: '2026-02-21 04:38'
updated_date: '2026-02-22 07:49'
labels:
  - logging
  - mpv
  - electron
  - quality
dependencies: []
priority: high
ordinal: 91000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow-up to TASK-33: `[main:mpv] MPV IPC connect requested.` still emits at INFO repeatedly during startup retries. Gate this request-attempt log behind debug level so normal runs stay quiet.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Non-debug runs do not emit repeated 'MPV IPC connect requested.' lines during retry loops.
- [x] #2 Debug runs still emit connect-request attempt logs for diagnosis.
- [x] #3 Connection behavior and retry timing are unchanged.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Moved MPV connect-request log in `src/core/services/mpv.ts` from INFO to DEBUG so retry loops are silent at default log level while preserving diagnostics in debug mode. Added regression tests in `src/core/services/mpv.test.ts` asserting zero connect-request logs at info level and one connect-request log at debug level. Validation: `bun run build && node dist/core/services/mpv.test.js` (pass).
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Add/adjust unit tests covering debug vs non-debug logging for connect requests.
- [x] #2 Run targeted mpv service tests and build/typecheck path used for this change.
<!-- DOD:END -->
