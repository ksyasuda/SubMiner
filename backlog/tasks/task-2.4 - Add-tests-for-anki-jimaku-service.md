---
id: TASK-2.4
title: Add tests for anki jimaku service
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:56'
updated_date: '2026-02-18 04:11'
labels: []
dependencies:
  - TASK-2.1
references:
  - investigation.md
  - src/core/services/anki-jimaku-service.ts
parent_task_id: TASK-2
ordinal: 7000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add dedicated tests for `anki-jimaku-service.ts` focusing on IPC handler registration, request dispatch, and error handling behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 IPC registration behavior is validated for all channels exposed by this service.
- [x] #2 Success-path behavior for core handler flows is validated.
- [x] #3 Failure-path behavior is validated with expected error propagation.
- [x] #4 `pnpm run test:core` remains green.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a lightweight registration-injection seam to `registerAnkiJimakuIpcRuntimeService` so runtime behavior can be tested without Electron IPC globals.

Added `src/core/services/anki-jimaku-service.test.ts` with coverage for:
- runtime handler surface registration
- integration disable path and runtime-options broadcast
- subtitle history clear and field-grouping response callbacks
- merge-preview guard error and integration success delegation
- Jimaku search request mapping/result capping
- downloaded-subtitle MPV command forwarding

Updated `package.json` `test:core` to include `dist/core/services/anki-jimaku-service.test.js`.

Verification: `pnpm run test:core` passed (21/21).
<!-- SECTION:NOTES:END -->
