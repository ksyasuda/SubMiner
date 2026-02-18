---
id: TASK-2.3
title: Add tests for mining service
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
  - src/core/services/mining-service.ts
parent_task_id: TASK-2
ordinal: 8000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add dedicated behavior tests for `mining-service.ts` covering sentence/card mining orchestration and error boundaries.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Happy-path behavior is covered for mining entry points.
- [x] #2 Guard/early-return behavior is covered for missing runtime state.
- [x] #3 Error paths are covered with expected logging/OSD behavior.
- [x] #4 `pnpm run test:core` remains green.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added `src/core/services/mining-service.test.ts` with focused coverage for:

- `copyCurrentSubtitleService` guard and success behavior
- `mineSentenceCardService` integration/connection guards and success path
- `handleMultiCopyDigitService` history-copy behavior with truncation messaging
- `handleMineSentenceDigitService` async error catch and OSD/log propagation

Updated `package.json` `test:core` to include `dist/core/services/mining-service.test.js`.

Verification: `pnpm run test:core` passed (20/20 after adding mining tests).

<!-- SECTION:NOTES:END -->
