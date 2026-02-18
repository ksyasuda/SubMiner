---
id: TASK-2.2
title: Add tests for overlay shortcut handler service
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
  - src/core/services/overlay-shortcut-handler.ts
parent_task_id: TASK-2
ordinal: 9000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add dedicated tests for `overlay-shortcut-handler.ts`, covering shortcut runtime handlers, fallback behavior, and key edge/error paths.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Shortcut registration/unregistration handler behavior is covered.
- [x] #2 Fallback handling paths are covered for valid and invalid input.
- [x] #3 Error and guard behavior is covered for missing dependencies/state.
- [x] #4 `pnpm run test:core` remains green.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added `src/core/services/overlay-shortcut-handler.test.ts` with coverage for:

- runtime handler dispatch for sync and async actions
- async error propagation to OSD/log handling
- local fallback action matching, including timeout forwarding
- `allowWhenRegistered` behavior for secondary subtitle toggle
- no-match fallback return behavior

Updated `package.json` `test:core` to include `dist/core/services/overlay-shortcut-handler.test.js`.

Verification: `pnpm run test:core` passed (19/19 at completion of this ticket).

<!-- SECTION:NOTES:END -->
