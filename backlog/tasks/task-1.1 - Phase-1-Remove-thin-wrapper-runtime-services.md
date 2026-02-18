---
id: TASK-1.1
title: 'Phase 1: Remove thin wrapper runtime services'
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:46'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
references:
  - plan.md
  - src/main.ts
  - src/core/services/index.ts
parent_task_id: TASK-1
ordinal: 12000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Inline trivial wrapper services into their call sites and delete redundant service/test files listed in Phase 1 of plan.md.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Wrapper logic from the Phase 1 file list is inlined at call sites without behavior changes.
- [x] #2 Phase 1 wrapper service files and corresponding trivial tests are removed from the codebase.
- [x] #3 `src/core/services/index.ts` exports are updated to remove deleted modules.
- [x] #4 `pnpm run build && pnpm run test:core` passes after Phase 1 completion.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Locate all Phase 1 wrapper service call sites and classify direct-inline substitutions vs orchestration-flow inlines.
2. Remove the lowest-risk wrappers first (`config-warning-runtime-service`, `app-logging-runtime-service`, `runtime-options-manager-runtime-service`, `overlay-modal-restore-service`, `overlay-send-service`) and update imports/exports.
3. Continue with startup and shutdown wrappers (`startup-resource-runtime-service`, `config-generation-runtime-service`, `app-shutdown-runtime-service`, `shortcut-ui-deps-runtime-service`) by inlining behavior into `main.ts` or direct callers.
4. Delete corresponding trivial test files for removed wrappers and clean `src/core/services/index.ts` exports.
5. Run `pnpm run build && pnpm run test:core`; fix regressions and update task notes/acceptance criteria incrementally.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Inlined wrapper behaviors into direct call sites in `main.ts` and `overlay-bridge-runtime-service.ts` for config warning/app logging, runtime options manager construction, generate-config bootstrap path, startup resource initialization, app shutdown sequence, overlay modal restore handling, overlay send behavior, and overlay shortcut local fallback invocation.

Deleted 16 Phase 1 files (9 wrapper services + 7 wrapper tests) and removed corresponding barrel exports from `src/core/services/index.ts`.

Updated `package.json` `test:core` list to remove deleted test entries so the script tracks current sources accurately.

Verification: `pnpm run build && pnpm run test:core` passes after refactor.

<!-- SECTION:NOTES:END -->
