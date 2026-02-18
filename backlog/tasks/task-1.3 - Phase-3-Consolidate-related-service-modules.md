---
id: TASK-1.3
title: 'Phase 3: Consolidate related service modules'
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:46'
updated_date: '2026-02-18 04:11'
labels: []
dependencies:
  - TASK-1.2
references:
  - plan.md
  - src/core/services/overlay-visibility-service.ts
  - src/core/services/overlay-manager-service.ts
  - src/core/services/overlay-shortcut-service.ts
  - src/core/services/numeric-shortcut-session-service.ts
  - src/core/services/app-ready-runtime-service.ts
parent_task_id: TASK-1
ordinal: 6000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Merge split modules for overlay visibility, broadcast, shortcuts, numeric shortcuts, and startup orchestration into cohesive service files.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Overlay visibility/runtime split is consolidated into a single service module.
- [x] #2 Overlay broadcast functions are merged with overlay manager responsibilities.
- [x] #3 Shortcut and numeric shortcut runtime/lifecycle splits are consolidated as described in plan.md.
- [x] #4 Startup bootstrap and app-ready runtime orchestration is consolidated into one startup module.
- [x] #5 `pnpm run build && pnpm run test:core` passes after Phase 3 completion.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Merge `overlay-visibility-runtime-service.ts` exports into `overlay-visibility-service.ts` and update imports/exports.
2. Merge overlay broadcast responsibilities from `overlay-broadcast-runtime-service.ts` into `overlay-manager-service.ts` while preserving current APIs used by `main.ts`.
3. Consolidate shortcut modules by absorbing lifecycle utilities into `overlay-shortcut-service.ts` and fallback-runner logic into `overlay-shortcut-runtime-service.ts` (or successor handler module), then remove obsolete files.
4. Merge numeric shortcut runtime/session split into a single `numeric-shortcut-service.ts` and update call sites/tests.
5. Merge startup bootstrap + app-ready orchestration into a single startup module, update imports, remove obsolete files, and run `pnpm run build && pnpm run test:core`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Merged overlay visibility runtime API into `overlay-visibility-service.ts` and removed `overlay-visibility-runtime-service.ts`.

Merged overlay broadcast behavior into `overlay-manager-service.ts` (including manager-level broadcasting) and removed `overlay-broadcast-runtime-service.ts` + test, with equivalent coverage moved into `overlay-manager-service.test.ts`.

Consolidated shortcut modules into `overlay-shortcut-service.ts` (lifecycle) and new `overlay-shortcut-handler.ts` (runtime handlers + local fallback), removing `overlay-shortcut-lifecycle-service.ts`, `overlay-shortcut-runtime-service.ts`, and `overlay-shortcut-fallback-runner.ts`.

Merged numeric shortcut runtime/session split into `numeric-shortcut-service.ts`; removed `numeric-shortcut-runtime-service.ts` and merged runtime test coverage into session tests.

Merged startup bootstrap + app-ready orchestration into `startup-service.ts`; removed `startup-bootstrap-runtime-service.ts` and `app-ready-runtime-service.ts` with tests updated to new module path.

Verification: `pnpm run build && pnpm run test:core` passes after consolidation.

<!-- SECTION:NOTES:END -->
