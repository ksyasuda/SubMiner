---
id: TASK-2.1
title: Remove unused scaffolding and clean exports
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:56'
updated_date: '2026-02-11 03:35'
labels: []
dependencies:
  - TASK-1
references:
  - investigation.md
  - src/core/action-bus.ts
  - src/core/actions.ts
  - src/core/app-context.ts
  - src/core/module-registry.ts
  - src/core/module.ts
  - src/modules/
  - src/ipc/
  - src/core/services/index.ts
parent_task_id: TASK-2
ordinal: 10000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove unused module-architecture scaffolding and IPC abstraction files identified as dead code, and clean service barrel exports that are not needed outside service internals.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Files under `src/core/{action-bus.ts,actions.ts,app-context.ts,module-registry.ts,module.ts}` are removed if unreferenced.
- [x] #2 Unused `src/modules/` and `src/ipc/` scaffolding files are removed if unreferenced.
- [x] #3 `src/core/services/index.ts` no longer exports symbols that are only consumed internally (`isGlobalShortcutRegisteredSafe`).
- [x] #4 Build and core tests pass after cleanup.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Verify all candidate files are truly unreferenced in runtime/test paths.
2. Delete dead scaffolding files and folders.
3. Remove unnecessary service barrel exports and fix any import fallout.
4. Run `pnpm run build` and `pnpm run test:core`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed unused scaffolding files from `src/core/`, `src/modules/`, and `src/ipc/` that were unreferenced by runtime code.

Updated `src/core/services/index.ts` to stop re-exporting `isGlobalShortcutRegisteredSafe`, which is only used internally by service files.

Verification:
- `pnpm run build` passed
- `pnpm run test:core` passed (18/18)
<!-- SECTION:NOTES:END -->
