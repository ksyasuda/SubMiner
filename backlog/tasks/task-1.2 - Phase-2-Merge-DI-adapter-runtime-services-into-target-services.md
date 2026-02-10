---
id: TASK-1.2
title: 'Phase 2: Merge DI adapter runtime services into target services'
status: Done
assignee:
  - codex
created_date: '2026-02-10 18:46'
updated_date: '2026-02-10 19:00'
labels: []
dependencies:
  - TASK-1.1
references:
  - plan.md
  - src/core/services/cli-command-service.ts
  - src/core/services/ipc-service.ts
  - src/core/services/tokenizer-service.ts
  - src/core/services/app-lifecycle-deps-runtime-service.ts
parent_task_id: TASK-1
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Absorb dependency adapter runtime services into core service modules and remove adapter files/tests while preserving runtime behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 CLI, IPC, tokenizer, and app lifecycle adapter logic is merged into their target service modules.
- [x] #2 Adapter service and adapter test files listed in Phase 2 are removed.
- [x] #3 Callers pass dependency shapes expected by updated services without redundant mapping layers.
- [x] #4 `pnpm run build && pnpm run test:core` passes after Phase 2 completion.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Audit `cli-command-deps-runtime-service.ts`, `ipc-deps-runtime-service.ts`, `tokenizer-deps-runtime-service.ts`, and `app-lifecycle-deps-runtime-service.ts` usage sites in `main.ts` and corresponding services.
2. For each adapter, move null-guarding and shape-normalization logic into its target service (`cli-command-service.ts`, `ipc-service.ts`, `tokenizer-service.ts`, `app-lifecycle-service.ts`) and simplify caller dependency objects.
3. Remove adapter service files/tests and update `src/core/services/index.ts` exports/import sites.
4. Run `pnpm run build && pnpm run test:core` and fix any typing/regression issues from the interface consolidation.
5. Update task notes with dependency-shape decisions to preserve handoff clarity.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Merged adapter-constructor logic into target services: `createCliCommandDepsRuntimeService` moved into `cli-command-service.ts`, `createIpcDepsRuntimeService` moved into `ipc-service.ts`, `createTokenizerDepsRuntimeService` moved into `tokenizer-service.ts`, and `createAppLifecycleDepsRuntimeService` moved into `app-lifecycle-service.ts`.

Deleted adapter service files and tests for cli-command deps, ipc deps, tokenizer deps, and app lifecycle deps.

Updated `src/core/services/index.ts` exports and `package.json` `test:core` entries to remove deleted adapter test modules.

Verification: `pnpm run build && pnpm run test:core` passes after consolidation.
<!-- SECTION:NOTES:END -->
