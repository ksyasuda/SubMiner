---
id: TASK-238.3
title: Introduce domain type entrypoints and shrink src/types.ts import surface
status: Done
assignee: []
created_date: '2026-03-26 20:49'
updated_date: '2026-03-27 00:14'
labels:
  - tech-debt
  - types
  - maintainability
milestone: m-0
dependencies: []
references:
  - src/types.ts
  - src/shared/ipc/contracts.ts
  - src/config/service.ts
  - docs/architecture/README.md
parent_task_id: TASK-238
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/types.ts` has become the repo-wide dumping ground for unrelated domains. Splitting it is still worthwhile, but a big-bang move would create noisy churn across a large import graph. Introduce domain entrypoints under `src/types/` and migrate the highest-churn imports first while leaving `src/types.ts` as a compatibility barrel until the new structure is proven.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 Domain-focused type modules exist for the main clusters currently mixed together in `src/types.ts` (for example Anki, config/runtime, subtitle/media, and integration/runtime-option types).
- [x] #2 `src/types.ts` becomes a thinner compatibility layer or barrel instead of the sole source of truth for every shared type.
- [x] #3 A meaningful set of imports is migrated to the new entrypoints without breaking the maintained typecheck/test lanes.
- [x] #4 The new structure is documented well enough that contributors can tell where new shared types should live.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inventory the main type clusters in `src/types.ts` and choose stable domain seams.
2. Create `src/types/` modules and re-export through `src/types.ts` so the migration can be incremental.
3. Migrate the highest-value import sites first, especially config/runtime and Anki-heavy surfaces.
4. Verify with `bun run typecheck` and the cheapest test lane covering touched domains.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented domain entrypoints under `src/types/` and kept `src/types.ts` as a compatibility barrel (`src/types/anki.ts`, `src/types/config.ts`, `src/types/integrations.ts`, `src/types/runtime.ts`, `src/types/runtime-options.ts`, `src/types/subtitle.ts`). Migrated the highest-value import surfaces away from `src/types.ts` in config/runtime/Anki-related modules and shared IPC surfaces. Added type-level regression coverage in `src/types-domain-entrypoints.type-test.ts`.

Aligned docs in `docs/architecture/README.md`, `docs/architecture/domains.md`, and `docs-site/changelog.md` to support the change and clear docs-site sync mismatch.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Task completed with commit `5dd8bb7f` (`refactor: split shared type entrypoints`). The refactor introduced domain type entrypoints, shrank the `src/types.ts` import surface, updated import consumers, and recorded verification evidence in the local verifier artifacts. Backlog now tracks TASK-238.3 as done.
<!-- SECTION:FINAL_SUMMARY:END -->
