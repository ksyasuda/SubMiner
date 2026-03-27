---
id: TASK-238.3
title: Introduce domain type entrypoints and shrink src/types.ts import surface
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
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
- [ ] #1 Domain-focused type modules exist for the main clusters currently mixed together in `src/types.ts` (for example Anki, config/runtime, subtitle/media, and integration/runtime-option types).
- [ ] #2 `src/types.ts` becomes a thinner compatibility layer or barrel instead of the sole source of truth for every shared type.
- [ ] #3 A meaningful set of imports is migrated to the new entrypoints without breaking the maintained typecheck/test lanes.
- [ ] #4 The new structure is documented well enough that contributors can tell where new shared types should live.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inventory the main type clusters in `src/types.ts` and choose stable domain seams.
2. Create `src/types/` modules and re-export through `src/types.ts` so the migration can be incremental.
3. Migrate the highest-value import sites first, especially config/runtime and Anki-heavy surfaces.
4. Verify with `bun run typecheck` and the cheapest test lane covering touched domains.
<!-- SECTION:PLAN:END -->
