---
id: TASK-238.5
title: Split immersion tracker query layer into focused read-model modules
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - tech-debt
  - stats
  - database
  - maintainability
milestone: m-0
dependencies:
  - TASK-238.3
references:
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/stats-server.ts
  - src/core/services/immersion-tracker-service.ts
parent_task_id: TASK-238
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/core/services/immersion-tracker/query.ts` has grown into a large mixed read/write/maintenance surface that owns library queries, timeline/detail queries, cleanup helpers, and rollup rebuild hooks. That size makes stats work harder to change safely. Split the query layer into focused read-model and maintenance modules so future stats/dashboard work does not keep landing in one 2500-line file.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Query responsibilities are grouped into focused modules such as library/session detail, vocabulary/kanji detail, and maintenance/cleanup helpers.
- [ ] #2 The stats server and immersion tracker service depend on stable exported query surfaces instead of one monolithic file.
- [ ] #3 The refactor preserves current SQL behavior and existing statistics outputs.
- [ ] #4 Existing stats/immersion tests still pass, with added focused coverage where extraction creates new seams.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inventory the major query clusters and choose modules that match current caller boundaries.
2. Extract without changing schema or response contracts unless a narrow cleanup is required for compile/test health.
3. Keep SQL ownership close to the domain module that consumes it; avoid a giant `queries/` dump with no structure.
4. Verify with the maintained stats/immersion test lane plus `bun run typecheck`.
<!-- SECTION:PLAN:END -->
