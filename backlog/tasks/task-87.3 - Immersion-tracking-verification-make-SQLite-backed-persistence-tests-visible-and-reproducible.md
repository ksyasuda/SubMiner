---
id: TASK-87.3
title: >-
  Immersion tracking verification: make SQLite-backed persistence tests visible
  and reproducible
status: To Do
assignee: []
created_date: '2026-03-06 03:19'
updated_date: '2026-03-06 03:21'
labels:
  - tests
  - immersion-tracking
milestone: m-0
dependencies: []
references:
  - src/core/services/immersion-tracker-service.test.ts
  - src/core/services/immersion-tracker/storage-session.test.ts
  - src/core/services/immersion-tracker-service.ts
  - package.json
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The immersion tracker is persistence-heavy, but its SQLite-backed tests are conditionally skipped in the standard Bun run when node:sqlite support is unavailable. That creates a blind spot around session finalization, telemetry persistence, and retention behavior. This task should establish a reliable automated verification path for the database-backed cases and make the prerequisite/runtime behavior explicit to contributors and CI.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Database-backed immersion tracking tests run in at least one documented automated command that is practical for contributors or CI to execute.
- [ ] #2 If the current runtime cannot execute the SQLite-backed tests, the repository exposes that limitation clearly instead of silently reporting a misleading green result.
- [ ] #3 Contributor-facing documentation explains how to run the immersion tracker verification lane and any environment prerequisites it depends on.
- [ ] #4 The resulting verification covers session persistence or finalization behavior that is not exercised by the pure seam tests alone.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Confirm which SQLite-backed immersion tests are currently skipped and why in the standard Bun environment.
2. Establish a reproducible command or lane for the DB-backed cases, or make the unsupported-runtime limitation explicit and actionable.
3. Document prerequisites and expected behavior for contributors and CI.
4. Verify at least one persistence/finalization path beyond the seam tests is exercised by the new lane.
<!-- SECTION:PLAN:END -->
