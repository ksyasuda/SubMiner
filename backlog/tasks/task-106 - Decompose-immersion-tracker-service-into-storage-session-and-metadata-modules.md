---
id: TASK-106
title: Decompose immersion tracker service into storage session and metadata modules
status: To Do
assignee: []
created_date: '2026-02-22 07:14'
updated_date: '2026-02-22 07:14'
labels:
  - refactor
  - maintainability
  - immersion-tracking
dependencies:
  - TASK-95
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/core/services/immersion-tracker-service.ts` remains large (~1100 LOC) and still mixes multiple concerns in one class:
- queue/flush orchestration,
- DB schema and SQL lifecycle,
- session state transitions,
- local media metadata probing (`ffprobe`/hashing).

Further decomposition is needed to keep ownership boundaries clear and reduce refactor risk.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract SQLite schema+statement setup into a storage module dedicated to DB lifecycle.
2. Extract session-state transition/event-recording logic into a session runtime module.
3. Extract local metadata probing (hash + ffprobe parsing) into a metadata adapter module.
4. Keep `ImmersionTrackerService` as orchestration facade over extracted collaborators.
5. Expand seam tests for extracted modules and reduce skipped tracker coverage where feasible.
6. Verify with source tracker tests and full build.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `immersion-tracker-service.ts` no longer embeds full schema SQL and metadata probing logic directly.
- [ ] #2 Extracted modules have focused tests for session transitions, DB writes, and metadata parsing.
- [ ] #3 Tracker behavior remains unchanged (session lifecycle, rollups, retention, queue semantics).
- [ ] #4 Build and tracker-related source tests pass.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Service file size reduced materially from current baseline.
- [ ] #2 Ownership boundaries documented in `docs/architecture.md` or relevant service docs.
- [ ] #3 No regression in `bun run test:core:src` immersion tracker coverage.
<!-- DOD:END -->

