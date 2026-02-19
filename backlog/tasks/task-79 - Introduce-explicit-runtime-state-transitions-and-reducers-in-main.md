---
id: TASK-79
title: Introduce explicit runtime state transitions and reducers in main
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - main-process
  - state-management
  - architecture
dependencies:
  - TASK-71
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Main runtime state is currently mutable from many callsites. This task introduces explicit state-transition functions (reducer-like actions) for critical state domains to improve traceability and reduce accidental cross-feature coupling.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Start with high-risk domains: overlay visibility, MPV connection/session, AniList/Jellyfin runtime status.
- Use action-style transition helpers (`setMediaPath`, `setOverlayVisibility`, `setAnilistState`) instead of direct field mutation.
- Keep transition helpers side-effect free; invoke effects in orchestration layer.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Inventory direct mutation hotspots in `src/main.ts` and runtime modules.
2. Define state transition API in `src/main/state.ts` for priority domains.
3. Migrate callsites incrementally to transition helpers.
4. Add unit tests for transition behavior and invariant enforcement.
5. Add lightweight debug tracing for transition events in debug mode.
6. Document state ownership and mutation rules in architecture docs.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Critical runtime state domains mutate through explicit transition helpers
- [ ] #2 State invariants are test-covered
- [ ] #3 Direct ad-hoc mutation in migrated domains is removed
- [ ] #4 Ownership/mutation rules documented
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Core tests pass after migration
- [ ] #2 No behavior regressions in startup/IPC/overlay flows
<!-- DOD:END -->

