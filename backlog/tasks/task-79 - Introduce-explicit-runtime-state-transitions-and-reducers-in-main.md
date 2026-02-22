---
id: TASK-79
title: Introduce explicit runtime state transitions and reducers in main
status: Done
assignee:
  - '@sudacode'
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 07:49'
labels:
  - main-process
  - state-management
  - architecture
dependencies:
  - TASK-71
priority: medium
ordinal: 77000
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
- [x] #1 Critical runtime state domains mutate through explicit transition helpers
- [x] #2 State invariants are test-covered
- [x] #3 Direct ad-hoc mutation in migrated domains is removed
- [x] #4 Ownership/mutation rules documented
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Implementation plan (writing-plans):
1) Add pure reducer-style transition helpers in `src/main/state.ts` for critical AniList domains: client-secret state, retry queue metadata, media-guess runtime state, and tracking in-flight flag. Add focused tests in `src/main/state.test.ts`.
2) Rewire `src/main.ts` to use transition helpers instead of direct mutation for migrated domains in runtime deps wiring (`buildAnilistStateRuntimeMainDepsHandler`, media-guess state deps, retry-update deps, post-watch in-flight setter).
3) Add/extend invariant tests in `src/main/runtime/anilist-state.test.ts` and `src/main/runtime/anilist-media-state.test.ts` (metadata preservation, idempotent reset, scoped-field resets).
4) Document ownership/mutation rules in `docs/architecture.md` under composition guidance and run verification gates (`bun run build`, `bun run test:core:src`).
5) Record evidence and AC/DoD progress in TASK-79 notes.

Parallelization: run code-wiring slice and tests/docs slice in parallel subagents where safe, then reconcile and run final gates in the main session.

Detailed step-by-step plan saved at `docs/plans/2026-02-21-task-79-runtime-state-reducers.md`.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: Started implementation pass via writing-plans/executing-plans workflow. Scoping to explicit transition helpers/reducers for critical AniList runtime state domains in `src/main.ts`/`src/main/state.ts`, plus invariant tests and architecture ownership rules docs.

Implemented explicit AniList runtime transition helpers/reducers in `src/main/state.ts` and rewired migrated `src/main.ts` mutation paths through them (client-secret state, retry queue metadata, media-guess runtime state, in-flight flag).

Added focused reducer tests in `src/main/state.test.ts` plus invariants coverage updates in `src/main/runtime/anilist-state.test.ts` and `src/main/runtime/anilist-media-state.test.ts`.

Documented runtime ownership/mutation rules for migrated domains in `docs/architecture.md` under Composition Pattern.

Validation evidence: `bun test src/main/state.test.ts src/main/runtime/anilist-state.test.ts src/main/runtime/anilist-media-state.test.ts` (12 pass), `bun run build` (pass), `bun run test:core:src` (219 pass, 6 skip, 0 fail).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Introduced explicit reducer-style runtime transitions for critical AniList state in the main process by adding pure transition helpers in `src/main/state.ts` and routing migrated writes in `src/main.ts` through those helpers. Removed ad-hoc direct mutation in migrated domains (client-secret state, retry queue metadata, media-guess runtime state, in-flight flag), added focused and invariant tests, and documented ownership/mutation rules in architecture docs. Validation passed with focused state suites, full build, and core source test lane.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Core tests pass after migration
- [x] #2 No behavior regressions in startup/IPC/overlay flows
<!-- DOD:END -->
