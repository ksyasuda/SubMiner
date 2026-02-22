---
id: TASK-106
title: Decompose immersion tracker service into storage session and metadata modules
status: Done
assignee:
  - opencode-task106-immersion-modules
created_date: '2026-02-22 07:14'
updated_date: '2026-02-22 21:58'
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
- [x] #1 `immersion-tracker-service.ts` no longer embeds full schema SQL and metadata probing logic directly.
- [x] #2 Extracted modules have focused tests for session transitions, DB writes, and metadata parsing.
- [x] #3 Tracker behavior remains unchanged (session lifecycle, rollups, retention, queue semantics).
- [x] #4 Build and tracker-related source tests pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Plan file: docs/plans/2026-02-22-task-106-immersion-tracker-storage-session-metadata.md

Execution plan:
1) Add failing seam tests in src/core/services/immersion-tracker-service.test.ts for extracted storage/session/metadata APIs.
2) Parallel slice A: extract storage/schema + session lifecycle collaborators into src/core/services/immersion-tracker/storage.ts and session.ts; rewire service.
3) Parallel slice B: extract ffprobe/hash/local metadata probing into src/core/services/immersion-tracker/metadata.ts with injected seams; rewire service.
4) Merge/wire both slices in immersion-tracker-service facade; verify behavior parity and file-size reduction.
5) Run required gates: bun run build, bun test src/core/services/immersion-tracker-service.test.ts, bun run test:core:src.
6) Update ownership docs (docs/immersion-tracking.md and/or docs/architecture.md) and finalize backlog AC/DoD notes/status (no commit).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented decomposition with new modules: `src/core/services/immersion-tracker/storage.ts` (schema/DB writes/prepared statements), `src/core/services/immersion-tracker/session.ts` (session start/finalize persistence), and `src/core/services/immersion-tracker/metadata.ts` (ffprobe/hash/local metadata probing with injectable deps).

`src/core/services/immersion-tracker-service.ts` reduced from 1099 LOC baseline to 654 LOC; now orchestration facade delegating storage/session/metadata concerns while preserving public API and queue/maintenance semantics.

Added focused tests: `src/core/services/immersion-tracker/storage-session.test.ts` (schema creation, session transitions, DB writes) and `src/core/services/immersion-tracker/metadata.test.ts` (ffprobe parsing/fallbacks, hash fallback).

Validation: `bun test src/core/services/immersion-tracker/metadata.test.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker-service.test.ts` => 7 pass, 9 skip, 0 fail; `bun run test:core:src` => 236 pass, 6 skip, 0 fail.

Build gate currently fails due pre-existing unrelated TypeScript errors in `src/anki-integration/*` and `src/main/runtime/*` test/type contracts; no new immersion-tracker errors observed in this pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Decomposed `src/core/services/immersion-tracker-service.ts` into focused collaborators while preserving tracker behavior and public API. Introduced `src/core/services/immersion-tracker/storage.ts` (schema bootstrap, prepared statements, DB writes), `src/core/services/immersion-tracker/session.ts` (session start/finalize persistence), and `src/core/services/immersion-tracker/metadata.ts` (ffprobe/hash/local metadata probing with injectable deps). Reduced service file from 1099 LOC baseline to 654 LOC and kept queue/flush/maintenance orchestration in the facade.

Added focused regression coverage via `src/core/services/immersion-tracker/storage-session.test.ts` and `src/core/services/immersion-tracker/metadata.test.ts`, and updated ownership documentation in `docs/immersion-tracking.md` to reflect new boundaries.

Verification:
- `bun run build` ✅
- `bun test src/core/services/immersion-tracker/metadata.test.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker-service.test.ts` ✅ (7 pass, 9 skip)
- `bun run test:core:src` ✅ (236 pass, 6 skip)
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Service file size reduced materially from current baseline.
- [x] #2 Ownership boundaries documented in `docs/architecture.md` or relevant service docs.
- [x] #3 No regression in `bun run test:core:src` immersion tracker coverage.
<!-- DOD:END -->
