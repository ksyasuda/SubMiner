---
id: TASK-98
title: Shift core tests to source level and trim dist coupling
status: In Progress
assignee:
  - opencode
created_date: '2026-02-21 07:15'
updated_date: '2026-02-21 09:56'
labels:
  - testing
  - maintainability
  - developer-experience
dependencies:
  - TASK-95
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Core test flow is heavily coupled to `dist/` artifacts, increasing cycle time and reducing test clarity. Move the majority of tests to source-level execution while retaining minimal dist smoke coverage.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Catalog current test commands and classify tests: source-level candidates vs required dist smoke tests.
2. Introduce/standardize source-level test runners for core suites.
3. Migrate high-volume `dist/core/services/*` tests to source-level equivalents.
4. Keep a small dist smoke suite for packaging/runtime sanity only.
5. Update `package.json` scripts and CI workflow steps to use new split.
6. Capture timing comparison (before/after) for local and CI runs.
7. Run verification gate for both lanes: source-level tests + dist smoke tests.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Majority of core regression tests run from source-level entrypoints.
- [x] #2 Dist suite reduced to explicit smoke scope with documented purpose.
- [x] #3 CI still validates packaged/runtime assumptions via smoke lane.
- [x] #4 Total default test cycle time improves measurably.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Baseline current dist-coupled matrix and capture `time bun run test:fast` before timing.
2) Update `package.json` scripts so default regression lanes (`test:config`, `test:core`, `test:fast`) run source-level tests (`bun test src/...`) instead of `dist` artifacts.
3) Introduce explicit `test:smoke:dist` command for minimal compiled/runtime smoke checks; keep launcher smoke behavior intact.
4) Update `.github/workflows/ci.yml` and `.github/workflows/release.yml` to run source lane as primary + dist smoke lane as secondary.
5) Document command matrix in `docs/development.md`, rerun timing commands, capture before/after delta in task notes, then finalize AC/DoD.

Detailed execution plan: `docs/plans/2026-02-21-task-98-source-tests-dist-smoke-split.md`
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Execution updates (2026-02-21):
- Default regression lane moved to source entrypoints in `package.json`:
  - `test:config` -> `test:config:src` (`bun test src/config/config.test.ts src/config/path-resolution.test.ts`)
  - `test:core` -> `test:core:src` (`bun test` over curated core list)
  - `test:fast` -> source config + source core
- Dist coupling trimmed to explicit smoke lane:
  - `test:smoke:dist` -> `test:config:smoke:dist` + `test:core:smoke:dist`
  - Smoke scope validates compiled/runtime assumptions across config, IPC/electron interop (`overlay-manager`, `anilist-token-store`), startup, renderer, main URL guard, and x11 tracker.
- CI/release wiring updated:
  - `.github/workflows/ci.yml`: `test:fast` (source) before build; `test:smoke:dist` after build.
  - `.github/workflows/release.yml` quality-gate: source tests, build, dist smoke.
- Docs updated in `docs/development.md` with source-vs-dist command matrix and smoke rationale.
- Timing evidence (repro commands):
  - Before: `time bun run test:fast` (dist-coupled) => `1.438 total`
  - After: `time bun run test:fast` (source lane) => `1.017 total`
  - Delta: ~`0.421s` faster (~29.3% reduction).
  - Dist smoke timing: `time bun run test:smoke:dist` => `0.206 total`.
- Verification:
  - PASS: `bun run test:config:src && bun run test:core:src`
  - PASS: `bun run test:smoke:dist`
  - BLOCKED (pre-existing unrelated workspace errors): `bun run build` currently fails in `src/main.ts` and `src/main/runtime/composers/mpv-runtime-composer.test.ts` from in-flight TASK-96/97 changes present in working tree; not introduced by TASK-98 edits.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Test command matrix documented in task notes.
- [ ] #2 CI config updated and passing with new source/dist split.
- [x] #3 Performance delta captured with reproducible timing commands.
<!-- DOD:END -->
