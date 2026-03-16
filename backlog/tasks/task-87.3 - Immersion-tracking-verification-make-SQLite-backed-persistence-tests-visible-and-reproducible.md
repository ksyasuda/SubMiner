---
id: TASK-87.3
title: >-
  Immersion tracking verification: make SQLite-backed persistence tests visible
  and reproducible
status: Done
assignee:
  - Kyle Yasuda
created_date: '2026-03-06 03:19'
updated_date: '2026-03-16 05:13'
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
ordinal: 90500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The immersion tracker is persistence-heavy, but its SQLite-backed tests are conditionally skipped in the standard Bun run when node:sqlite support is unavailable. That creates a blind spot around session finalization, telemetry persistence, and retention behavior. This task should establish a reliable automated verification path for the database-backed cases and make the prerequisite/runtime behavior explicit to contributors and CI.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Database-backed immersion tracking tests run in at least one documented automated command that is practical for contributors or CI to execute.
- [x] #2 If the current runtime cannot execute the SQLite-backed tests, the repository exposes that limitation clearly instead of silently reporting a misleading green result.
- [x] #3 Contributor-facing documentation explains how to run the immersion tracker verification lane and any environment prerequisites it depends on.
- [x] #4 The resulting verification covers session persistence or finalization behavior that is not exercised by the pure seam tests alone.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Implementation plan recorded in `docs/plans/2026-03-06-immersion-sqlite-verification.md`.

1. Update `src/core/services/immersion-tracker-service.test.ts` and `src/core/services/immersion-tracker/storage-session.test.ts` so unsupported `node:sqlite` runtimes emit an explicit skip reason instead of a silent top-level skip alias.
2. Add a dedicated `package.json` SQLite verification lane that runs both immersion persistence suites together under a runtime with `node:sqlite` support, likely via built `dist/**` tests executed by Node.
3. Wire that lane into `.github/workflows/ci.yml` and `.github/workflows/release.yml` so automated verification includes a real DB-backed persistence/finalization check.
4. Document the new command, prerequisites, and coverage in `README.md`, including the distinction between Bun's default lane and the reproducible SQLite lane.
5. Validate the final lane by running the dedicated command and confirming it exercises persistence/finalization behavior beyond the seam-only tests.

Execution adjustment: the reproducible lane uses `node --experimental-sqlite --test ...` because Node 22 exposes `node:sqlite` behind the experimental flag. Running that lane also exposed placeholder-count mismatches in `src/core/services/immersion-tracker/storage.ts`, so the final implementation includes a small SQL placeholder fix required for the new cross-runtime verification path.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Confirmed Bun 1.3.5 lacks `node:test` `t.skip()` support, so explicit unsupported-runtime messaging is surfaced with file-level warnings while the SQLite-backed tests remain conditionally skipped.

Added `test:immersion:sqlite:src`, `test:immersion:sqlite:dist`, and `test:immersion:sqlite` scripts; the source lane now prints explicit warnings when `node:sqlite` is unavailable, and the dist lane runs both SQLite-backed immersion suites under Node with `--experimental-sqlite`.

Wired the dist SQLite lane into `.github/workflows/ci.yml` and `.github/workflows/release.yml` after the bundle build, with explicit `actions/setup-node@v4` provisioning for Node 22.12.0.

Fixed SQL prepared-statement placeholder counts in `src/core/services/immersion-tracker/storage.ts`, which the new Node-backed SQLite lane surfaced immediately.

Verification: `bun run test:immersion:sqlite:src` -> pass with explicit unsupported-runtime warnings and 10 skips under Bun 1.3.5; `bun run test:immersion:sqlite` -> pass with 14/14 tests under Node 22.12.0 + `--experimental-sqlite`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added an explicit SQLite-backed immersion verification lane and documented it so persistence-heavy coverage is no longer hidden behind Bun-only skips. `package.json` now exposes source and dist SQLite scripts, the source test files print actionable warnings when `node:sqlite` is unavailable, and `README.md` explains the dedicated contributor command plus its Node 22 `--experimental-sqlite` prerequisite.

Automated verification now includes the new dist lane in both `.github/workflows/ci.yml` and `.github/workflows/release.yml` after build output is available. While wiring the reproducible Node lane, it exposed placeholder-count mismatches in `src/core/services/immersion-tracker/storage.ts`; fixing those placeholders makes the SQLite-backed persistence/finalization tests pass cross-runtime, covering session finalization, telemetry persistence, and storage-session write paths.
<!-- SECTION:FINAL_SUMMARY:END -->
