---
id: TASK-87.1
title: >-
  Testing workflow: make standard test commands reflect the maintained test
  surface
status: Done
assignee:
  - OpenCode
created_date: '2026-03-06 03:19'
updated_date: '2026-03-16 05:13'
labels:
  - tests
  - maintainability
milestone: m-0
dependencies: []
references:
  - package.json
  - src/main-entry-runtime.test.ts
  - src/anki-integration/anki-connect-proxy.test.ts
  - src/main/runtime/jellyfin-remote-playback.test.ts
  - src/main/runtime/registry.test.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
ordinal: 86500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The current package scripts hand-enumerate a small subset of test files, which leaves the standard green signal misleading. A local audit found 241 test/type-test files under src/ and launcher/, but only 53 unique files referenced by the standard package.json test scripts. This task should redesign the runnable test matrix so maintained tests are either executed by the standard commands or intentionally excluded through a documented rule, instead of silently drifting out of coverage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The repository has a documented and reproducible test matrix for standard development commands, including which suites belong in the default lane versus slower or environment-specific lanes.
- [x] #2 The standard test entrypoints stop relying on a brittle hand-maintained allowlist for the currently covered unit and integration suites, or an explicit documented mechanism exists that prevents silent omission of new tests.
- [x] #3 Representative tests that were previously outside the standard lane from src/main/runtime, src/anki-integration, and entry/runtime surfaces are executed by an automated command and included in the documented matrix.
- [x] #4 Documentation for contributors explains which command to run for fast verification, full verification, and environment-specific verification.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update `package.json` to replace the current file-by-file test allowlists with a documented lane matrix: keep `test`/`test:fast` as the quick default lane, add `test:full` for the maintained source test surface, and add `test:env` for slower or environment-specific checks.
2. Use directory-based discovery for maintained suites so new tests under stable surfaces such as `src/main`, `src/anki-integration`, and `launcher` are not silently omitted by default script maintenance.
3. Split environment-specific verification into explicit commands for checks such as launcher smoke/plugin coverage and sqlite-gated tests, instead of leaving them undocumented or mixed into the default signal.
4. Update `README.md` with a contributor-facing testing matrix that explains fast, full, and environment-specific verification, including any prerequisites or expected skip behavior.
5. Verify the matrix by running representative targeted tests plus the documented lane commands, demonstrating that previously omitted entry/runtime, Anki integration, and main runtime tests now belong to automated commands.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reviewed task context via Backlog MCP plus repo audit. Current package.json test scripts still rely on hand-maintained file allowlists and omit large maintained areas including src/main/runtime, src/anki-integration, and src/main-entry-runtime.test.ts. Preparing an implementation plan and contributor-facing test matrix update before code changes.

Saved detailed implementation plan to docs/plans/2026-03-06-testing-workflow-test-matrix.md and recorded the approved direction in the Backlog task before implementation.

Implemented a lane-based test matrix. Added `scripts/run-test-lane.mjs` so Bun-managed `src/**` and launcher unit lanes discover files automatically while excluding a small explicit Node-only set instead of relying on large hand-maintained allowlists. Added `test:node:compat` for `ipc`, `anki-jimaku-ipc`, `overlay-manager`, `config-validation`, `startup-config`, and `registry` suites, kept `test:env` for launcher smoke/plugin plus SQLite-backed immersion checks, and updated `README.md` with the contributor-facing matrix and exclusions.

Validated the new matrix with `bun run test:fast`, `bun run test:full`, `bun run test:env`, `bun run test:src`, `bun run test:launcher:unit:src`, `bun run test:node:compat`, and targeted `bun test src/core/services/anilist/anilist-updater.test.ts`. Representative previously omitted surfaces now run through automated commands: `src/main-entry-runtime.test.ts` via `test:fast`, `src/anki-integration/anki-connect-proxy.test.ts` via `test:fast`/`test:src`, and `src/main/runtime/registry.test.ts` via `test:node:compat`/`test:full`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Reworked the repository test matrix so standard commands reflect the maintained test surface without relying on brittle file allowlists. Added automated Bun discovery lanes for Bun-compatible `src/**` and launcher unit suites, a documented Node compatibility lane for Electron/sqlite-sensitive tests, and updated the contributor docs with fast/full/environment-specific guidance plus explicit exclusions. Verified with `bun run test:fast`, `bun run test:full`, and `bun run test:env`, along with the component lanes and targeted regression coverage for the updated AniList guessit test seam.
<!-- SECTION:FINAL_SUMMARY:END -->
