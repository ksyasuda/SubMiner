---
id: TASK-88
title: >-
  Fix second-instance --start handling when overlay runtime is already
  initialized
status: Done
assignee:
  - codex
created_date: '2026-03-06 07:30'
updated_date: '2026-03-06 07:31'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/cli-command.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/cli-command.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Restore the CLI command guard so a second-instance `--start` request does not reconnect or reinitialize overlay work when the overlay runtime is already active, while preserving other second-instance commands.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Second-instance `--start` logs that the app is already running when the overlay runtime is initialized.
- [x] #2 Second-instance `--start` does not reconnect the MPV client when the overlay runtime is already initialized.
- [x] #3 Second-instance commands that include non-start actions still execute those actions.
- [x] #4 Regression coverage documents the guarded second-instance behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the failing `handleCliCommand` second-instance `--start` regression in `src/core/services/cli-command.test.ts`.
2. Update `src/core/services/cli-command.ts` so second-instance `--start` is ignored when the overlay runtime is already initialized, while still allowing non-start actions in the same invocation.
3. Run focused CLI command tests, then rerun the core test target if practical, and record acceptance criteria/results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reproduced the failing second-instance `--start` regression in `src/core/services/cli-command.test.ts` before editing.

Restored a guard in `src/core/services/cli-command.ts` that ignores second-instance `--start` when the overlay runtime is already initialized, but still allows other flags in the same invocation to run.

Verification: `bun test src/core/services/cli-command.test.ts`, `bun run test:core:src`, and `bun run test` all pass; the six immersion tracker tests remain skipped as before.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored the missing second-instance `--start` guard in `src/core/services/cli-command.ts`.

- Added an `ignoreSecondInstanceStart` check so `handleCliCommand` logs `Ignoring --start because SubMiner is already running.` when a second-instance `--start` arrives after the overlay runtime is already initialized.
- Updated start gating so pure duplicate `--start` requests no longer reconnect the MPV client, while combined commands such as `--start --toggle-visible-overlay` still execute their non-start behavior and can connect through those paths.
- Verified the regression with existing CLI command coverage and reran the broader test targets.

Tests run:
- `bun test src/core/services/cli-command.test.ts`
- `bun run test:core:src`
- `bun run test`

Notes:
- The six skipped immersion tracker tests are unchanged from the pre-fix baseline.
<!-- SECTION:FINAL_SUMMARY:END -->
