---
id: TASK-13
title: Fix second-instance --start when texthooker-only instance is running
status: Done
assignee: []
created_date: '2026-02-11 23:47'
updated_date: '2026-02-18 04:11'
labels:
  - bugfix
  - cli
  - overlay
dependencies: []
priority: high
ordinal: 51000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When SubMiner is already running in texthooker-only mode, a subsequent `--start` command from a second instance is currently ignored. This can leave users without an initialized overlay runtime even though startup commands were issued. Adjust CLI command handling so `--start` on second-instance initializes overlay runtime when it is not yet initialized, while preserving current ignore behavior when overlay runtime is already active.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Second-instance `--start` initializes overlay runtime when current instance has deferred/not-initialized overlay runtime.
- [x] #2 Second-instance `--start` remains ignored (existing behavior) when overlay runtime is already initialized.
- [x] #3 CLI command service tests cover both behaviors and pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Patched CLI second-instance `--start` handling in `src/core/services/cli-command-service.ts` to initialize overlay runtime when deferred.

Added regression test for deferred-runtime start path and updated initialized-runtime second-instance tests in `src/core/services/cli-command-service.test.ts`.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Fixed overlay startup regression path where a second-instance `--start` could be ignored even when the primary instance was running in texthooker-only/deferred overlay mode.

Changes:

- Updated `handleCliCommandService` logic so `ignoreStart` applies only when source is second-instance, `--start` is present, and overlay runtime is already initialized.
- Added explicit overlay-runtime initialization path for second-instance `--start` when runtime is not initialized.
- Kept existing behavior for already-initialized runtime (still logs and ignores redundant `--start`).
- Added and updated tests in `cli-command-service.test.ts` to validate both deferred and initialized second-instance startup behaviors.

Validation:

- `pnpm run build` succeeded.
- `node dist/core/services/cli-command-service.test.js` passed (11/11).
<!-- SECTION:FINAL_SUMMARY:END -->
