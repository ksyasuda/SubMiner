---
id: TASK-87
title: Fix plugin --start flow so tokenization/frequency/N+1 initialize correctly
status: Done
assignee: []
created_date: '2026-02-19 19:05'
updated_date: '2026-02-19 23:18'
labels: []
dependencies: []
priority: medium
ordinal: 69000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Plugin startup flow (`--texthooker` then `--start`) can miss tokenization initialization, which makes frequency highlighting and N+1 appear broken even when dictionaries/config are correct. Ensure `--start` initializes overlay runtime when needed, including second-instance handoff mode.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `--start` initializes overlay runtime when runtime is not yet initialized.
- [x] #2 `second-instance --start` is ignored only when overlay runtime is already initialized.
- [x] #3 Regression tests cover both `--start` behaviors and pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: subtitle tokenization path in `src/main.ts` only runs when overlay windows exist; command runtime did not initialize overlay runtime for `--start`, especially visible in plugin handoff (`--texthooker` -> `--start`). Updated `src/core/services/cli-command.ts` to initialize overlay runtime for `--start`, and narrowed second-instance ignore behavior to already-initialized runtime only. Added regression tests in `src/core/services/cli-command.test.ts`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed plugin/start regression by ensuring `--start` initializes overlay runtime when needed and no longer gets dropped in second-instance handoff before runtime init. Added/updated CLI command tests; targeted suite passes.
<!-- SECTION:FINAL_SUMMARY:END -->
