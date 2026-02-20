---
id: TASK-68
title: Allow trailing commas in JSONC config parsing
status: Done
assignee: []
created_date: '2026-02-18 10:13'
updated_date: '2026-02-19 23:18'
labels:
  - config
  - jsonc
dependencies: []
priority: medium
ordinal: 67000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Permit trailing commas in `config.jsonc` parsing so normal JSONC edits do not fail strict reload/watcher startup.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Config parser accepts trailing commas in JSONC objects/arrays
- [x] #2 Invalid malformed JSONC still fails strict reload
- [x] #3 Coverage added to prevent regression
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Enabled `allowTrailingComma: true` in `src/config/service.ts` JSONC parse options while preserving strict parse error handling.

Added regression test `accepts trailing commas in jsonc` in `src/config/config.test.ts`.

Updated docs note in `docs/configuration.md` that JSONC config supports comments and trailing commas.

Validated with `bun run build && bun run test:config:dist`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
SubMiner now accepts trailing commas in `config.jsonc` by enabling JSONC parser trailing-comma support in the strict config load path. Strict reload behavior remains intact for malformed JSON/JSONC, and a regression test now covers trailing-comma acceptance. Documentation was updated accordingly.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Config test suite passes after parser change
<!-- DOD:END -->
