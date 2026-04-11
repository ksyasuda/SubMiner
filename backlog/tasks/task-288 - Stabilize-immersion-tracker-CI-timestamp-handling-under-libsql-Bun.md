---
id: TASK-288
title: Stabilize immersion-tracker CI timestamp handling under libsql/Bun
status: Done
assignee: []
created_date: '2026-04-11 21:34'
updated_date: '2026-04-11 21:43'
labels:
  - bug
  - ci
  - immersion-tracker
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`bun run test:fast` is currently failing because large millisecond timestamps are not handled safely through the libsql/Bun path. Fix timestamp parsing/storage so lifetime/library and session-event queries return correct wall-clock values in CI and runtime.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Large wall-clock timestamps round-trip correctly through immersion-tracker lifetime/library queries under the repo's Bun/libsql runtime.
- [x] #2 Session-event timestamps round-trip correctly for real wall-clock values used by runtime event inserts.
- [x] #3 Targeted immersion-tracker regression coverage passes, and the previously failing `test:fast` lane no longer fails on these timestamp assertions.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause split in two places: Bun/libsql corrupts large millisecond timestamp strings when coerced through `Number(...)`, and `imm_session_events.ts_ms` being `INTEGER` let runtime event inserts/readbacks return `-2147483648` on CI/runtime.

Fix shipped by parsing timestamp strings without the broken `Number(largeString)` path, migrating `imm_session_events.ts_ms` to `TEXT`, ordering/retention queries via `CAST(ts_ms AS REAL)`, and avoiding `Number(currentMs)` when reusing already-normalized timestamp strings.

Added regression coverage for both real runtime event inserts and schema migration/repair of previously truncated session-event rows.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed immersion-tracker timestamp handling under Bun/libsql so large wall-clock millisecond values survive runtime writes, query reads, and schema migration. `bun run test:fast`, `bun run typecheck`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`, and `bun run changelog:lint` all pass after the patch.
<!-- SECTION:FINAL_SUMMARY:END -->
