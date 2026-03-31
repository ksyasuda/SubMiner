---
id: TASK-261
title: Fix immersion tracker SQLite timestamp truncation
status: Done
assignee: []
created_date: '2026-03-31 01:45'
updated_date: '2026-03-31 19:37'
labels:
  - immersion-tracker
  - sqlite
  - bug
dependencies: []
references:
  - src/core/services/immersion-tracker
priority: medium
ordinal: 179500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current-epoch millisecond values are being truncated by the libsql driver when bound as numeric parameters, which corrupts session, telemetry, lifetime, and rollup timestamps.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Current-epoch millisecond timestamps persist correctly in session, telemetry, lifetime, and rollup tables
- [x] #2 Startup backfill and destroy/finalize flows keep retained sessions and lifetime summaries consistent
- [x] #3 Regression tests cover the destroyed-session, startup backfill, and distinct-day/distinct-video lifetime semantics
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-31 assessment: epoch-ms timestamp writes now route through `toDbMs()` / `toDbTimestamp()` in `src/core/services/immersion-tracker/query-shared.ts`, which avoids libsql numeric-parameter truncation by binding BigInt/string values before they hit SQLite. The fix is wired through the session, storage/telemetry, lifetime, and rollup-maintenance paths in `src/core/services/immersion-tracker/session.ts`, `src/core/services/immersion-tracker/storage.ts`, `src/core/services/immersion-tracker/lifetime.ts`, and `src/core/services/immersion-tracker/maintenance.ts`.

Acceptance coverage is present: `bun test src/core/services/immersion-tracker-service.test.ts` passed with explicit regressions for destroy/finalize persistence, startup backfill when retained sessions exist but lifetime tables are empty, startup reconciliation of stale active sessions, `rebuildLifetimeSummaries`, and distinct-day / distinct-video lifetime semantics. `bun test src/core/services/immersion-tracker/time.test.ts src/core/services/immersion-tracker/maintenance.test.ts` also passed.

Remaining action item before close: fix the two `src/main/runtime/stats-cli-command.test.ts` cleanup-lifetime assertions that currently use Bun-misparsed underscored millisecond literals (`1_710_000_000_000` evaluates to `-2147483648` under Bun 1.3.11), rerun that verification lane, then write the final summary and mark the task Done.
<!-- SECTION:NOTES:END -->
