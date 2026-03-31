---
id: TASK-261
title: Fix immersion tracker SQLite timestamp truncation
status: In Progress
assignee: []
created_date: '2026-03-31 01:45'
labels:
  - immersion-tracker
  - sqlite
  - bug
dependencies: []
references:
  - src/core/services/immersion-tracker
priority: medium
ordinal: 1200
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current-epoch millisecond values are being truncated by the libsql driver when bound as numeric parameters, which corrupts session, telemetry, lifetime, and rollup timestamps.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Current-epoch millisecond timestamps persist correctly in session, telemetry, lifetime, and rollup tables
- [ ] #2 Startup backfill and destroy/finalize flows keep retained sessions and lifetime summaries consistent
- [ ] #3 Regression tests cover the destroyed-session, startup backfill, and distinct-day/distinct-video lifetime semantics
<!-- AC:END -->
