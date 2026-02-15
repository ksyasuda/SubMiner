---
id: TASK-32
title: Add pluggable immersion tracking persistence layer for non-SQLite backends
status: To Do
assignee: []
created_date: '2026-02-13 19:30'
updated_date: '2026-02-14 00:46'
labels:
  - analytics
  - backend
  - database
  - architecture
  - immersion
dependencies:
  - TASK-28
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
## Scope
Create a storage abstraction for immersion analytics so existing SQLite-first tracking can evolve to external backends (PostgreSQL/MySQL/other) without rewriting core analytics behavior.

## Backend portability and performance contract
- Define a canonical interface that avoids SQL dialect leakage and returns stable query result shapes for session, video, and trend analytics.
- Decouple all `TASK-28` persistence callers from raw SQLite access behind a repository/adapter boundary.
- Keep domain model and query DTOs backend-agnostic; SQL specifics live in adapters.

## Target architecture
- Add provider config: `storage.provider` default `sqlite`.
- Use dependency injection/service factory to select adapter at startup.
- Add adapter contract tests to validate `Session`, `Video`, `Telemetry`, `Event`, and `Rollup` operations behave identically across backends.
- Include migration and rollout contract with compatibility/rollback notes.

## Numeric operational constraints to preserve (from TASK-28)
- All adapters must support write batching (or equivalent) with flush cadence equivalent to 25 records or 500ms.
- All adapters must enforce bounded in-memory write queue (default 1000 rows) and explicit overflow policy.
- All adapters should preserve index/query shapes needed for ~150ms p95 read targets on session/video/time-window queries at ~1M events.
- All adapters should support retention semantics equivalent to: events 7d, telemetry 30d, daily rollups 365d, monthly rollups 5y, and prune-on-startup + every 24h + idle vacuum schedule.

## Async and isolation constraint
- Adapter API must support enqueue/write-queue semantics so tokenization/render loops never block on persistence.
- Background worker owns DB I/O; storage adapter exposes non-blocking API surface to tracking pipeline.

Acceptance Criteria:
--------------------------------------------------
- [ ] #1 Define a stable `ImmersionTrackingStore` interface covering session lifecycle, telemetry counters, event writes, and analytics queries (timeline, video stats, throughput, and streak/trend summaries).
- [ ] #2 Implement a DI-bound storage repository that encapsulates all DB interaction behind the interface.
- [ ] #3 Refactor `TASK-28`-related data writes/reads in the tracking pipeline to depend on the abstraction (not raw SQLite calls).
- [ ] #4 Document migration path from current schema to backend-adapter-based persistence.
- [ ] #5 Add config for storage provider (`sqlite` default) and connection options, with validation and clear fallback behavior.
- [ ] #6 Create adapter contracts for at least one external SQL engine target (PostgreSQL/MySQL contract/schema mapping), even if execution is deferred with a follow-up task.
- [ ] #7 Add tests or validation plan for adapter boundary behavior (mock adapter + contract tests, and error handling behavior).
- [ ] #8 Expose retention and write profile defaults in backend contracts: 25 events/500ms batching, queue cap 1000, event payload cap 256B, overflow policy, and retention windows equivalent to TASK-28.
- [ ] #9 Preserve performance contract semantics in adapters: query/index assumptions for sub-150ms p95 local reads on ~1M event scale and same read-path shapes as TASK-28.

Definition of Done:
--------------------------------------------------
- [ ] #1 Storage interface with required method signatures and query contracts is documented in code and backlog docs.
- [ ] #2 Default SQLite adapter remains the primary implementation and passes existing/ planned immersion tracking expectations.
- [ ] #3 Non-SQLite implementation path is explicitly represented in config and adapter scaffolding.
- [ ] #4 Tracking pipeline is fully storage-engine agnostic and can support a new adapter without schema churn.
- [ ] #5 Migration and rollout strategy is documented (phased migration, backfill, compatibility behavior, rollback plan).
- [ ] #6 Backend contract includes operational and storage-growth semantics that preserve performance/size behavior from TASK-28.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Define a stable `ImmersionTrackingStore` interface covering session lifecycle, telemetry counters, event writes, and analytics queries (timeline, video stats, throughput, and streak/trend summaries).
- [ ] #2 Implement a DI-bound storage repository that encapsulates all DB interaction behind the interface.
- [ ] #3 Refactor `TASK-28`-related data writes/reads in the tracking pipeline to depend on the abstraction (not raw SQLite calls).
- [ ] #4 Document migration path from current schema to backend-adapter-based persistence.
- [ ] #5 Add config for storage provider (`sqlite` default) and connection options, with validation and clear fallback behavior.
- [ ] #6 Create adapter contracts for at least one external SQL engine target (PostgreSQL/MySQL contract/schema mapping), even if execution is deferred with a follow-up task.
- [ ] #7 Add tests or validation plan for adapter boundary behavior (mock adapter + contract tests, and error handling behavior).
- [ ] #8 Expose retention and write profile defaults in backend contracts: 25 events/500ms batching, queue cap 1000, event payload cap 256B, overflow policy, and retention windows equivalent to TASK-28.
- [ ] #9 Preserve performance contract semantics in adapters: query/index assumptions for sub-150ms p95 local reads on ~1M event scale and same read-path shapes as TASK-28.
- [ ] #10 #9 Storage interface must expose async/fire-and-forget write contract for telemetry/event ingestion (no blocking calls available to UI/tokenization path).
- [ ] #11 #10 Adapter boundary must guarantee: persistence errors from tracker are observed internally and surfaced as background diagnostics, never bubbled into tokenization/render execution path.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Priority deferred from medium to low: this is premature until TASK-28 (SQLite tracking) ships and a concrete second backend need emerges. The SQLite-first design in TASK-28 already accounts for future portability via schema contracts and adapter boundaries. Revisit after TASK-28 has been in use.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Storage interface with required method signatures and query contracts is documented in code and backlog docs.
- [ ] #2 Default SQLite adapter remains the primary implementation and passes existing/ planned immersion tracking expectations.
- [ ] #3 Non-SQLite implementation path is explicitly represented in config and adapter scaffolding.
- [ ] #4 Tracking pipeline is fully storage-engine agnostic and can support a new adapter without schema churn.
- [ ] #5 Migration and rollout strategy is documented (phased migration, backfill, compatibility behavior, rollback plan).
- [ ] #6 Backend contract includes operational and storage-growth semantics that preserve performance/size behavior from TASK-28.
<!-- DOD:END -->
