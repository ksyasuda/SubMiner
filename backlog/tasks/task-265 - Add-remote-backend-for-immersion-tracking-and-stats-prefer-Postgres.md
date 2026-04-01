---
id: TASK-265
title: Add remote backend for immersion tracking and stats (prefer Postgres)
status: To Do
assignee: []
created_date: '2026-04-01 00:47'
labels: []
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker-service.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/storage.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/sqlite.ts
  - /home/sudacode/projects/japanese/SubMiner/src/stats-daemon-runner.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/stats-server.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/boot/services.ts
  - /home/sudacode/projects/japanese/SubMiner/package.json
documentation:
  - /home/sudacode/projects/japanese/SubMiner/docs/architecture/README.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/docs/architecture/stats-trends-data-flow.md
  - /home/sudacode/projects/japanese/SubMiner/README.md
  - /home/sudacode/projects/japanese/SubMiner/config.example.jsonc
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Enable immersion tracking/stats to use a remote authoritative backend so multiple devices can share the same history.

Current state: `ImmersionTrackerService` opens a local `immersion.sqlite` file from the app data/config path, `stats-daemon-runner` points at that same local file, and `config.example.jsonc` only exposes `immersionTracking.dbPath` for a local path override. The stats API/dashboard reads from the same tracker service and assumes the local database is the source of truth.

Goal: add a remote backend option that avoids shared filesystem/database-file syncing between devices. Do not use SSH/rsync/shared network filesystem as the primary sync strategy for live multi-device use.

Backend choice: prefer Postgres if it can be integrated without a broad new dependency surface or destabilizing the current runtime; otherwise use the least invasive remote backend that can be shipped with the current stack and document the tradeoff clearly. Preserve the current local SQLite mode as the default/offline fallback if possible.

This ticket should cover the full product/architecture change: configuration, storage access, stats reads, startup/error handling, migration/bootstrap from existing local data, tests, and docs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The app can be configured to use a remote authoritative backend for immersion tracking instead of only a local `immersion.sqlite` file.
- [ ] #2 The chosen backend persists tracker writes and serves the existing stats read models across app restarts.
- [ ] #3 Two devices can point at the same remote backend without relying on a shared filesystem or raw SQLite file sync.
- [ ] #4 Local SQLite remains supported as the default or fallback mode for offline use.
- [ ] #5 If the remote backend is unavailable or misconfigured, startup/write paths fail with actionable errors instead of silent data loss.
- [ ] #6 A migration or bootstrap path exists to move existing local immersion data into the remote backend or seed a new device from it.
- [ ] #7 Config/examples/docs explain the backend choice, required connection/setup details, and any security/network assumptions.
- [ ] #8 Tests cover backend selection plus at least one representative write/read path against the remote backend.
<!-- AC:END -->
