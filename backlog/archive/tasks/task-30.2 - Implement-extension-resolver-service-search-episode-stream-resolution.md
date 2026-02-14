---
id: TASK-30.2
title: Implement extension resolver service (search + episode + stream resolution)
status: To Do
assignee: []
created_date: '2026-02-13 18:32'
updated_date: '2026-02-13 18:34'
labels: []
dependencies:
  - TASK-30.1
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Build a dedicated service in main process that queries configured extension repos and normalizes results into a unified internal model, including optional playback metadata. Keep transport abstracted so future backends (local process, remote API, Manatán-compatible source) can be swapped without changing renderer contracts.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Create a typed internal model for source, series, episode, and playable candidate with fields for quality/audio/headers/referrer/userAgent.
- [ ] #2 Implement provider abstraction with pluggable fetch/execution strategy from config.
- [ ] #3 Add services for searchAnime, listEpisodes, resolveStream (or equivalent) with cancellation/error boundaries.
- [ ] #4 Normalize all provider responses into deterministic field names and stable IDs.
- [ ] #5 Include resilient handling for empty/no-result/no-URL cases and network faults with explicit error categories.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Phase 2 — Core service: provider integration and stream resolution
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Resolver never leaks raw provider payload to renderer.
- [ ] #2 Streaming URL output includes reason for failure when unavailable.
- [ ] #3 Service boundaries allow unit-level validation of request/response mapping logic.
- [ ] #4 No blocking calls on Electron UI/main thread; all I/O is async and cancellable.
<!-- DOD:END -->
