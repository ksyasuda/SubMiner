---
id: TASK-29.2
title: Implement AniList retry/backoff queue for failed post-watch updates
status: Done
assignee: []
created_date: '2026-02-17 04:13'
updated_date: '2026-02-18 04:11'
labels:
  - anilist
  - reliability
  - queue
dependencies: []
parent_task_id: TASK-29
priority: medium
ordinal: 19000
---

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Failed AniList mutations are enqueued with retry metadata and exponential backoff.
- [ ] #2 Transient API/network failures retry automatically without blocking playback.
- [ ] #3 Queue is idempotent per media+episode update key and survives app restarts.
- [ ] #4 Permanent failures surface clear diagnostics and dead-letter state.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Implemented persistent AniList retry queue with exponential backoff, dead-lettering after max attempts, queue snapshot state wiring, and retry processing integrated into playback-triggered AniList update flow.

<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [ ] #1 Queue service integrated into AniList post-watch update path.
- [ ] #2 Backoff/retry behavior covered by unit tests.
<!-- DOD:END -->
