---
id: TASK-211
title: Recover anime episode progress from subtitle timing when checkpoints are missing
status: Done
assignee:
  - '@Codex'
created_date: '2026-03-20 10:15'
updated_date: '2026-03-20 10:22'
labels:
  - stats
  - bug
milestone: m-1
dependencies: []
references:
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker/__tests__/query.test.ts
---

## Description

Anime episode progress can still show `0%` for older sessions that have watch-time and subtitle timing but no persisted `ended_media_ms` checkpoint. Recover progress from the latest retained subtitle/event segment end so already-recorded sessions render a useful progress percentage.

## Acceptance Criteria

- [x] `getAnimeEpisodes` returns the latest known session position even when `ended_media_ms` is null but subtitle/event timing exists.
- [x] Existing ended-session metrics and aggregation totals do not regress.
- [x] Regression coverage locks the fallback behavior.

## Implementation Notes

Added a query-side fallback for anime episode progress: when the newest session for a video has no persisted `ended_media_ms`, `getAnimeEpisodes` now uses the latest retained subtitle-line or session-event `segment_end_ms` from that same session. This recovers useful progress for already-recorded sessions that have timing data but predate or missed checkpoint persistence.

Verification: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts` passed. `bun run typecheck` passed.
