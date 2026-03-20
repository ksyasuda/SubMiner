---
id: TASK-202
title: Use ended session media position for anime episode progress
status: Done
assignee:
  - Codex
created_date: '2026-03-19 14:55'
updated_date: '2026-03-19 17:36'
labels:
  - stats
  - ui
  - bug
milestone: m-1
dependencies: []
references:
  - stats/src/components/anime/EpisodeList.tsx
  - stats/src/types/stats.ts
  - src/core/services/immersion-tracker/session.ts
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker/storage.ts
priority: medium
ordinal: 105720
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The anime episode list currently computes the `Progress` column from cumulative `totalActiveMs / durationMs`, which can exceed the intended watch-position meaning after rewatches or repeated sessions. Persist the playback position at the time a session ends and drive episode progress from that stored stop position instead.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Session finalization persists the playback position reached when the session ended.
- [x] #2 Anime episode queries expose the most recent ended-session media position for each episode.
- [x] #3 Episode-list progress renders from ended media position instead of cumulative active watch time.
- [x] #4 Regression coverage locks storage/query/UI behavior for the new progress source.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add failing regression coverage for persisted ended media position and episode progress rendering.
2. Add `ended_media_ms` to the immersion-session schema and persist `lastMediaMs` when ending a session.
3. Thread the new field through episode queries/types and render episode progress from `endedMediaMs / durationMs`.
4. Run targeted verification plus typecheck, then record the outcome.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->

Added nullable `ended_media_ms` storage to immersion sessions, persisted `lastMediaMs` when sessions finalize, and exposed the most recent ended-session media position through anime episode queries/types. The anime episode list now renders `Progress` from `endedMediaMs / durationMs` instead of cumulative active watch time, so rewatches no longer inflate the displayed percentage.

Verification:

- `bun test src/core/services/immersion-tracker/storage-session.test.ts`
- `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
- `bun test stats/src/lib/yomitan-lookup.test.tsx stats/src/lib/stats-ui-navigation.test.tsx`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun x prettier --check 'src/core/services/immersion-tracker/types.ts' 'src/core/services/immersion-tracker/storage.ts' 'src/core/services/immersion-tracker/session.ts' 'src/core/services/immersion-tracker/query.ts' 'src/core/services/immersion-tracker/storage-session.test.ts' 'src/core/services/immersion-tracker/__tests__/query.test.ts' 'stats/src/types/stats.ts' 'stats/src/components/anime/EpisodeList.tsx' 'stats/src/lib/yomitan-lookup.test.tsx' 'stats/src/lib/stats-ui-navigation.test.tsx' 'backlog/tasks/task-202 - Use-ended-session-media-position-for-anime-episode-progress.md' 'changes/2026-03-19-stats-ended-media-progress.md'`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core 'src/core/services/immersion-tracker/types.ts' 'src/core/services/immersion-tracker/storage.ts' 'src/core/services/immersion-tracker/session.ts' 'src/core/services/immersion-tracker/query.ts' 'src/core/services/immersion-tracker/storage-session.test.ts' 'src/core/services/immersion-tracker/__tests__/query.test.ts' 'stats/src/types/stats.ts' 'stats/src/components/anime/EpisodeList.tsx' 'stats/src/lib/yomitan-lookup.test.tsx' 'stats/src/lib/stats-ui-navigation.test.tsx' 'backlog/tasks/task-202 - Use-ended-session-media-position-for-anime-episode-progress.md' 'changes/2026-03-19-stats-ended-media-progress.md'`
- Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260319-173511-AV7kUg/`

<!-- SECTION:OUTCOME:END -->
