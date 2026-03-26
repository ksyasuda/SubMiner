---
id: TASK-177
title: Track intentional Yomitan lookups in stats
status: Done
assignee:
  - codex
created_date: '2026-03-17 09:15'
updated_date: '2026-03-18 05:28'
labels:
  - stats
  - immersion-tracking
  - yomitan
milestone: m-1
dependencies: []
references:
  - vendor/subminer-yomitan/ext/js/app/frontend.js
  - src/core/services/immersion-tracker-service.ts
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/ipc.ts
  - src/preload.ts
  - stats/src/components/sessions/SessionDetail.tsx
  - stats/src/components/library/MediaHeader.tsx
  - stats/src/components/anime/AnimeDetailView.tsx
documentation:
  - docs/plans/2026-03-17-yomitan-lookup-stats-design.md
priority: medium
ordinal: 114500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a dedicated intentional-Yomitan lookup metric so the stats app can show when and how often the user performed real Yomitan lookups while watching video. Keep existing annotation/known-word lookup counters unchanged. Surface the new metric in session detail, episode/media detail, and anime detail, including lookup rate based on words seen.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Successful Yomitan searches while watching create a dedicated Yomitan lookup event and aggregate counter without changing existing lookupCount or lookupHits behavior
- [x] #2 Session detail shows Yomitan lookup timeline markers plus lookup count and lookup rate using words seen
- [x] #3 Episode/media detail shows aggregated Yomitan lookup count and lookup rate using episode totals
- [x] #4 Anime detail shows aggregated Yomitan lookup count and lookup rate using anime totals
- [x] #5 Automated tests cover the new lookup event path, aggregate queries, and affected stats UI surfaces
- [x] #6 Internal docs/plans reflect the approved design and implementation approach
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a SubMiner-specific Yomitan lookup signal emitted from vendored Yomitan on searchSuccess and bridge it through renderer, preload, and main IPC to a tracker hook.
2. Extend immersion tracking with a dedicated Yomitan lookup event type and yomitanLookupCount aggregate, preserving existing lookupCount and lookupHits semantics.
3. Update session, media, anime, and anime-episode queries plus shared stats types to expose the new aggregate count.
4. Update stats UI to show Yomitan lookup markers in session detail and lookup count/rate at session, episode/media, and anime levels using lookups per 100 words copy.
5. Verify with focused unit tests first, then repo typecheck/test/build lanes, and finalize TASK-177 with implementation notes and acceptance-criteria checks.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Approved design recorded in docs/plans/2026-03-17-yomitan-lookup-stats-design.md.

Observed pre-existing local changes in tracker/query/session stats files; implementation plan must preserve those edits while layering Yomitan lookup tracking on top.

Implemented a dedicated Yomitan lookup signal on vendored searchSuccess, bridged it through renderer/preload/main IPC, and persisted YOMITAN_LOOKUP events plus yomitanLookupCount without changing existing annotation lookup counters.

Extended stats queries/types for session, media, anime, and episode aggregates; updated session detail, media header, episode list, and anime overview to show Yomitan lookup counts and lookup rate copy as lookups per 100 words.

Focused verification passed for IPC, tracker service/query, and stats UI tests. stats typecheck still has pre-existing unrelated failures in stats/src/components/anime/AnilistSelector.tsx and stats/src/lib/reading-utils.ts.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added intentional Yomitan lookup tracking end-to-end: vendored Yomitan searchSuccess now emits a SubMiner event, the app records dedicated YOMITAN_LOOKUP events and yomitanLookupCount aggregates, and the stats UI surfaces lookup counts/rates for sessions, episodes/media, and anime. Focused regression tests pass for the IPC bridge, tracker persistence/querying, and new stats UI helpers/components. Full `bun run typecheck:stats` remains blocked by unrelated existing errors in `stats/src/components/anime/AnilistSelector.tsx` and `stats/src/lib/reading-utils.ts`.
<!-- SECTION:FINAL_SUMMARY:END -->
