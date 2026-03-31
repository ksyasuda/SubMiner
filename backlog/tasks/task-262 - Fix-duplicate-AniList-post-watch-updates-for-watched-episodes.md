---
id: TASK-262
title: Fix duplicate AniList post-watch updates for watched episodes
status: Done
assignee:
  - codex
created_date: '2026-03-31 19:03'
updated_date: '2026-03-31 19:05'
labels:
  - bug
  - anilist
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Watching an episode can currently produce two AniList activity updates for the same episode. The duplicate happens when the post-watch flow drains a queued retry for the current episode and then proceeds to run the live post-watch update for that same media/episode in the same pass. User report says this reproduces both when crossing the watched threshold naturally and when using the mark-watched keybinding. Fix the duplicate so one successful watch produces at most one AniList progress update for a given mediaKey/episode pair.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A watched episode triggers at most one AniList post-watch progress update for a given media key and episode during a single post-watch pass, even if that episode already exists in the retry queue.
- [x] #2 Both watched-threshold and manual mark-watched flows are protected by regression coverage for the duplicate-update case.
- [x] #3 Relevant user-visible change note is added if required by repo policy.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the duplicate in a unit test around `createMaybeRunAnilistPostWatchUpdateHandler` by simulating a ready retry for the same `mediaKey::episode` the live path would also submit.
2. Fix the handler so that after processing a queued retry, it does not perform a second live update when the retry already satisfied the current attempt key.
3. Run focused AniList runtime tests and adjacent immersion tests to confirm both threshold-driven and manual mark-watched entry points stay covered through the shared post-watch path.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a regression in `src/main/runtime/anilist-post-watch.test.ts` for the case where `processNextAnilistRetryUpdate()` already satisfies the current `mediaKey::episode` before the live path runs.

Updated `createMaybeRunAnilistPostWatchUpdateHandler` to re-check `hasAttemptedUpdateKey(attemptKey)` immediately after draining the retry queue and short-circuit before a second live AniList submission.

Verification: `bun test src/main/runtime/anilist-post-watch.test.ts src/main/runtime/anilist-post-watch-main-deps.test.ts`; `bun test src/core/services/immersion-tracker-service.test.ts --test-name-pattern 'recordPlaybackPosition marks watched at 85% completion|markActiveVideoWatched'`; `bun run typecheck`; `bun run changelog:lint`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed duplicate AniList post-watch submissions by short-circuiting the live update path when a ready retry item already handled the current `mediaKey::episode` in the same pass. Added a focused regression test for the retry-plus-live duplicate scenario and a changelog fragment documenting the fix.
<!-- SECTION:FINAL_SUMMARY:END -->
