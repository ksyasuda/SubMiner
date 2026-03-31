---
id: TASK-219
title: Restore streamed video progress in anime episodes
status: Done
assignee:
  - codex
created_date: '2026-03-22 21:25'
updated_date: '2026-03-31 19:37'
labels:
  - stats
  - immersion-tracker
  - youtube
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/query.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker-service.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/__tests__/query.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker-service.test.ts
priority: medium
ordinal: 178500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Episode progress for streamed media can stay at `0%` because some remote sessions persist `ended_media_ms = 0` even when subtitle timing and watch activity clearly advanced, and the anime episode query currently treats `0` as a valid progress checkpoint.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Anime episode progress ignores zero-valued session checkpoints and falls back to subtitle/event timing
- [x] #2 New streamed sessions persist meaningful progress even when playback-position updates are missing or sparse
- [x] #3 Regression tests cover the zero-checkpoint remote-session case
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored anime episode progress handling for streamed sessions by ignoring zero-valued `ended_media_ms` checkpoints and falling back to subtitle/event timing, with regression coverage for the remote-session zero-checkpoint case.
<!-- SECTION:FINAL_SUMMARY:END -->
