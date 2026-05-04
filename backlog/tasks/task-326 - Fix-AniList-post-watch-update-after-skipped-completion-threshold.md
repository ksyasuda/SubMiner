---
id: TASK-326
title: Fix AniList post-watch update after skipped completion threshold
status: In Progress
assignee: []
created_date: '2026-05-04 00:33'
labels:
  - anilist
  - bug
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
AniList episode progress should sync reliably when playback reaches or passes the watched trigger point, even if mpv progress events jump over the exact threshold. Investigate why a completed watched episode did not update AniList and fix the root cause for post-watch tracking.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When playback moves from before the completion threshold to any later position at or beyond the threshold, AniList queues or sends the episode progress update once.
- [x] #2 If playback is already past the completion threshold and the update has not yet been recorded for the current media/episode, AniList still queues or sends the update.
- [x] #3 AniList progress updates remain deduplicated for the same media/episode watch completion.
- [x] #4 A regression test covers the skipped-threshold or already-past-threshold case.
<!-- AC:END -->

## Notes

- Fixed mpv `time-pos` ordering so post-watch checks read the fresh playback position after seeks.
- Wired manual mark-watched to run a forced AniList post-watch sync after the local watched mark succeeds.
- Added regressions for time-position ordering, manual watched sync, forced post-watch updates, and the Little Witch Academia filename parse.
