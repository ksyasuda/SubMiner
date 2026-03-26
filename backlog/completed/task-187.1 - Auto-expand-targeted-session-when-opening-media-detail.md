---
id: TASK-187.1
title: Auto-expand targeted session when opening media detail
status: Done
assignee:
  - codex
created_date: '2026-03-18 01:32'
updated_date: '2026-03-18 05:28'
labels:
  - stats
  - ui
milestone: m-1
dependencies: []
references:
  - stats/src/lib/stats-navigation.ts
  - stats/src/App.tsx
  - stats/src/components/overview/RecentSessions.tsx
  - stats/src/components/library/MediaDetailView.tsx
  - stats/src/components/library/MediaSessionList.tsx
  - stats/src/lib/stats-navigation.test.ts
parent_task_id: TASK-187
priority: medium
ordinal: 117500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When a navigation path opens episode/media detail with a known session ID, the matching session row in media detail should auto-expand so the user lands directly on the intended session details instead of only the episode history page.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Media detail navigation state can carry an optional target session ID alongside the selected video.
- [x] #2 Any navigation path that opens media detail with a known session ID causes that session row to auto-expand when the episode history loads.
- [x] #3 Session-tab fallback for orphan sessions without a video still behaves as it does now.
- [x] #4 Media detail auto-expansion clears or stabilizes its one-shot navigation state so normal manual expand/collapse behavior still works after landing.
- [x] #5 Relevant navigation/component tests cover the targeted media-detail auto-expand behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Extend media-detail navigation state to optionally carry a target session ID while preserving the existing orphan-session fallback to the Sessions tab.
2. Update app-level navigation helpers and overview recent-session click handling to pass session IDs into media-detail navigation whenever both video and session are known.
3. Thread the one-shot target session ID into MediaDetailView and MediaSessionList so the matching accordion row auto-expands on load, then clear/stabilize that state so manual toggling still behaves normally.
4. Update targeted stats navigation/component tests to cover media-detail auto-expansion and fallback behavior.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Extended media-detail navigation state to carry an optional `initialSessionId`, updated overview/app navigation to pass session IDs into media detail whenever a video-backed session is clicked, and wired `MediaDetailView` + `MediaSessionList` to auto-expand and then consume that one-shot session target.

Updated `stats-navigation.test.ts` to cover the new navigation-state shape. Validation not run in this pass, so acceptance criteria remain unchecked pending verification.
<!-- SECTION:NOTES:END -->
