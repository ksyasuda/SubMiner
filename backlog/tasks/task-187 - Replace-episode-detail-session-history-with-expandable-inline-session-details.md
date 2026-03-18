---
id: TASK-187
title: Replace episode detail session history with expandable inline session details
status: Done
assignee:
  - codex
created_date: '2026-03-17 23:42'
updated_date: '2026-03-18 05:28'
labels:
  - stats
  - ui
milestone: m-1
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/library/MediaDetailView.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/library/MediaSessionList.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionRow.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionDetail.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionsTab.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/overview/OverviewTab.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/overview/RecentSessions.tsx
documentation:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-17-episode-detail-session-accordion-design.md
  - >-
    /Users/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-17-episode-detail-session-accordion.md
priority: medium
ordinal: 102500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the dedicated episode detail page so its session history uses the same expandable session-row behavior as the Sessions page, including inline timeline details and session deletion, instead of navigating away to the Sessions tab. Also update home-page session navigation so recent session links open the associated episode detail page rather than the Sessions tab.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Dedicated episode detail session history uses expandable inline rows styled like the Sessions page instead of linking to the Sessions tab.
- [x] #2 Expanding a session on the episode detail page shows the full existing session detail panel, including the timeline chart and stats.
- [x] #3 Episode detail session rows retain a session delete control with the same behavior and safeguards as the Sessions page.
- [x] #4 Home-page recent session navigation opens the associated episode detail page when a session is tied to a video, instead of routing to the Sessions tab.
- [x] #5 Relevant stats tests cover the inline session expansion/delete behavior and the updated home-page navigation path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for media-detail session accordion structure and for overview-to-media-detail navigation, keeping orphan-session fallback coverage.
2. Rework MediaSessionList to reuse SessionRow and SessionDetail with local expansion state and delete affordance matching the Sessions page.
3. Move media-detail session mutation/delete ownership into MediaDetailView so deletes update the current episode page immediately.
4. Add app-level direct media-detail navigation from overview/home-page session rows when videoId exists; keep Sessions-tab fallback for sessions without videoId.
5. Run targeted tests, stats build, and the SubMiner core verification lane; then update TASK-187 with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added approved design/plan docs at docs/plans/2026-03-17-episode-detail-session-accordion-design.md and docs/plans/2026-03-17-episode-detail-session-accordion.md before implementation.

MediaDetailView now owns local session state and delete handling, derives displayed media aggregates from the current session list, and renders MediaSessionList as an inline accordion instead of a session-page link list.

MediaSessionList now reuses SessionRow and full SessionDetail so episode-level session history matches Sessions-page dropdown behavior and keeps the same delete affordance.

Overview/home-page recent session navigation now prefers dedicated media detail when session.videoId exists and falls back to the Sessions tab only for orphan sessions without videoId.

Verification passed: bun test stats/src/lib/stats-navigation.test.ts stats/src/lib/stats-ui-navigation.test.tsx stats/src/lib/media-session-list.test.tsx; bun run build:stats; bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core stats/src/App.tsx stats/src/components/overview/OverviewTab.tsx stats/src/components/overview/RecentSessions.tsx stats/src/components/library/MediaDetailView.tsx stats/src/components/library/MediaSessionList.tsx stats/src/lib/stats-navigation.ts stats/src/lib/stats-navigation.test.ts stats/src/lib/stats-ui-navigation.test.tsx stats/src/lib/media-session-list.test.tsx => passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Dedicated episode detail pages now show inline expandable session rows using the same shared SessionRow + SessionDetail UI as the Sessions page, including per-session delete controls. Overview/home-page recent session clicks now open the episode detail page whenever a backing video exists, with Sessions-tab fallback only for sessions missing videoId. Added navigation and media-session-list tests plus design/implementation docs, and verified the change with targeted tests, stats bundle build, and the SubMiner core verification lane.
<!-- SECTION:FINAL_SUMMARY:END -->
