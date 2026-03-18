---
id: TASK-186
title: Remove stats Library tab and add episode detail navigation from anime page
status: Done
assignee:
  - codex
created_date: '2026-03-17 23:19'
updated_date: '2026-03-17 23:29'
labels:
  - stats
  - ui
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/stats/src/App.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/layout/TabBar.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/anime/AnimeDetailView.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/anime/EpisodeList.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/library/MediaDetailView.tsx
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the stats UI so watched-file detail is no longer exposed as a top-level Library tab. Users should open dedicated episode detail pages from the anime detail page while preserving inline quick-peek session expansion.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Stats navigation no longer shows a top-level Library tab.
- [x] #2 Anime episode rows keep inline quick-peek expansion and also expose an explicit control to open the dedicated episode detail page.
- [x] #3 Dedicated episode detail navigation lands on the existing watched-file detail view with a back action that returns to the originating anime detail page.
- [x] #4 Relevant stats component tests cover the new navigation flow and removed tab behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add app-level stats navigation state for dedicated media detail so anime flows can open watched-file detail without a Library tab.
2. Remove the Library tab from the tab bar and top-level tab panels while preserving existing Overview/Anime/Trends/Vocabulary/Sessions behavior.
3. Update anime detail episode list to keep row expansion for quick peek and add an explicit button that opens the dedicated detail view for the selected episode.
4. Reuse MediaDetailView for episode detail and adjust its back action to return to the originating anime detail page.
5. Add or update stats component tests to cover the removed Library tab and the new anime-to-episode-detail navigation flow.
6. Run targeted stats tests, then targeted SubMiner verification lanes if needed for touched files.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented app-level stats navigation state for dedicated media detail and removed the Library tab from the tab bar and top-level panels.

Anime episode rows now keep inline quick-peek expansion and expose a visible Details button that opens the dedicated watched-file detail view.

Reused MediaDetailView for anime-origin episode navigation with a Back to Anime label and app-level return path.

Verification: bun test stats/src/lib/stats-navigation.test.ts stats/src/lib/stats-ui-navigation.test.tsx; bun run build:stats; bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core ... => passed.

Observed unrelated existing stats workspace issues outside this task when running bun run typecheck:stats, including AnilistSelector/reading-utils/vocabulary-tab and an outdated AnimeOverviewStats test signature.

Reopened for bugfix: episode Details button is a no-op when anime detail is open from within AnimeTab because app-level selectedAnimeId is not retained there. Follow-up fix will pass animeId explicitly through the callback chain instead of depending on App route state.

Bugfix: the Details button now passes animeId explicitly from AnimeTab/AnimeDetailView into app-level media-detail navigation, so dedicated episode navigation works even when the anime page was opened from within the tab rather than seeded by App state.

Bugfix verification: bun test stats/src/lib/stats-navigation.test.ts stats/src/lib/stats-ui-navigation.test.tsx; bun run build:stats => passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Removed the stats Library tab and replaced that navigation path with app-level dedicated media-detail routing from the anime page. Episode rows still support inline quick peek, and now also provide a Details button that opens the dedicated episode view and returns cleanly to the anime detail page. Added navigation-focused tests for the removed tab and anime-origin media-detail flow, and verified the change with targeted tests, stats bundle build, and the repo core verification lane.
<!-- SECTION:FINAL_SUMMARY:END -->
