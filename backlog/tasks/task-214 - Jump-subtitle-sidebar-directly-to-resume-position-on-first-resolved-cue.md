---
id: TASK-214
title: Jump subtitle sidebar directly to resume position on first resolved cue
status: In Progress
assignee: []
created_date: '2026-03-21 11:15'
updated_date: '2026-03-21 11:15'
labels:
  - bug
  - ux
  - overlay
  - subtitles
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When playback starts from a resumed timestamp while the subtitle sidebar is open, the sidebar currently smooth-scrolls from the top of the cue list to the resumed cue. Change the first resolved active-cue positioning to jump immediately to the resume location while preserving smooth auto-follow for later playback-driven cue advances.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The first active cue resolved after open/resume uses an instant jump instead of smooth-scrolling through the list.
- [x] #2 Normal subtitle-sidebar auto-follow remains smooth after the first active cue has been positioned.
- [x] #3 Regression coverage distinguishes the initial jump behavior from later smooth auto-follow updates.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-21: Fixed by treating the first auto-scroll from `previousActiveCueIndex < 0` as `behavior: 'auto'` in the subtitle sidebar scroll helper. Added renderer regression coverage for initial jump plus later smooth follow.
<!-- SECTION:NOTES:END -->
