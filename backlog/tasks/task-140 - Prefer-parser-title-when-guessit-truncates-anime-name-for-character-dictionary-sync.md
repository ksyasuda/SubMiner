---
id: TASK-140
title: Fix guessit title parsing for character dictionary sync
status: Done
assignee: []
created_date: '2026-03-09 00:00'
updated_date: '2026-03-09 00:25'
labels:
  - dictionary
  - anilist
  - bug
  - guessit
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/anilist/anilist-updater.ts
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/anilist/anilist-updater.test.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix AniList character dictionary auto-sync for filenames where `guessit` misparses the full path and our title extraction keeps only the first array segment, causing AniList resolution to match the wrong anime and abort merged dictionary refresh.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AniList media guessing passes basename-only targets to `guessit` so parent folder names do not corrupt series title detection.
- [x] #2 Guessit title arrays are combined into one usable title instead of truncating to the first segment.
- [x] #3 Regression coverage includes the Bunny Girl Senpai filename shape that previously resolved to the wrong AniList entry.
- [x] #4 Verification confirms the targeted AniList guessing tests pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root repro: `guessit` parsed the Bunny Girl Senpai full path as `title: ["Rascal", "Does-not-Dream-of-Bunny-Girl-Senapi"]`, and our `firstString` helper kept only `Rascal`, which resolved to AniList 3490 (`rayca`) and produced zero character results. Fixed by sending basename-only input to `guessit` and joining multi-part guessit title arrays.
<!-- SECTION:NOTES:END -->
