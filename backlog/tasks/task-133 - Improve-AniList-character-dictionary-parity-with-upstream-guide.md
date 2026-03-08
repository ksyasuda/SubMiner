---
id: TASK-133
title: Improve AniList character dictionary parity with upstream guide
status: To Do
assignee: []
created_date: '2026-03-08 21:06'
updated_date: '2026-03-08 21:35'
labels:
  - dictionary
  - anilist
  - planning
dependencies: []
references:
  - >-
    https://github.com/bee-san/Japanese_Character_Name_Dictionary/blob/main/docs/agents_read_me.md
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.test.ts
documentation:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-08-anilist-character-dictionary-parity-design.md
  - >-
    /Users/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-08-anilist-character-dictionary-parity.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Plan and implement guide-faithful parity improvements for the AniList character dictionary flow inside SubMiner's current single-media generation path. Scope includes AniList first/last name hints, hint-aware reading generation for kanji/native names, expanded honorific coverage, 160x200 JPEG thumbnail handling, and AniList 429 retry/backoff behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 AniList character queries include first/last name fields and preserve them through runtime data models.
- [ ] #2 Dictionary generation uses hint-aware name splitting and reading generation for kanji and mixed native names, not only kana-only readings.
- [ ] #3 Honorific generation is expanded substantially toward upstream coverage and is covered by regression tests.
- [ ] #4 Character and voice-actor images are resized or re-encoded to bounded JPEG thumbnails with fallback behavior.
- [ ] #5 AniList requests handle 429 responses with bounded exponential backoff and tests cover retry behavior.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Approved design and implementation plan captured on 2026-03-08. Scope stays within current single-media AniList dictionary flow; excludes username-driven CURRENT-list fetching and Yomitan auto-update schema work.
<!-- SECTION:NOTES:END -->
