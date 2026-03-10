---
id: TASK-101
title: Index AniList character alternative names in the character dictionary
status: Done
assignee: []
created_date: '2026-03-07 00:00'
updated_date: '2026-03-08 00:11'
labels:
  - dictionary
  - anilist
dependencies: []
references:
  - src/main/character-dictionary-runtime.ts
  - src/main/character-dictionary-runtime.test.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Index AniList character alternative names in generated character dictionaries so aliases like Shadow resolve during subtitle lookup instead of falling through to unrelated generic dictionary entries.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Character fetch reads AniList alternative character names needed for lookup coverage
- [x] #2 Generated term banks include alias-derived terms for subtitle lookups like シャドウ
- [x] #3 Regression coverage proves alternative-name indexing works end to end
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Character dictionary generation now requests AniList `name.alternative`, indexes those aliases as term candidates, and expands mixed aliases like `Minoru Kagenou (影野ミノル)` into usable outer/inner variants. Also extended kana alias synthesis so the AniList alias `Shadow` emits `シャドウ`, which matches the subtitle token the user hit in The Eminence in Shadow.

Bumped the character-dictionary snapshot format to invalidate stale cached snapshots, and updated merged-dictionary rebuilds to refresh invalid snapshots before composing the ZIP so old cache files do not hard-fail the merge path.

Verified with `bun test src/main/character-dictionary-runtime.test.ts` and `bun run tsc --noEmit`.

<!-- SECTION:FINAL_SUMMARY:END -->
