---
id: TASK-94
title: Add kana aliases for AniList character dictionary entries
status: Done
assignee: []
created_date: '2026-03-06 09:20'
updated_date: '2026-03-16 05:13'
labels:
  - dictionary
  - tokenizer
  - anilist
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.test.ts
priority: high
ordinal: 84500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Generate katakana/hiragana-friendly aliases from AniList romanized character names so subtitle katakana names like カズマ match character dictionary entries even when AniList native name is kanji.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 AniList character dictionary generation adds kana aliases for romanized names when native name is not already kana-only
- [x] #2 Generated dictionary entries allow katakana subtitle names like カズマ to resolve against a kanji-native AniList character entry
- [x] #3 Regression tests cover alias generation and resulting term bank output
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added katakana aliases synthesized from AniList romanized character names during character dictionary generation, so kanji-native entries such as 佐藤和真 / Satou Kazuma now also emit terms like カズマ and サトウカズマ with hiragana readings. Added regression coverage verifying generated term-bank output for the Konosuba case.

Verified with `bun test src/main/character-dictionary-runtime.test.ts` and `bun run tsc --noEmit`.
<!-- SECTION:FINAL_SUMMARY:END -->
