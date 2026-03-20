---
id: TASK-203
title: Restore known and JLPT annotation for reading-mismatch subtitle tokens
status: Done
assignee:
  - Codex
created_date: '2026-03-19 18:25'
updated_date: '2026-03-19 18:25'
labels:
  - subtitle
  - bug
dependencies: []
references:
  - src/core/services/tokenizer/annotation-stage.ts
  - src/core/services/tokenizer/annotation-stage.test.ts
priority: medium
ordinal: 105721
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Some subtitle tokens lose both known-word coloring and JLPT underline even though the popup resolves a valid dictionary term. Repro example: `大体` in `大体 僕だって困ってたんですよ！` can be known via kana-only Anki data (`だいたい`) while JLPT lookup should still resolve from the kanji surface/headword.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Subtitle annotation can mark a token known via its reading when the configured headword/surface lookup misses.
- [x] #2 JLPT eligibility no longer drops valid kanji terms just because their reading contains repeated kana patterns.
- [x] #3 Regression coverage locks the combined known + JLPT case for `大体`.
<!-- AC:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->

Known-word annotation now falls back to the token reading after the configured headword/surface lookup misses, so kana-only known-card entries still light up matching subtitle tokens. JLPT eligibility now ignores repeated-kana noise checks on the reading when a real surface/headword is present, which preserves JLPT tagging for words like `大体`.

Verification:

- `bun test src/core/services/tokenizer/annotation-stage.test.ts`

<!-- SECTION:OUTCOME:END -->
