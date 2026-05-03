---
id: TASK-319
title: Suppress annotations for expressive interjection subtitles
status: Done
assignee:
  - Codex
created_date: '2026-05-03 03:18'
updated_date: '2026-05-03 03:20'
labels:
  - bug
  - subtitle-annotations
dependencies: []
references:
  - src/core/services/tokenizer/subtitle-annotation-filter.ts
  - src/core/services/tokenizer/annotation-stage.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Interjection-only subtitle tokens such as ハァ and はっ should remain hoverable as tokens but must not receive known, N+1, frequency, or JLPT annotation styling. Current behavior can still annotate these forms when dictionary/POS metadata does not trip the existing exclusion gate.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Standalone ハァ/はっ-style interjection tokens have annotation metadata cleared even when dictionary metadata exists.
- [x] #2 Filtering remains scoped so content-bearing non-interjection tokens still receive annotations.
- [x] #3 Regression coverage exercises the reported subtitle pattern: ハァ… / （ガーフィール）はっ！
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression coverage around annotation filtering for the reported interjection forms, including katakana ハァ and small-tsu はっ with surrounding subtitle punctuation/name text.
2. Tighten the shared subtitle annotation exclusion gate so expressive kana interjections clear annotation metadata without relying only on MeCab pos1=感動詞.
3. Run the focused tokenizer/annotation tests, then update acceptance criteria and notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented via shared subtitle annotation exclusion term normalization: added はぁ so katakana ハァ normalizes into the existing term gate. Existing small-tsu kana SFX logic already covers はっ. Regression confirms both reported forms clear known/N+1/frequency/JLPT metadata while a normal noun keeps frequency annotation.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added a regression for the reported subtitle pattern ハァ… / （ガーフィール）はっ！, with annotation metadata present on both interjection tokens.
- Extended the shared subtitle annotation exclusion term set so ハァ normalizes to はぁ and is stripped of annotation styling. Existing はっ handling remains covered by small-tsu kana SFX filtering.
- Added a change fragment for the user-visible bug fix.

Verification:
- bun test src/core/services/tokenizer/annotation-stage.test.ts
- bun test src/core/services/tokenizer/annotation-stage.test.ts src/core/services/tokenizer.test.ts src/renderer/subtitle-render.test.ts
- bun run typecheck
<!-- SECTION:FINAL_SUMMARY:END -->
