---
id: TASK-176
title: Exclude interjections and sound effects from subtitle annotations
status: Done
assignee:
  - codex
created_date: '2026-03-15 12:07'
updated_date: '2026-03-16 05:13'
labels:
  - bug
  - tokenizer
  - renderer
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/annotation-stage.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.test.ts
  - /home/sudacode/projects/japanese/SubMiner/src/renderer/subtitle-render.ts
  - >-
    /home/sudacode/projects/japanese/SubMiner/src/renderer/subtitle-render.test.ts
priority: high
ordinal: 16500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle tokens that are not useful annotation targets, especially interjections and sound-effect / onomatopoeia-style exclamations such as `ぐはっ` and `はあ`, can still survive tokenization and become interactive hover annotations. Keep the subtitle text visible, but remove these tokens from annotation payloads so they do not render hover targets or dictionary popovers.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Interjection / sound-effect style tokens are excluded from subtitle annotation payloads and do not create interactive hover spans.
- [x] #2 Excluded tokens remain visible in rendered subtitle text as plain text.
- [x] #3 Regression tests cover at least one MeCab-tagged interjection case and one rendering-visible/plain-text case.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression coverage proving excluded tokens still come through visibly in subtitle text but no longer survive as annotation tokens.
2. Introduce a shared annotation-eligibility predicate in the tokenizer annotation stage for interjections / SFX-like tokens.
3. Filter subtitle token payloads through that predicate before renderer hover ranges/spans are built.
4. Verify with targeted tokenizer and renderer tests.
<!-- SECTION:PLAN:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a subtitle-annotation exclusion pass after token annotation so interjections and obvious SFX-style tokens are removed from returned token payloads while the original subtitle text stays intact. Coverage now includes MeCab-tagged `感動詞`, repeated-kana interjections such as `ああ`, a mixed `ぐはっ 猫` tokenizer case, and a renderer check proving omitted tokens stay visible as plain text instead of interactive hover spans.
<!-- SECTION:FINAL_SUMMARY:END -->
