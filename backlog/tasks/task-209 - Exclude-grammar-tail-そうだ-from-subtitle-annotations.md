---
id: TASK-209
title: Exclude grammar-tail そうだ from subtitle annotations
status: Done
assignee:
  - codex
created_date: '2026-03-20 04:06'
updated_date: '2026-03-20 04:33'
labels:
  - bug
  - tokenizer
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/annotation-stage.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer/annotation-stage.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/tokenizer.test.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Sentence-final grammar-tail `そうだ` tokens can still receive subtitle annotation styling, including frequency highlighting, when Yomitan returns a standalone `そうだ` token and MeCab enriches it as an auxiliary-stem/coupla pattern (`名詞|助動詞`, `助動詞語幹`). Keep the subtitle text visible, but treat this grammar tail like other grammar-only endings so it renders without annotation metadata.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Sentence-final grammar-tail `そうだ` tokens enriched as auxiliary-stem/copula patterns do not receive frequency highlighting or other subtitle annotation metadata.
- [x] #2 The preceding lexical token in cases like `与えるそうだ` keeps its existing annotation behavior.
- [x] #3 Regression tests cover the annotation-stage exclusion and end-to-end subtitle tokenization for the `そうだ` grammar-tail case.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression coverage for the reported `与えるそうだ` case at both annotation-stage and tokenizeSubtitle levels.
2. Reproduce failure by modeling the MeCab-enriched grammar-tail shape (`名詞|助動詞`, `特殊`, `助動詞語幹`) that currently keeps frequency metadata.
3. Update subtitle-annotation exclusion logic to recognize auxiliary-stem/copula grammar tails via POS metadata plus normalized tail text, not a raw sentence-specific string match.
4. Re-run targeted tokenizer and annotation-stage tests, then record the verification commands and outcome in the task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Investigated reported `与えるそうだ` case. MeCab tags `そう` as `名詞,特殊,助動詞語幹` and `だ` as `助動詞`; after overlap enrichment the Yomitan token becomes `pos1=名詞|助動詞`, `pos2=特殊`, `pos3=助動詞語幹`, which currently escapes subtitle-annotation exclusion and can keep a frequency rank.

Implemented a POS-shape subtitle-annotation exclusion for MeCab-enriched auxiliary-stem grammar tails. The new predicate keys off merged tokens whose POS tags stay within `名詞/助動詞/助詞` and whose POS3 includes `助動詞語幹`, which clears annotation metadata for `そうだ`-style tails without hard-coding the full subtitle text.

Verification: `bun test src/core/services/tokenizer/annotation-stage.test.ts`, `bun test src/core/services/tokenizer.test.ts --test-name-pattern 'explanatory ending|interjection|single-kana merged tokens from frequency highlighting|auxiliary-stem そうだ grammar tails|composite function/content token from frequency highlighting|keeps frequency for content-led merged token with trailing colloquial suffixes'`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added regression coverage for `与えるそうだ` and updated subtitle annotation exclusion logic to drop annotation metadata for MeCab-enriched auxiliary-stem grammar tails. The fix is POS-driven rather than sentence-specific, so `そうだ`-style grammar endings stay visible/hoverable as plain text while neighboring lexical tokens keep their existing frequency/JLPT behavior.
<!-- SECTION:FINAL_SUMMARY:END -->
