---
id: TASK-298
title: Exclude kana grammar-helper merges like ことに from subtitle annotations
status: Done
assignee:
  - codex
created_date: '2026-04-26 00:08'
updated_date: '2026-04-26 00:15'
labels:
  - tokenizer
  - annotations
  - bug
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix subtitle tokenizer annotation behavior where all-hiragana grammar-helper merged tokens such as `ことに` can be marked as N+1. Current likely path: Yomitan emits `ことに` with headword `こと`; MeCab enrichment supplies content-led POS (`名詞|助詞`, likely `非自立|格助詞`); shared subtitle annotation filter does not exclude this family unless it matches narrower rules such as `これで` or explanatory endings.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `ことに`-style kana grammar-helper merges are not marked known, N+1, JLPT, or frequency-highlighted when their MeCab metadata indicates a non-independent noun plus helper particle.
- [x] #2 Regression coverage demonstrates the reported subtitle phrase does not mark `ことに` as N+1 while preserving annotation for real lexical content tokens.
- [x] #3 Existing tokenizer annotation tests pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Approved approach (user: "let's do it"):
1. Add a regression test for the reported `ことに` case using Yomitan token `ことに` -> headword `こと` and MeCab metadata `名詞|助詞` / `非自立|格助詞`; assert all annotation fields are stripped while nearby lexical content can still be N+1.
2. Verify the new test fails before production changes.
3. Update the shared subtitle annotation filter to exclude conservative kana-only grammar-helper merges: merged surface differs from headword, surface is kana-only, first POS component is `名詞`, first POS2 component is `非自立`, and remaining POS components are grammar helpers (`助詞`/`助動詞`).
4. Run targeted tokenizer/annotation tests and update the task acceptance criteria/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Red test initially passed with headword `こと` because `こと` is already in `JLPT_EXCLUDED_TERMS` and the shared subtitle annotation filter checks that set. Updated regression to the live-risk shape `surface=ことに`, `headword=事`, with MeCab POS `名詞|助詞` / `非自立|格助詞`; this failed before the filter change and passed after.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented a conservative shared subtitle annotation filter for kana-only non-independent noun helper merges. Tokens such as `ことに` with a kanji dictionary headword like `事` are now stripped of known-word, N+1, JLPT, and frequency metadata when MeCab shows the first component as `名詞/非自立` and trailing components as grammar helpers.

Added unit coverage in `src/core/services/tokenizer/annotation-stage.test.ts` and an integration-style tokenizer regression for the reported phrase shape in `src/core/services/tokenizer.test.ts`, verifying `ことに` stays plain while a real lexical token can still become the N+1 target.

Validation: `bun test src/core/services/tokenizer/annotation-stage.test.ts`; `bun test src/core/services/tokenizer.test.ts`; `bun run test:fast`; `bun run changelog:lint`.
<!-- SECTION:FINAL_SUMMARY:END -->
