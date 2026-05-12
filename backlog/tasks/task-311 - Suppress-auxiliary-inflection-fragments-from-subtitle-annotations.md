---
id: TASK-311
title: Suppress auxiliary inflection fragments from subtitle annotations
status: Done
assignee: []
created_date: '2026-05-02 09:07'
updated_date: '2026-05-02 09:10'
labels:
  - tokenizer
  - annotations
  - bug
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Suppress standalone Japanese auxiliary/inflection subtitle fragments such as `れる` and `れた` from frequency/JLPT/N+1/known annotation styling while keeping lexical verbs such as `くれ` / `くれる` annotatable. Tokens must remain hoverable; only annotation metadata should be stripped.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `れる` and `れた`-style standalone helper fragments render as plain hoverable subtitle tokens.
- [x] #2 Lexical verbs like `くれ` / `くれる` remain eligible for annotation.
- [x] #3 Regression tests cover unit filter behavior and tokenizer integration.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented with TDD. Added failing coverage first for standalone `れる`/`れた` auxiliary fragments and a lexical `くれ`/`くれる` guard. Updated the shared subtitle annotation filter to strip annotation metadata for kana-only auxiliary inflection fragments identified by MeCab POS (`助動詞` only, or `動詞/接尾` with optional trailing `助動詞`) while preserving lexical `くれ` as `くれる` when tagged `動詞/自立`. Added tokenizer integration coverage for `れた` and neighboring lexical N+1 behavior.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Suppressed annotation metadata for standalone auxiliary inflection fragments such as `れる` and `れた` in subtitle tokens, leaving them hoverable but plain. Preserved lexical `くれ` -> `くれる` verb metadata when MeCab tags it as `動詞/自立`.

Added unit and tokenizer regression coverage, plus a release fragment in `changes/311-auxiliary-inflection-annotation-filter.md`.

Validation: targeted annotation/tokenizer tests passed; `bun run typecheck` passed; `bun run changelog:lint` passed. `bun run test:fast` was attempted twice and failed in unrelated `src/core/services/subsync.test.ts` cross-file state (`window.electronAPI` undefined), while `bun test src/core/services/subsync.test.ts` passes by itself.
<!-- SECTION:FINAL_SUMMARY:END -->
