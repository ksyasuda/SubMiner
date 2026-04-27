---
id: TASK-307
title: Exclude kana-only words from N+1 subtitle targets
status: Done
assignee:
  - codex
created_date: '2026-04-27 01:52'
updated_date: '2026-04-27 01:57'
labels:
  - tokenizer
  - annotations
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle N+1 annotation is over-targeting kana-only or hiragana/katakana tokens that collapse to dictionary words. Adjust targeting so kana-only tokens are not selected as N+1 candidates, while preserving tokenization/hover behavior and other annotation metadata where existing filters allow it.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Kana-only subtitle tokens are not marked as N+1 targets.
- [x] #2 Kanji or mixed lexical tokens can still be marked as N+1 targets when they are the single unknown candidate in a sentence.
- [x] #3 Regression coverage demonstrates the kana-only N+1 exclusion.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression in `src/core/services/tokenizer.test.ts` showing a kana-only Yomitan token is not selected as the single N+1 target, while a mixed lexical token in the same style still can be targeted.
2. Implement the smallest filter in `src/token-merger.ts`: N+1 candidate selection rejects tokens whose surface is entirely kana; word-count behavior remains governed by existing annotation/POS filters.
3. Run the focused tokenizer tests, then update task acceptance criteria/final summary.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented a surface-level kana-only guard in N+1 candidate selection. Kept existing word-count/POS filtering behavior intact; updated tokenizer and annotation-stage expectations where old tests intentionally allowed kana-only N+1 targets.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added kana-only surface detection to `isNPlusOneCandidateToken` so hiragana/katakana-only subtitle tokens are not selected as N+1 targets.
- Added/updated tokenizer and annotation-stage regressions for kana-only targets while preserving non-kana N+1 behavior.
- Added changelog fragment `changes/307-kana-nplusone-targets.md`.

Verification:
- `bun test src/core/services/tokenizer.test.ts --test-name-pattern "kana-only N\+1"` failed before the fix with `true !== false`.
- `bun test src/core/services/tokenizer/annotation-stage.test.ts src/core/services/tokenizer.test.ts` passed.
- `bun run typecheck` passed.
- `bun run test:fast` passed.
- `bun run changelog:lint` passed.
- `bunx prettier --check src/core/services/tokenizer.test.ts src/core/services/tokenizer/annotation-stage.test.ts src/token-merger.ts changes/307-kana-nplusone-targets.md` passed.
<!-- SECTION:FINAL_SUMMARY:END -->
