---
id: TASK-341
title: Fix frequency highlight for honorific prefix noun tokens
status: Done
assignee:
  - codex
created_date: '2026-05-05 02:08'
updated_date: '2026-05-05 02:10'
labels:
  - bug
  - tokenizer
  - frequency
dependencies: []
documentation:
  - docs/architecture/2026-03-15-renderer-performance-design.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User reported subtitle token `ご機嫌` in `（フランク）ご機嫌が良くないようだな アンドリュー` shows Yomitan/JPDB rank 5484 in popup but is not highlighted as frequent. Frequency annotation currently excludes merged tokens containing default-excluded POS parts such as `接頭詞`; ordinal prefix-noun tokens already have an exception. Desired outcome: honorific prefix + noun lexical tokens like `ご機嫌` keep their valid frequency rank so renderer can apply frequent-token styling, while standalone prefixes and noisy merged grammar fragments remain excluded.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `ご機嫌`-style honorific prefix + noun tokens retain a finite frequency rank after annotation/tokenization when frequency highlighting is enabled.
- [x] #2 Standalone prefix/noise tokens remain excluded from frequency annotation.
- [x] #3 Regression test covers the reported `ご機嫌` rank 5484 behavior.
- [x] #4 Relevant tokenizer/annotation tests pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression around honorific prefix + noun token frequency retention, using `ご機嫌` with rank 5484 and POS `接頭詞|名詞` / `名詞接続|一般`.
2. Implement a narrow annotation-stage exception for lexical honorific prefix-noun tokens, adjacent to the existing ordinal prefix-noun allowance.
3. Verify standalone prefix/noise exclusion behavior remains covered.
4. Run targeted tokenizer/annotation tests and update acceptance criteria/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
TDD red verified: `bun test src/core/services/tokenizer.test.ts -t "honorific prefix-noun"` failed with `actual: undefined`, `expected: 5484` before implementation.

Implemented a narrow honorific prefix-noun frequency allowance for merged `お`/`ご`/`御` + noun tokens with POS `接頭詞|名詞` and prefix POS2 `名詞接続`. Existing standalone prefix/noise exclusion tests still pass.

Verification: `bun test src/core/services/tokenizer.test.ts src/core/services/tokenizer/annotation-stage.test.ts` passed (164 tests); `bun run typecheck` passed; `bunx prettier --check src/core/services/tokenizer/annotation-stage.ts src/core/services/tokenizer.test.ts` passed. Repo-wide `bun run format:check:src` still fails on pre-existing `src/core/services/stats-window.ts` formatting.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed frequency annotation for lexical honorific prefix-noun tokens such as `ご機嫌`. The annotation filter now allows merged `お`/`ご`/`御` prefix + noun tokens with MeCab POS `接頭詞|名詞` / `名詞接続|...` to retain a valid frequency rank, while standalone prefixes and existing noise filters remain excluded.

Added a tokenizer regression for the reported `ご機嫌` case asserting rank `5484` is preserved after MeCab enrichment and annotation.

Verification:
- `bun test src/core/services/tokenizer.test.ts -t "honorific prefix-noun"` failed before the fix with `undefined` vs `5484`, then passed after the fix.
- `bun test src/core/services/tokenizer.test.ts src/core/services/tokenizer/annotation-stage.test.ts` passed (164 tests).
- `bun run typecheck` passed.
- `bunx prettier --check src/core/services/tokenizer/annotation-stage.ts src/core/services/tokenizer.test.ts` passed.

Note: repo-wide `bun run format:check:src` currently fails on unrelated existing formatting in `src/core/services/stats-window.ts`.
<!-- SECTION:FINAL_SUMMARY:END -->
