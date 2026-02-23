---
id: TASK-88
title: Remove MeCab fallback tokenizer and simplify Yomitan token flow
status: Done
assignee: []
created_date: '2026-02-20 00:00'
updated_date: '2026-02-23 01:44'
labels:
  - tokenizer
  - refactor
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Remove the MeCab fallback tokenization path and associated merge-selection complexity in subtitle tokenization. Treat Yomitan parser output as the single source of token boundaries/grouping, and keep only minimal normalization needed for downstream known-word, JLPT, and frequency annotation.

<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->

1. Remove MeCab fallback execution from `tokenizeSubtitle` and delete dead fallback-specific branches.
2. Remove merge/candidate-selection code that is only needed to reconcile MeCab-vs-Yomitan tokenization strategies.
3. Keep Yomitan parsing pipeline with minimal structural token normalization only.
4. Update MeCab usage so it is no longer required for tokenization fallback (retain only explicitly needed behavior, if any).
5. Update docs/config notes to reflect Yomitan-only tokenization flow.
6. Add regression tests for Yomitan-only success/failure paths and token annotation continuity.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Subtitle tokenization no longer falls back to MeCab when Yomitan parsing fails.
- [x] #2 Token grouping logic is simplified to rely on Yomitan structure; redundant custom merge-selection logic removed.
- [x] #3 Known-word, JLPT, frequency, and N+1 annotations still work on Yomitan-derived tokens.
- [x] #4 If Yomitan parsing fails, behavior is explicit and tested (for example `tokens: null` without MeCab fallback path).
- [x] #5 Documentation reflects that tokenization flow is Yomitan-first and Yomitan-only.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Removed MeCab fallback tokenization from `src/core/services/tokenizer.ts`; `tokenizeSubtitle` now returns `tokens: null` when Yomitan parsing/selecting yields no tokens.

Simplified parse candidate selection in `src/core/services/tokenizer/parser-selection-stage.ts` to scanning-parser sources only; added null behavior when only `mecab` parse candidates are present.

Updated tokenizer regression suites to reflect Yomitan-only flow while preserving annotation continuity checks (known-word, JLPT, frequency, N+1) in `src/core/services/tokenizer.test.ts` and `src/core/services/tokenizer/parser-selection-stage.test.ts`.

Updated docs to remove MeCab fallback positioning and clarify Yomitan-only tokenization in `docs/usage.md` and `docs/troubleshooting.md`.

Validation run passed:

- `bun test src/core/services/tokenizer/parser-selection-stage.test.ts src/core/services/tokenizer.test.ts`
- `bun test src/core/services/subtitle-processing-controller.test.ts`
- `bun run build`
- `bun run docs:build`
<!-- SECTION:NOTES:END -->

## Definition of Done

<!-- DOD:BEGIN -->

- [x] #1 `src/core/services/tokenizer.ts` no longer contains MeCab fallback tokenization branch.
- [x] #2 Tests cover Yomitan-only pipeline and failure behavior regressions.
- [x] #3 Any removed MeCab-only merge helpers are deleted with no unused exports/imports.
- [x] #4 Build and relevant tokenizer/subtitle tests pass.
<!-- DOD:END -->
