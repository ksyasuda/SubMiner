---
id: TASK-88
title: Remove MeCab fallback tokenizer and simplify Yomitan token flow
status: To Do
assignee: []
created_date: '2026-02-20 00:00'
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
- [ ] #1 Subtitle tokenization no longer falls back to MeCab when Yomitan parsing fails.
- [ ] #2 Token grouping logic is simplified to rely on Yomitan structure; redundant custom merge-selection logic removed.
- [ ] #3 Known-word, JLPT, frequency, and N+1 annotations still work on Yomitan-derived tokens.
- [ ] #4 If Yomitan parsing fails, behavior is explicit and tested (for example `tokens: null` without MeCab fallback path).
- [ ] #5 Documentation reflects that tokenization flow is Yomitan-first and Yomitan-only.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `src/core/services/tokenizer.ts` no longer contains MeCab fallback tokenization branch.
- [ ] #2 Tests cover Yomitan-only pipeline and failure behavior regressions.
- [ ] #3 Any removed MeCab-only merge helpers are deleted with no unused exports/imports.
- [ ] #4 Build and relevant tokenizer/subtitle tests pass.
<!-- DOD:END -->
