---
id: TASK-107
title: 'Fix Yomitan scan-token fallback fragmentation on exact-source misses'
status: Done
assignee: []
created_date: '2026-03-07 01:10'
updated_date: '2026-03-07 01:12'
labels: []
dependencies: []
priority: high
ordinal: 9007
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Left-to-right Yomitan scanning can emit bogus fallback tokens when `termsFind` returns entries but none of their headwords carries an exact primary source for the consumed substring. Repro: `だが それでも届かぬ高みがあった` currently yields trailing fragments like `があ` / `た`, which blocks the real `あった` token from receiving frequency highlighting.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Scanner skips `termsFind` fallback entries that are not backed by an exact primary source for the consumed substring.
- [x] #2 Repro line no longer yields bogus trailing fragments such as `があ`.
- [x] #3 Regression coverage added for the scan-token path.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Removed the scan-token helper fallback that previously emitted a token from the first returned headword even when Yomitan did not report an exact primary source for the consumed substring. Added a focused regression test covering `だが それでも届かぬ高みがあった`, ensuring bogus `があ` fragmentation is skipped so the later `あった` exact match can still be tokenized and highlighted.

Verification:

- `bun test src/core/services/tokenizer/yomitan-parser-runtime.test.ts src/core/services/tokenizer.test.ts --timeout 20000`

<!-- SECTION:FINAL_SUMMARY:END -->
