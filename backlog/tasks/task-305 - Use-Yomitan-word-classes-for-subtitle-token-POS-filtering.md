---
id: TASK-305
title: Use Yomitan word classes for subtitle token POS filtering
status: Done
assignee: []
created_date: '2026-04-26 05:56'
updated_date: '2026-04-26 05:59'
labels:
  - tokenizer
  - yomitan
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Subtitle annotation filtering currently uses Yomitan token spans, then enriches those spans by running MeCab over the full normalized subtitle line. Add support for carrying Yomitan headword wordClasses from termsFind into SubMiner tokens so dictionary-backed tokens can provide coarse POS/tag metadata without vendored Yomitan changes. MeCab whole-line enrichment should remain a fallback/source of detailed POS data when Yomitan classes are absent.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Yomitan scanner tokens preserve matched headword wordClasses when termsFind returns them.
- [x] #2 Subtitle tokenization maps recognized Yomitan wordClasses to coarse PartOfSpeech/POS metadata before annotation filtering.
- [x] #3 Whole-line MeCab enrichment remains available for missing or more detailed POS metadata and does not break existing subtitle annotation behavior.
- [x] #4 Focused tokenizer tests cover wordClasses extraction and POS mapping.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression coverage for Yomitan scanner wordClasses payload and subtitle POS mapping.
2. Extend the app-owned Yomitan scanner payload to carry matched headword wordClasses when present.
3. Map recognized Yomitan wordClasses to SubMiner coarse PartOfSpeech/POS metadata before annotation filtering.
4. Keep MeCab whole-line enrichment as fallback/detail-fill for missing POS fields.
5. Run focused tokenizer tests and typecheck.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented app-only wordClasses extraction from termsFind results; no vendored Yomitan changes required. Recognized classes currently map prt, aux, v*, adj-i/adj-ix, adj-na, and noun-like classes to SubMiner POS metadata. MeCab enrichment now skips only tokens with complete pos1/pos2/pos3 and otherwise fills missing fields while preserving existing coarse pos1. Verification: bun test src/core/services/tokenizer/yomitan-parser-runtime.test.ts src/core/services/tokenizer.test.ts; bun run typecheck.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented app-only Yomitan wordClasses support for subtitle token annotation filtering. The scanner now carries matched headword wordClasses from termsFind results, tokenizer maps recognized classes into SubMiner coarse POS metadata before annotation, and MeCab whole-line enrichment continues to fill missing detailed POS fields without requiring vendored Yomitan changes.

Tests run:
- bun test src/core/services/tokenizer/yomitan-parser-runtime.test.ts src/core/services/tokenizer.test.ts
- bun run typecheck

Note: the working tree already had unrelated tokenizer/annotation edits and task-304 before this work; those were left intact.
<!-- SECTION:FINAL_SUMMARY:END -->
