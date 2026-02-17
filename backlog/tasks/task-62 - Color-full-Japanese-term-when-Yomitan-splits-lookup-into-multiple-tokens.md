---
id: TASK-62
title: Color full Japanese term when Yomitan splits lookup into multiple tokens
status: Done
assignee: []
created_date: '2026-02-16 23:03'
updated_date: '2026-02-16 23:11'
labels: []
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Users should see one continuous highlight for a looked-up term even when Yomitan returns the term as multiple adjacent tokens, so color feedback matches the selected word/phrase.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When a looked-up Japanese term is represented as multiple adjacent tokens from Yomitan, the UI applies highlight color to the entire contiguous term instead of only one token.
- [x] #2 Existing highlighting behavior for single-token matches remains unchanged.
- [x] #3 Automated coverage or reproducible verification demonstrates the multi-token case is rendered correctly.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update Yomitan parse-result mapping so each parse line is treated as one logical token (combine segment text/reading and preserve the selected headword from segment metadata).
2. Add regression coverage for furigana-split parse lines to ensure frequency/highlight metadata applies to the full combined token.
3. Rebuild and run tokenizer tests to verify multi-segment and single-segment behavior remain correct.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented line-level token mapping in `src/core/services/tokenizer-service.ts` so segmented Yomitan line parts (e.g. furigana-split pieces) are merged into one `MergedToken` with one headword, one surface span, and one reading string.

Added/updated tokenizer tests in `src/core/services/tokenizer-service.test.ts` covering segmented-line behavior and aligned several existing fixtures/assertions to current runtime behavior so the full tokenizer suite is green.

Validation run: `pnpm run build && node dist/core/services/tokenizer-service.test.js` (38/38 passing).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed partial token coloring caused by Yomitan segmented parse lines by changing tokenizer mapping to treat each parse line as one logical token instead of one token per segment. The new mapping concatenates segment text/reading, carries the selected headword from segment metadata, and preserves correct span offsets so frequency/known-word/JLPT classifications apply to the full term span. Added regression coverage for furigana-split tokens and updated related parser fixture tests to reflect line-level token semantics. Verified with `pnpm run build` and `node dist/core/services/tokenizer-service.test.js` (38 passing).
<!-- SECTION:FINAL_SUMMARY:END -->
