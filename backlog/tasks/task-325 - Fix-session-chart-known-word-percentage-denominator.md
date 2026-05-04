---
id: TASK-325
title: Fix session chart known-word percentage denominator
status: Done
assignee: []
created_date: '2026-05-04 01:19'
updated_date: '2026-05-04 01:23'
labels:
  - stats
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Session detail known-word percentages should use the same filtered vocabulary occurrence rows for both known and total word counts. Current chart can divide known persisted word occurrences by raw token totals, causing excluded tokens to depress the known percentage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Session known-word timeline API exposes cumulative filtered total word counts alongside known counts.
- [x] #2 Session detail chart computes known/unknown areas from filtered totals, not raw timeline token counts, when known-word data is available.
- [x] #3 Session summary known-word rate uses filtered persisted word totals where available and preserves safe fallback behavior when known-word data is unavailable.
- [x] #4 Regression tests cover filtered denominator behavior for the API and chart data path.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented in-place fix using existing persisted word occurrence rows. `/api/stats/sessions/:id/known-words-timeline` now returns cumulative `totalWordsSeen` from filtered persisted occurrences, and session known-word rates divide by the same filtered total. Session detail chart builds known/unknown areas from `totalWordsSeen` instead of raw timeline `tokensSeen`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Known-word percentages on session charts now use filtered persisted word totals for both numerator and denominator. No migration/backfill required; data comes from existing `imm_word_line_occurrences`. Added regression coverage for the API response/rate and chart data builder.
<!-- SECTION:FINAL_SUMMARY:END -->
