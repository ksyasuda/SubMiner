---
id: TASK-326
title: Make stats word metrics honor filtering rules
status: Done
assignee: []
created_date: '2026-05-04 01:35'
updated_date: '2026-05-04 02:08'
labels:
  - stats
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Audit stats app metrics that show or derive from word totals and make them use filtered persisted vocabulary occurrences where the UI concept is learned/seen words. Preserve raw telemetry only where it is intentionally playback/token telemetry.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Stats UI word totals, word rates, lookup-per-word rates, and chart word series use filtered persisted word occurrences where available.
- [x] #2 Known-word metrics continue to use the same filtered denominator as known counts.
- [x] #3 Trend, overview, library, session, and episode surfaces are audited with regression coverage for changed data paths.
- [x] #4 Fallback behavior remains safe for sessions without persisted vocabulary occurrences.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Audit finding: raw `tokensSeen` / `totalTokensSeen` still feeds overview hints, dashboard aggregation, trends activity/progress/anime cumulative/library summary, lookup-per-100-word rates, session rows/recent sessions/episode sessions, and library/anime/media headers. Vocabulary and known unique word summaries already use persisted filtered vocabulary rows. Recommended design: query-time filtered word totals from existing `imm_word_line_occurrences`, with raw-token fallback only when a session has no persisted occurrence rows.

Implemented shared query-time filtered word counts. Session summaries, overview hints, daily/monthly rollups, anime/media library/detail rows, anime episode rows, episode/media sessions, trends activity/progress/anime cumulative, library summary, and lookup-per-100-word ratios now use filtered persisted word occurrences. Fallback remains raw token totals only for sessions with no persisted subtitle-line rows.

Follow-up implemented: Vocab frequency tables now apply the same tokenizer vocabulary predicate at read time, because old `imm_words` rows can predate current tokenizer exclusion rules. Vocabulary persistence and cleanup also mirror the broader subtitle-annotation grammar filters. Added common frequency stop terms observed in the stats vocabulary list to the shared tokenizer exclusion set so those rows are filtered consistently across subtitle annotations, persistence, cleanup, stats reads, and SQL word-count aggregates.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stats word metrics now honor filtering rules through the read-model query layer. Existing persisted `imm_word_line_occurrences` provide the filtered denominator; no migration/backfill needed. Vocab tables filter stored rows on read using tokenizer vocabulary rules, so legacy noisy rows stop appearing without a migration. Added regressions for session/overview/rollup fallback behavior, trends/library lookup-rate behavior, vocabulary read filtering, cleanup filtering, and shared stop-term filtering.
<!-- SECTION:FINAL_SUMMARY:END -->
