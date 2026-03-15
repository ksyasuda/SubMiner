---
id: TASK-171
title: Add normalized immersion word and kanji occurrence tracking
status: Done
assignee:
  - codex
created_date: '2026-03-14 11:30'
updated_date: '2026-03-14 11:48'
labels:
  - immersion
  - stats
  - database
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-14-immersion-occurrence-tracking-design.md
  - /home/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-14-immersion-occurrence-tracking.md
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add normalized occurrence tables for immersion-tracked words and kanji so stats can map vocabulary back to the exact anime, episode, timestamp, and subtitle line where each item appeared. Preserve repeated tokens within the same line via counted occurrences instead of deduping, while avoiding duplicated token text storage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 The immersion schema adds normalized subtitle-line and counted occurrence tables for words and kanji, with additive migration support for existing databases.
- [x] #2 Subtitle-line tracking writes one subtitle-line row per seen line plus counted word/kanji occurrences linked back to the line, session, video, and anime context.
- [x] #3 Query surfaces can map a word or kanji back to anime/episode/line/timestamp rows without breaking current top-level vocabulary and kanji stats.
- [x] #4 Focused regression coverage exists for schema, counted occurrence persistence, and reverse-mapping queries.
- [x] #5 Verification covers the SQLite immersion lane and any broader lanes required by touched service/API files.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add red tests for new line/occurrence schema and migration shape in the SQLite immersion lane.
2. Add red tests for service-level subtitle persistence that writes one line row plus counted word/kanji occurrences.
3. Implement additive schema, write-path plumbing, and counted occurrence upserts with minimal disruption to existing aggregate tables.
4. Add reverse-mapping query surfaces for word and kanji occurrences, plus focused API/service exposure only where needed.
5. Run focused SQLite verification first, then broader verification only if touched runtime/API files require it.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-14: Design approved in-thread. Chosen shape: `imm_subtitle_lines` plus counted bridge tables `imm_word_line_occurrences` and `imm_kanji_line_occurrences`, retaining repeated tokens within a line via `occurrence_count`.
2026-03-14: Implemented additive schema version bump to 7. `recordSubtitleLine(...)` now queues one normalized subtitle-line write that owns aggregate word/kanji upserts plus counted bridge-row inserts.
2026-03-14: Added reverse-mapping query surfaces for exact word triples and single kanji lookups. No stats API/UI consumer was widened in this change.
2026-03-14: Verification commands run:
  - `bun test src/core/services/immersion-tracker-service.test.ts`
  - `bun test src/core/services/immersion-tracker/storage-session.test.ts`
  - `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
  - `bun run typecheck`
  - `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker-service.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/__tests__/query.test.ts`
  - `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker-service.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/__tests__/query.test.ts`
  - `bun run test:immersion:sqlite:src`
2026-03-14: Verification results:
  - targeted tracker/query tests: passed
  - verifier lane selection: `core`
  - verifier result: passed (`typecheck`, `test:fast`)
  - verifier artifacts: `.tmp/skill-verification/subminer-verify-20260314-114630-abO7mb/`
  - maintained immersion SQLite lane: passed
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added normalized subtitle-line occurrence tracking to immersion stats with three additive tables: `imm_subtitle_lines`, `imm_word_line_occurrences`, and `imm_kanji_line_occurrences`.

`recordSubtitleLine(...)` now preserves repeated allowed tokens and repeated kanji within the same subtitle line via `occurrence_count`, while still updating canonical `imm_words` and `imm_kanji` aggregates.

Added reverse-mapping queries for exact word triples and kanji so callers can fetch anime/video/session/line/timestamp context for each occurrence without duplicating token text storage.

Verified with targeted tracker/query tests, `bun run typecheck`, verifier-selected `core` coverage, and the maintained `bun run test:immersion:sqlite:src` lane.
<!-- SECTION:FINAL_SUMMARY:END -->
