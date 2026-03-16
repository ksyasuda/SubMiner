---
id: TASK-172
title: Wire immersion occurrence drilldown into stats API and vocabulary drawer
status: Done
assignee:
  - codex
created_date: '2026-03-14 12:05'
updated_date: '2026-03-16 05:13'
labels:
  - immersion
  - stats
  - ui
dependencies:
  - TASK-171
ordinal: 18500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Expose the new immersion word/kanji occurrence queries through the stats server and add a right-side Vocabulary drawer that shows recent occurrence rows with paging when a word or kanji is clicked.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Stats server exposes word and kanji occurrence endpoints with bounded recent-first paging.
- [x] #2 Stats client/types support loading occurrence pages for a selected word or kanji.
- [x] #3 Vocabulary tab opens a right drawer for the selected word/kanji, shows recent occurrences, and supports loading more.
- [x] #4 Focused regression coverage exists for the server endpoint contract, and the stats UI still typechecks/builds.
- [x] #5 Verification covers the cheapest sufficient backend and stats-UI lanes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-14: Design approved in-thread. Chosen UX: click a word chip or kanji glyph to open a right-side drawer with recent-first occurrences, initial cap 50, plus “Load more”.
2026-03-14: Implemented `/api/stats/vocabulary/occurrences` and `/api/stats/kanji/occurrences` with `limit` + `offset` paging. The drawer uses direct stats HTTP client calls and keeps existing aggregate vocabulary data flow intact.
2026-03-14: Verification commands run:
  - `bun test src/core/services/__tests__/stats-server.test.ts`
  - `bun run typecheck`
  - `cd stats && bun run build`
  - `bun run docs:test`
  - `bun run docs:build`
  - `cd stats && bunx tsc --noEmit`
  - `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/stats-server.ts src/core/services/__tests__/stats-server.test.ts`
2026-03-14: Verification results:
  - stats server endpoint tests: passed
  - root typecheck: passed
  - stats UI production build: passed
  - docs-site test/build: passed
  - `cd stats && bunx tsc --noEmit`: passed after removing stale `hasCoverArt` prop usage in the library stats UI
  - verifier result: passed (`typecheck`, `test:fast`)
  - verifier artifacts: `.tmp/skill-verification/subminer-verify-20260314-120900-J0VvB0/`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Wired occurrence drilldown into the stats server and Vocabulary tab. Words and kanji now open a right-side drawer that loads recent occurrence rows 50 at a time and supports “Load more”.

Added bounded recent-first occurrence endpoints to the stats HTTP API, extended the stats client/type surface, and made word chips plus kanji glyphs selectable with active-state styling.

Updated the immersion-tracking docs to mention vocabulary occurrence drilldown. Verified with focused stats-server tests, root typecheck, stats UI production build, docs-site test/build, the repo verifier core lane, and a direct `stats` package typecheck after removing the stale `MediaHeader` prop mismatch.
<!-- SECTION:FINAL_SUMMARY:END -->
