---
id: TASK-185
title: Clarify library stats word-count labels
status: Done
assignee:
  - codex
created_date: '2026-03-17 22:58'
updated_date: '2026-03-18 05:28'
labels:
  - bug
  - stats
  - ui
milestone: m-1
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/library/MediaHeader.tsx
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/stats-server.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/query.ts
priority: medium
ordinal: 104500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the library/media detail stats header so occurrence-based subtitle counts are not presented as unique-word vocabulary totals. The UI should clearly distinguish subtitle word occurrences from unique known-word headword coverage to avoid misleading comparisons.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Library/media detail view labels subtitle occurrence totals with wording that does not imply unique vocabulary counts
- [x] #2 Known-words summary in the same view explicitly communicates that its denominator is unique words/headwords
- [x] #3 Frontend tests cover the updated copy so the mismatch does not regress
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused stats frontend test for MediaHeader copy that distinguishes occurrence totals from unique known-word coverage.
2. Update MediaHeader labels so occurrence-based totals no longer imply unique vocabulary counts.
3. Update the known-words label copy to explicitly state it is based on unique words/headwords.
4. Run targeted stats tests and record results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause confirmed: library header compares occurrence-based totalWordsSeen against unique-headword known-words summary. Awaiting plan approval before code changes.

Updated library header copy to label totalWordsSeen as word occurrences and known-word coverage as known unique words. Added an optional initialKnownWordsSummary prop to support deterministic server-render tests without changing runtime behavior.

Verification: `bun test stats/src/lib/yomitan-lookup.test.tsx` passes. `bun run typecheck:stats` remains blocked by preexisting unrelated errors in stats/src/components/anime/AnilistSelector.tsx, stats/src/lib/reading-utils.ts, stats/src/lib/reading-utils.test.ts, and stats/src/lib/vocabulary-tab.test.ts.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Clarified the library/media header so occurrence-based subtitle counts are no longer presented as if they were unique vocabulary totals. The header now labels `totalWordsSeen` as `word occurrences`, and the known-words summary explicitly says `known unique words`, which matches the backend's DISTINCT headword calculation.

For regression coverage, added a focused MediaHeader render test that exercises the exact mismatch case (30 occurrences vs 34 unique words) and verifies the new copy. Also updated one stale AnimeOverviewStats assertion in the same targeted test file so the focused stats test lane is green.

Tests run:
- `bun test stats/src/lib/yomitan-lookup.test.tsx` ✅
- `bun run typecheck:stats` ⚠️ blocked by preexisting unrelated errors in AnilistSelector and reading-utils/vocabulary-tab stats files.
<!-- SECTION:FINAL_SUMMARY:END -->
