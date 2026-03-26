---
id: TASK-182
title: Fix session stats chart known-word totals exceeding total words
status: Done
milestone: m-1
assignee:
  - codex
created_date: '2026-03-17 16:07'
updated_date: '2026-03-18 05:28'
labels: []
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/stats-server.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/stats/src/components/sessions/SessionDetail.tsx
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/__tests__/stats-server.test.ts
  - /Users/sudacode/projects/japanese/SubMiner/stats/src/hooks/useSessions.ts
ordinal: 109500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the session detail stats display so the known-word series cannot exceed the total-word series for the same sample. Ground the fix in the actual immersion-tracker metrics used by the stats UI and cover the regression with automated tests.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Session detail data uses a consistent cumulative word metric so known-word counts do not exceed total words for a sample
- [x] #2 Automated tests cover the session known-word timeline contract and reproduce the regression scenario
- [x] #3 Session stats UI still renders the timeline and tooltip values correctly after the fix
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test for `/api/stats/sessions/:id/known-words-timeline` covering a Japanese-style session where telemetry word counts can be lower than token-derived known-word counts.
2. Update the stats known-word timeline contract/server implementation so the series is expressed in the same cumulative unit used for total words in the session detail view.
3. Adjust the session detail UI/types to consume the corrected series and keep tooltip/legend copy coherent.
4. Run targeted tests for stats server and stats UI transforms, then summarize any wider verification skipped.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented server-side known-word timeline fix to preserve stored line positions and accumulate known-word occurrences rather than compressed unique-headword counts.

Updated session-facing stats views to prefer `tokensSeen` over `wordsSeen` when available so displayed session word totals align with the session chart and lookup-rate denominator.

Verification: `bun test src/core/services/__tests__/stats-server.test.ts`, `bun test stats/src/lib/yomitan-lookup.test.tsx`, `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`, `bun run typecheck` all passed.

Verification skipped/blocker: `bun run typecheck:stats` still fails in pre-existing unrelated files `stats/src/components/anime/AnilistSelector.tsx`, `stats/src/lib/reading-utils.test.ts`, and `stats/src/lib/reading-utils.ts`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the session stats mismatch that let known words outrun total words. The stats server now preserves actual subtitle-line positions and accumulates known-word occurrences for the session timeline, while session-facing stats views prefer token-based word totals when available. Added regression coverage for the known-word timeline API and for session-row word-count rendering, plus a user-visible changelog fragment.
<!-- SECTION:FINAL_SUMMARY:END -->
