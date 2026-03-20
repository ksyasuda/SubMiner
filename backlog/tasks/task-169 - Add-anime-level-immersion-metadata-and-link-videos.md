---
id: TASK-169
title: Add anime-level immersion metadata and link videos
status: Done
assignee:
  - codex
created_date: '2026-03-13 19:34'
updated_date: '2026-03-16 05:13'
labels:
  - immersion
  - stats
  - database
  - anilist
milestone: m-1
dependencies: []
references:
  - >-
    /home/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-13-immersion-anime-metadata-design.md
  - >-
    /home/sudacode/projects/japanese/SubMiner/docs/plans/2026-03-13-immersion-anime-metadata.md
ordinal: 20500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add first-class anime metadata to the immersion tracker so stats can group sessions and videos by anime, season, and episode instead of relying only on per-video canonical titles. The new model should deduplicate anime-level metadata across rewatches and multiple files, use guessit-first filename parsing with built-in parser fallback, and create provisional anime rows even when AniList lookup fails.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The immersion schema includes a new anime-level table plus additive video linkage/parsed metadata fields needed for anime, season, and episode stats.
- [x] #2 Media ingest creates or reuses anime rows, stores parsed season/episode metadata on videos, and upgrades provisional anime rows when AniList data becomes available.
- [x] #3 Query surfaces expose anime-level aggregation suitable for library/detail/episode stats without breaking current video/session queries.
- [x] #4 Focused regression coverage exists for schema/storage/query/service behavior, including provisional anime rows and guessit-first parser fallback behavior.
- [x] #5 Verification covers the SQLite immersion lane and any broader lanes required by the touched runtime/query files.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add red tests for the new schema shape in the SQLite immersion lane before changing storage code.
2. Implement `imm_anime` plus additive `imm_videos` metadata fields and focused storage helpers for provisional anime creation and AniList upgrade.
3. Add a guessit-first parser helper with built-in fallback and wire media ingest to persist anime/video metadata during `handleMediaChange(...)`.
4. Add anime-level query surfaces for library/detail/episode aggregation and expose them only where needed.
5. Run focused SQLite verification first, then broader verification lanes only if touched runtime/API files require them.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-03-13: Design approved in-thread. Initial scope excluded migration/backfill work, but implementation was corrected in-thread to add a legacy DB migration/backfill path based on filename parsing.
2026-03-13: Detailed implementation plan written at `docs/plans/2026-03-13-immersion-anime-metadata.md`.
2026-03-13: Task 6 export/API work was intentionally skipped because no current stats API/UI consumer needs the anime query surface yet, and widening the contract would have touched unrelated dirty stats files.
2026-03-13: Verification commands run:
  - `bun test src/core/services/immersion-tracker/storage-session.test.ts`
  - `bun test src/core/services/immersion-tracker/metadata.test.ts`
  - `bun test src/core/services/immersion-tracker-service.test.ts`
  - `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
  - `bun run test:immersion:sqlite:src`
  - `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker/metadata.ts src/core/services/immersion-tracker/metadata.test.ts src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/__tests__/query.test.ts src/core/services/immersion-tracker-service.ts src/core/services/immersion-tracker-service.test.ts`
  - `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/immersion-tracker/storage.ts src/core/services/immersion-tracker/storage-session.test.ts src/core/services/immersion-tracker/metadata.ts src/core/services/immersion-tracker/metadata.test.ts src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/types.ts src/core/services/immersion-tracker/__tests__/query.test.ts src/core/services/immersion-tracker-service.ts src/core/services/immersion-tracker-service.test.ts`
2026-03-13: Verification results:
  - `bun run test:immersion:sqlite:src`: passed
  - verifier lane selection: `core`
  - verifier result: passed (`bun run typecheck`, `bun run test:fast`)
  - verifier artifacts: `.tmp/skill-verification/subminer-verify-20260313-214533-Ciw3L0/`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added `imm_anime`, additive `imm_videos` anime/parser metadata fields, and a legacy migration/backfill path that links existing videos to provisional anime rows from parsed filenames.

Added focused storage helpers for normalized anime identity reuse, later AniList upgrades, and per-video season/episode/parser metadata linking. Media ingest now parses and links anime metadata during `handleMediaChange(...)`.

Added anime-level query surfaces for library/detail/episode aggregation and regression coverage for schema, migration, storage, parser fallback, service ingest wiring, and anime stats queries.

Verified with the focused SQLite lane plus verifier-selected `core` coverage (`typecheck`, `test:fast`). No stats API/UI export was added yet because there is no current consumer for the new anime query surface.
<!-- SECTION:FINAL_SUMMARY:END -->
