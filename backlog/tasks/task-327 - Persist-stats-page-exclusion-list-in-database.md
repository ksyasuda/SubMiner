---
id: TASK-327
title: Persist stats page exclusion list in database
status: Done
assignee: []
created_date: '2026-05-04 01:39'
updated_date: '2026-05-04 01:49'
labels:
  - feature
  - stats
  - database
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add database-backed persistence for the stats page exclusion list. On first load with the new schema, seed the new table from the existing exclusion list source so existing user choices are preserved. After migration, update database rows whenever the exclusion list is changed or saved so it persists across browser sessions indefinitely.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A new small database table stores stats page exclusion entries.
- [x] #2 First load with the new schema seeds the table from the existing exclusion list source.
- [x] #3 Subsequent exclusion list save/change operations update the database-backed list.
- [x] #4 Regression coverage verifies migration/seed behavior and persistence updates.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented DB-backed stats exclusion list using schema version 18 and new `imm_stats_excluded_words` table. Added read/replace query helpers, service methods, and `/api/stats/excluded-words` GET/PUT routes. Stats frontend now loads exclusions from DB, seeds the empty DB table from legacy `localStorage` on first load, and writes each toggle/restore/clear through the API while keeping localStorage in sync for compatibility. Added focused regression coverage for schema/read-replace, API routes, API client, and frontend bootstrap/update behavior. Verification: `bun run typecheck` passed; `bun test src/core/services/__tests__/stats-server.test.ts stats/src/lib/api-client.test.ts stats/src/hooks/useExcludedWords.test.ts` passed; `bun test src/core/services/immersion-tracker/storage-session.test.ts` passed; `bun run docs:test` passed; `bun run format:check:stats` passed; `bun run changelog:lint` passed. Blocked/unrelated: `bun run typecheck:stats` fails in existing stats files (`AnilistSelector.tsx`, `reading-utils*`, `session-grouping.test.ts`, `yomitan-lookup.test.tsx`); `bun run test:immersion:sqlite:src` fails existing `recordSubtitleLine counts exact Yomitan tokens for session metrics` expected 4 got 3; `bun run docs:build` fails missing `@catppuccin/vitepress/theme/macchiato/mauve.css` import.

Added `src/core/services/__tests__/stats-server.test.ts` and `stats/src/hooks/useExcludedWords.test.ts` to the `test:core:src` allowlist so the new DB exclusion route/client/store regressions run in the maintained fast source lane.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Persisted the stats vocabulary exclusion list in SQLite with new schema version 18 table `imm_stats_excluded_words`. Added backend read/replace helpers and `/api/stats/excluded-words` GET/PUT routes, then wired the stats frontend exclusion store to load DB rows, seed an empty DB from legacy browser localStorage on first load, and update the DB on toggle/restore/clear. Updated docs and added changelog fragment. Focused tests and root typecheck pass; broader stats/docs/sqlite gates are blocked by unrelated existing failures recorded in notes.
<!-- SECTION:FINAL_SUMMARY:END -->
