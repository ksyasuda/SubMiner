---
id: TASK-218
title: 'Delete zero-session media from stats library and trends'
status: Done
assignee:
  - codex
created_date: '2026-03-22 16:20'
updated_date: '2026-03-22 21:10'
labels:
  - stats
  - immersion-tracker
priority: medium
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/query.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/lifetime.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/maintenance.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/__tests__/query.test.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Deleting the last retained session for a video still left stale lifetime media rows and trend rollups behind, so the stats dashboard could continue showing ghost entries in Library and Trends after all sessions were gone.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Deleting the final session for a video removes that media from Library queries and detail reads
- [x] #2 Deleting the final session for a video removes stale daily/monthly trend rollups for that media
- [x] #3 Regression coverage proves zero-session media disappears from affected stats surfaces after deletion
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression around deleting the only retained session for a video while preexisting lifetime and rollup rows exist.
2. Patch the deletion path to rebuild lifetime and rollup state from retained sessions inside the same transaction.
3. Run focused immersion-tracker tests plus the repo-native verifier core lane and record results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added a query regression that seeds a finished session plus stale lifetime media/anime rows and daily/monthly rollups, deletes that only session, and asserts Library, Anime detail, and Trends all drop the media immediately.

Refactored lifetime rebuild logic so it can run inside an existing delete transaction, then reused that helper from `deleteSession`, `deleteSessions`, and `deleteVideo`.

Added a rollup rebuild helper that clears existing daily/monthly rollups and reconstructs them from retained telemetry inside the current transaction so deleted sessions cannot leave ghost trend points behind.

Verification passed:
- `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
- `bun test src/core/services/immersion-tracker-service.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/lifetime.ts src/core/services/immersion-tracker/maintenance.ts src/core/services/immersion-tracker/__tests__/query.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/core/services/immersion-tracker/query.ts src/core/services/immersion-tracker/lifetime.ts src/core/services/immersion-tracker/maintenance.ts src/core/services/immersion-tracker/__tests__/query.test.ts`

Verifier artifact dir: `.tmp/skill-verification/subminer-verify-20260322-210718-n6sGL8`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Delete paths now rebuild lifetime summaries and trend rollups after removing sessions, so when the last session for a video disappears the stats database also drops that media from Library, related detail reads, and chart data. Added a regression proving a video with only stale lifetime/rollup rows vanishes after its final session is deleted, and verified the change with focused immersion-tracker tests plus the SubMiner core verification lane.
<!-- SECTION:FINAL_SUMMARY:END -->
