---
id: TASK-178
title: 'Address PR #19 Codex review feedback on immersion session deletion'
status: Done
assignee:
  - codex
created_date: '2026-03-17 14:59'
updated_date: '2026-03-18 05:28'
labels:
  - pr-review
  - immersion-tracker
  - stats
milestone: m-1
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker-service.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/query.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker-service.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/immersion-tracker/__tests__/query.test.ts
priority: medium
ordinal: 113500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess the open Codex review items on PR #19 and fix verified deletion-path regressions in immersion tracking so dashboard deletes cannot corrupt tracker state or leave stale aggregate stats.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Deleting the active immersion session is rejected safely and does not leave the tracker in a flush-failure loop
- [x] #2 Deleting sessions rebuilds or updates vocabulary and kanji aggregates so stats no longer include removed session data
- [x] #3 Regression tests cover the active-session deletion guard and aggregate cleanup after session deletion
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression tests for deleting the active session through ImmersionTrackerService and for deleteSession/deleteSessions keeping imm_words and imm_kanji aggregates in sync after rows are removed.
2. Verify the failures are caused by the current deletion path, then patch the service guard and query-layer aggregate maintenance with the smallest safe change.
3. Re-run focused tests for the touched files, then run SubMiner verification lanes appropriate for core/runtime-compat changes and record results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verified Codex PR #19 findings against current code: active-session/video deletes could orphan the live tracker session, and deleteSession/deleteSessions/deleteVideo left imm_words/imm_kanji aggregates stale after subtitle-line removal.

Implemented service guards that ignore deletes targeting the active session or active video and log a warning instead of deleting live tracker rows.

Updated query-layer delete helpers to capture affected word/kanji ids before deletion, remove session/video rows in a transaction, then recompute surviving imm_words/imm_kanji frequency and first/last-seen values from remaining subtitle-line occurrences, deleting orphan aggregate rows when no occurrences remain.

Focused verification passed: bun test src/core/services/immersion-tracker-service.test.ts and bun test src/core/services/immersion-tracker/__tests__/query.test.ts.

SubMiner verifier: classify_subminer_diff.sh selected lane core; verify_subminer_change.sh passed typecheck and failed on unrelated existing launcher test `stats command tolerates slower dashboard startup before timing out` in launcher/main.test.ts (timeout waiting for dashboard startup response).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the open Codex PR #19 review items on immersion deletion paths and fixed the confirmed regressions. ImmersionTrackerService now ignores delete requests that target the currently active session or its active video, preventing the dashboard from deleting the live parent rows that subsequent telemetry/event flushes still depend on. On the query side, session/video deletion now captures affected vocabulary and kanji aggregate ids before removing subtitle/session rows, then recomputes imm_words and imm_kanji frequency plus first/last seen timestamps from surviving line occurrences inside the same transaction, deleting orphan aggregate rows when no occurrences remain.

Regression coverage was added for active-session delete protection, active-video delete protection, and aggregate rebuild after session deletion. Focused verification passed with `bun test src/core/services/immersion-tracker-service.test.ts` and `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`. Repo-native verification selected the `core` lane; `bun run typecheck` passed, while `bun run test:fast` failed in an unrelated launcher test (`launcher/main.test.ts`: `stats command tolerates slower dashboard startup before timing out`) that times out waiting for dashboard startup response.
<!-- SECTION:FINAL_SUMMARY:END -->
