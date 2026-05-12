---
id: TASK-346
title: Fix stats session detail when recent media is missing
status: Done
assignee:
  - Codex
created_date: '2026-05-12 06:41'
updated_date: '2026-05-12 06:44'
labels:
  - bug
  - stats
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Stats overview can show a completed session, but clicking it opens a detail view that says "Media not found". The details view should resolve the session/media consistently for recently completed local playback so users can inspect session cards, words, timeline, and media stats after a video finishes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A completed session listed on the stats overview opens a usable details view instead of "Media not found" when its media is still present in stats data.
- [x] #2 Regression coverage reproduces the overview-to-detail lookup mismatch and verifies the corrected behavior.
- [x] #3 Relevant stats/detail documentation is updated if behavior or APIs change.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression test in `src/core/services/immersion-tracker/__tests__/query.test.ts` covering a video/session visible from session summaries before `imm_lifetime_media` exists.
2. Update `getMediaDetail` in `src/core/services/immersion-tracker/query-library.ts` so detail rows can resolve from `imm_videos` plus session metrics when lifetime summary is absent.
3. Run the focused query test lane and update task notes/acceptance criteria.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented root-cause fix in `getMediaDetail`: media detail now resolves from `imm_videos` plus session metrics when `imm_lifetime_media` is not populated yet, while still preferring lifetime summary totals when available. Added regression test for session-visible/media-detail-missing mismatch. No docs update required because the API shape is unchanged; added changelog fragment `changes/346-stats-session-detail.md`. Verification: `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`, `bun run typecheck`, `bun run format:check:src`, `bun run test:fast`, `bun run changelog:lint`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Fixed stats media detail lookup so sessions visible in overview can open detail even before lifetime media summaries exist.
- Preserved lifetime-summary totals when available and added a regression test for the missing-lifetime case.
- Added `changes/346-stats-session-detail.md` for the user-visible fix.

Tests:
- `bun test src/core/services/immersion-tracker/__tests__/query.test.ts`
- `bun run typecheck`
- `bun run format:check:src`
- `bun run test:fast`
- `bun run changelog:lint`
<!-- SECTION:FINAL_SUMMARY:END -->
