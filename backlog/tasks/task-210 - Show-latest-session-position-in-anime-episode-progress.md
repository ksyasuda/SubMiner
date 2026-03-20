---
id: TASK-210
title: Show latest session position in anime episode progress
status: Done
assignee:
  - '@Codex'
created_date: '2026-03-20 04:09'
updated_date: '2026-03-20 04:25'
labels:
  - stats
  - bug
  - ui
milestone: m-1
dependencies: []
references:
  - stats/src/components/anime/EpisodeList.tsx
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker/session.ts
  - src/core/services/immersion-tracker-service.ts
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Anime episode rows in stats can show watch time and lookups from the latest session while the Progress column stays blank because it only reads `ended_media_ms` from ended sessions. Update the progress source so a just-watched episode reflects the latest known session stop position without falling back to cumulative watch time.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Anime episode progress uses the latest known session position for the episode, including the most recent active session when available.
- [x] #2 Ended-session progress remains correct and does not regress to cumulative watch time.
- [x] #3 Regression coverage locks query and/or UI behavior for active-session and ended-session episode progress.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression coverage for anime episode progress when the latest session is still active but has a known playback position.
2. Persist the latest playback position on the active `imm_sessions` row during playback so stats queries can read it before session finalization.
3. Update anime episode queries to use the newest known session position for progress while preserving ended-session behavior.
4. Run targeted verification for immersion tracker, stats query, and cheap repo checks; record results and task outcome.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause: stale active-session recovery rebuilt session state with `lastMediaMs = null`, so `finalizeSessionRecord` overwrote persisted progress checkpoints with `ended_media_ms = NULL` during startup reconciliation.

Implemented telemetry-flush checkpointing to persist `lastMediaMs` onto the active `imm_sessions` row, preserved that checkpoint through stale-session reconciliation, and updated anime episode progress queries to read the latest known non-null session position across active or ended sessions.

Verification: targeted regressions passed (`bun test src/core/services/immersion-tracker-service.test.ts --test-name-pattern 'flushTelemetry checkpoints latest playback position on the active session row|startup finalizes stale active sessions and applies lifetime summaries'`, `bun test src/core/services/immersion-tracker/__tests__/query.test.ts --test-name-pattern 'getAnimeEpisodes prefers the latest session media position when the latest session is still active|getAnimeEpisodes returns latest ended media position and aggregate metrics'`), broader tracker/query suite passed (`bun test src/core/services/immersion-tracker-service.test.ts src/core/services/immersion-tracker/__tests__/query.test.ts`), `bun run typecheck` passed via verifier, `bun run changelog:lint` passed.

Verification blocker: `.agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core ...` reported `bun run test:fast` failure from pre-existing `scripts/update-aur-package.test.ts` (`mapfile: command not found` under bash), unrelated to this change set.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Persist anime episode progress checkpoints before session finalization so stats can survive crashes/restarts and still show the latest known watch position. Telemetry flushes now checkpoint `lastMediaMs` onto the active `imm_sessions` row, stale-session recovery preserves that checkpoint when finalizing recovered sessions, and `getAnimeEpisodes` now reads the newest non-null session position whether it came from an active or ended session.

Added regressions for active-session checkpoint persistence, stale-session recovery preserving `ended_media_ms`, and episode queries preferring the latest known session position. Verification passed for the targeted and broader immersion tracker/query suites, plus `bun run typecheck` and `bun run changelog:lint`. The verifier's `bun run test:fast` step still fails on the pre-existing `scripts/update-aur-package.test.ts` bash `mapfile` issue, which is outside this task's scope.
<!-- SECTION:FINAL_SUMMARY:END -->
