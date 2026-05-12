---
id: TASK-327
title: Restore stats daemon deferral when launching playback
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-04 01:15'
updated_date: '2026-05-04 01:17'
labels:
  - bug
  - stats
  - runtime
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launching a video while a background stats daemon is already running must not fail with stats.serverPort already in use. Normal in-app stats startup should reuse the live daemon URL instead of binding a second stats server, while preserving managed playback shutdown behavior from TASK-316.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A live background stats daemon from another process causes in-app stats URL resolution to return the daemon URL without starting a local stats server.
- [x] #2 Dead or stale daemon state is removed and local stats startup still works.
- [x] #3 Managed playback shutdown behavior remains covered by existing tests.
- [x] #4 Focused regression tests pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Update `src/main/runtime/stats-server-routing.test.ts` first so a live foreign daemon must return its daemon URL and skip local server startup.
2. Run the focused routing test to confirm the regression fails red.
3. Update `src/main/runtime/stats-server-routing.ts` to return `{ source: 'background' }` for live foreign daemon state, clear stale/self-owned state, and keep local startup fallback unchanged.
4. Run focused stats routing tests plus managed playback tests touched by TASK-316.
5. Update changelog and task acceptance/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented via TDD: first changed `stats-server-routing.test.ts` to require live foreign daemon deferral and observed the expected red failure. Then restored `stats-server-routing.ts` to return the daemon URL with `source: 'background'` when daemon state belongs to a live other process. Stale/dead and self-owned stale cleanup paths remain local fallback.

Verification passed: `bun test src/main/runtime/stats-server-routing.test.ts`; focused runtime suite for stats daemon + TASK-316 managed playback files; `bun run typecheck`; `bun run test:fast`.

`bun run changelog:lint` is blocked by pre-existing unrelated `changes/326-anilist-time-position-post-watch.md` missing valid `type` metadata; `changes/327-stats-daemon-deferral.md` follows the expected fragment format.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Restored in-app stats startup deferral to a live background stats daemon from another process, returning the daemon URL and skipping local stats server binding.
- Kept stale/dead daemon cleanup and local stats startup fallback behavior intact.
- Added a changelog fragment for the restored port-conflict fix.

Verification:
- `bun test src/main/runtime/stats-server-routing.test.ts`
- `bun test src/main/runtime/stats-server-routing.test.ts src/core/services/mpv.test.ts src/core/services/mpv-protocol.test.ts src/main/runtime/mpv-client-event-bindings.test.ts src/main/runtime/mpv-main-event-bindings.test.ts src/main/runtime/mpv-main-event-main-deps.test.ts src/main/runtime/stats-cli-command.test.ts src/stats-daemon-control.test.ts`
- `bun run typecheck`
- `bun run test:fast`

Blocked check:
- `bun run changelog:lint` fails on unrelated pre-existing `changes/326-anilist-time-position-post-watch.md` metadata, not this change.
<!-- SECTION:FINAL_SUMMARY:END -->
