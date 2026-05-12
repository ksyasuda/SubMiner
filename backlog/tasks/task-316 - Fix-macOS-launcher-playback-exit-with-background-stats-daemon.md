---
id: TASK-316
title: Fix macOS launcher playback exit with background stats daemon
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-03 00:32'
updated_date: '2026-05-03 00:36'
labels:
  - bug
  - macos
  - mpv
  - stats
  - runtime
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launching a video on macOS when SubMiner is not already running should not leave the regular SubMiner app/tray alive after mpv closes. A separately running background stats daemon must remain non-blocking and must not be used as a foreground app dependency during playback startup/shutdown.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Closing a launcher/plugin-managed mpv session exits the launcher-started regular SubMiner app/tray after mpv closes.
- [x] #2 Explicit background/no-argument app launches still remain alive as before.
- [x] #3 A live background stats daemon is ignored by normal in-app stats server routing during regular app startup/playback, so the regular app never depends on or connects to that background daemon.
- [x] #4 Regression coverage demonstrates the managed playback shutdown and stats-daemon isolation behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regressions first: stats routing should ignore a live foreign background daemon for normal app URL/server startup, and managed playback disconnect should request app quit directly without reconnecting or depending on overlay/youtube disconnect guards.
2. Implement the narrow runtime changes in `src/main/runtime/stats-server-routing.ts` and, if needed, mpv disconnect plumbing in `src/core/services/mpv.ts` / event deps.
3. Preserve explicit persistent background/no-arg behavior by keeping `--managed-playback` as the only playback-exit marker.
4. Run focused tests (`stats-server-routing`, mpv client/protocol/event tests), then typecheck if focused checks pass.
5. Update changelog and task acceptance/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented regular app stats routing isolation from live background daemon state and explicit managed-playback quit-on-disconnect wiring in main mpv event deps. Existing `MpvIpcClient` socket-close managed playback quit path remains covered.

`bun run test:fast` was attempted after focused verification. It failed in the broad `test:core:src` lane with Bun/node:test nested-test runner errors across many unrelated files and one transient subsync renderer API failure; rerunning the concrete subsync failure alone passed. Focused runtime tests, typecheck, and changelog lint remain green.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Regular app stats server routing no longer returns or depends on a live background daemon URL; it validates/cleans state, then uses the local app stats server path.
- Managed playback is now explicitly treated as a quit-on-disconnect launch mode in main mpv event deps, in addition to the existing mpv socket-close quit request.
- Added regressions for background daemon isolation and managed playback quit-on-disconnect classification.
- Added changelog fragment `changes/316-macos-playback-stats-daemon.md`.

Verification:
- `bun test src/main/runtime/stats-server-routing.test.ts src/core/services/mpv.test.ts src/core/services/mpv-protocol.test.ts src/main/runtime/mpv-client-event-bindings.test.ts src/main/runtime/mpv-main-event-bindings.test.ts src/main/runtime/mpv-main-event-main-deps.test.ts`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun test src/core/services/subsync.test.ts --test-name-pattern "deterministic _retimed"`

Blocked broader gate:
- `bun run test:fast` failed in `test:core:src` with Bun/node:test nested-test runner errors across unrelated files; the concrete subsync failure from that run passed when isolated.
<!-- SECTION:FINAL_SUMMARY:END -->
