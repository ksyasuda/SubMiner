---
id: TASK-301
title: Fix launcher-managed video close leaving background app alive
status: Done
assignee:
  - Codex
created_date: '2026-04-26 03:29'
updated_date: '2026-04-26 03:44'
labels:
  - bug
  - launcher
  - mpv
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launcher/plugin-managed video playback should not leave the SubMiner background app or tray icon running after the video closes unless the user explicitly launched SubMiner in background mode via --background or by starting with no app arguments. This is a regression after crash-avoidance work that added background startup for launcher-managed playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Closing a launcher-managed video exits the launcher-started SubMiner app/tray instead of leaving it alive.
- [x] #2 Explicit background launches still keep SubMiner alive after windows close.
- [x] #3 No-argument app startup behavior remains unchanged.
- [x] #4 Regression coverage exercises the launcher-managed playback shutdown lifecycle.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression coverage first: plugin auto-start should tag launcher-managed playback, and app mpv shutdown handling should quit only when started in that managed playback mode.
2. Add a narrow CLI flag/state field for launcher-managed playback, separate from explicit persistent background mode.
3. Have plugin pass the new flag with its background start command.
4. On mpv shutdown/disconnect, request app quit only when managed playback mode is active; preserve explicit --background and no-arg startup persistence.
5. Run focused plugin/app tests, then relevant launcher/core gates if feasible.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented managed playback shutdown by adding a `--managed-playback` app flag that the mpv plugin passes only for launcher-managed starts. The main mpv shutdown path now quits the app when initial args indicate managed playback, while explicit background/no-arg startup remains persistent. Added plugin start-gate and mpv protocol regression coverage.

Implemented managed playback lifecycle: mpv plugin auto-start passes --background --managed-playback; app quits on mpv shutdown only when initial args include managedPlayback. Explicit --background and no-arg startup remain persistent. Installed updated mpv plugin to ~/.config/mpv/scripts/subminer via make install-plugin.

Retest showed tray still remained. Root cause: relying on mpv's JSON IPC shutdown event was insufficient; the app may only see the socket close. Added managed-playback quit on MpvIpcClient onClose before reconnect scheduling, with regression coverage.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Launcher-managed playback now starts SubMiner with an internal --managed-playback marker alongside --background. The app requests quit either when mpv sends shutdown or when the mpv IPC socket closes, but only for managed playback mode; explicit background/no-arg startup remains persistent. Added CLI, mpv protocol, mpv socket-close, and plugin regression coverage plus a launcher changelog fragment. Rebuilt the app/launcher and confirmed focused checks, typecheck, build, plugin tests, dist smoke, and formatting.
<!-- SECTION:FINAL_SUMMARY:END -->
