---
id: TASK-301
title: Fix launcher-managed video close leaving background app alive
status: Done
assignee:
  - Codex
created_date: '2026-04-26 03:29'
updated_date: '2026-04-26 03:33'
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
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Launcher-managed video playback now exits the SubMiner background app/tray when mpv shuts down, avoiding a leftover background process after closing a launcher-started video. Explicit `--background` and no-argument app startup remain persistent because the quit path is gated on the new `--managed-playback` flag. The mpv plugin passes that flag for auto-started playback, and mpv protocol shutdown requests app quit only in that managed mode.

Verification: plugin start-gate regression coverage, mpv protocol shutdown regression coverage, CLI managed-playback parse coverage, plus broader local gates run before handoff.
<!-- SECTION:FINAL_SUMMARY:END -->
