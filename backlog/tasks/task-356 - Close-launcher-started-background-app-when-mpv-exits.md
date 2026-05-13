---
id: TASK-356
title: Close launcher-started background app when mpv exits
status: Done
assignee:
  - codex
created_date: '2026-05-13 01:37'
updated_date: '2026-05-13 01:40'
labels:
  - bug
  - launcher
  - mpv
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When SubMiner is started through the launcher-managed mpv flow, closing the mpv window should also close the background Electron app instead of leaving it running in the tray. Preserve intentional tray/background behavior for normal app startup.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher-managed mpv sessions signal or otherwise cause the spawned background app to quit when the mpv process exits.
- [x] #2 Normal background/tray startup remains available when SubMiner is launched without a launcher-managed playback session.
- [x] #3 A regression test covers the launcher mpv close/shutdown behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a launcher command regression test for mpv plugin auto-start playback: no direct startOverlay call, mpv exits, launcher marks the session as managed and runs cleanup.
2. Add a small launcher mpv lifecycle helper to mark a SubMiner app session as launcher-managed when the launcher relies on plugin auto-start.
3. Wire playback-command to call that helper only for launcher-managed playback paths where mpv plugin auto-start is expected.
4. Run the focused launcher tests, then update TASK-356 acceptance criteria/notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented launcher ownership marking for plugin-auto-start playback sessions. Direct startOverlay already marks launcher ownership; the plugin-auto-start branch now does the same before waiting for mpv exit, so existing cleanup sends the app --stop when mpv closes. Added regression coverage in launcher/commands/playback-command.test.ts. Verification: bun test launcher/commands/playback-command.test.ts; bun test launcher/mpv.test.ts launcher/commands/playback-command.test.ts; bun run test:launcher:src; bun run typecheck. Typecheck initially caught a nullable test fixture assignment and passed after fixing it.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added markOverlayManagedByLauncher() to centralize launcher ownership tracking for SubMiner app sessions.
- Mark plugin-auto-start playback sessions as launcher-managed, so closing mpv triggers existing cleanup and stops the background app instead of leaving it in the tray.
- Added a regression test covering mpv exit after launcher-managed plugin auto-start playback.

Tests:
- bun test launcher/commands/playback-command.test.ts
- bun test launcher/mpv.test.ts launcher/commands/playback-command.test.ts
- bun run test:launcher:src
- bun run typecheck
<!-- SECTION:FINAL_SUMMARY:END -->
