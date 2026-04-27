---
id: TASK-306
title: Separate background stats daemon from regular SubMiner app
status: Done
assignee: []
created_date: '2026-04-27 00:56'
updated_date: '2026-04-27 01:00'
labels:
  - stats
  - runtime
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Background stats mode should run only the stats data/server pieces. It must not bring up tray UI or expose the regular mpv connection surface, and stopping should remain CLI-only.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launching stats background mode starts a separate stats daemon process rather than booting the regular SubMiner runtime.
- [x] #2 Background stats mode does not create or keep a tray icon.
- [x] #3 Background stats mode does not start mpv IPC/client surfaces that let mpv connect to the app.
- [x] #4 Background stats mode remains stoppable through the stats stop command line path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add entry-runtime tests for public stats background/stop daemon detection.
2. Implement early public stats daemon command detection and route it before regular app boot.
3. Run targeted tests and update task status/criteria.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented early public stats daemon routing in main-entry runtime. Direct `--stats-background` and `--stats-stop` now resolve to daemon control before single-instance lock and before loading `main.js`, matching the existing internal launcher daemon flags. Installed missing Bun dependencies to run targeted tests.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added `resolveStatsDaemonCommandAction` and updated entry detection so public `--stats-background` / `--stats-stop` invocations route through the isolated stats daemon control path.
- Reused that action resolution in `stats-daemon-entry` so public stop commands map to stop instead of the default start path.
- Added regression coverage for public daemon detection/action resolution.

Verification:
- `bun test src/main-entry-runtime.test.ts launcher/commands/command-modules.test.ts src/main/runtime/stats-cli-command.test.ts src/stats-daemon-control.test.ts`
- `bun run typecheck`
<!-- SECTION:FINAL_SUMMARY:END -->
