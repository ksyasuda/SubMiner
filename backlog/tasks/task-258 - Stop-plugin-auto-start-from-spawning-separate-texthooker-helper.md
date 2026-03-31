---
id: TASK-258
title: Stop plugin auto-start from spawning separate texthooker helper
status: Done
assignee:
  - codex
created_date: '2026-03-30 06:25'
updated_date: '2026-03-31 19:37'
labels:
  - bug
  - texthooker
  - launcher
  - plugin
  - startup
dependencies: []
priority: high
ordinal: 158500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Change the mpv/plugin auto-start path so normal SubMiner startup owns texthooker and websocket startup inside the main `--start` app instance. Keep standalone `subminer texthooker` / plain `--texthooker` available for explicit external use, but stop the plugin from spawning a second helper subprocess during regular auto-start.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Plugin auto-start includes texthooker on the main `--start` command when texthooker is enabled.
- [x] #2 Plugin auto-start no longer spawns a separate standalone `--texthooker` helper subprocess during normal startup.
- [x] #3 Regression coverage fails before the fix and passes after the fix for the plugin auto-start path.
- [x] #4 Standalone external `subminer texthooker` / plain `--texthooker` entrypoints remain available for explicit helper use.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Flip the mpv/plugin start-gate regression so enabled texthooker is folded into the main `--start` command and standalone helper subprocesses are rejected.
2. Update plugin process command construction so `start` includes `--texthooker` when enabled and the separate helper-launch path becomes a no-op for normal auto-start.
3. Run plugin Lua regressions, adjacent launcher tests, and typecheck to verify behavior and preserve explicit standalone `--texthooker` entrypoints.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Design approved by user: normal in-app startup should own texthooker/websocket; `texthookerOnlyMode` should stay explicit external-only.

Root cause path: mpv/plugin auto-start in `plugin/subminer/process.lua` launched `binary_path --start ...` and then separately spawned `binary_path --texthooker --port ...`. That created the standalone helper process observed live (`SubMiner --texthooker --port 5174`) instead of relying on the normal app instance.

Fix: `build_command_args('start', overrides)` now appends `--texthooker` when texthooker is enabled, and the old helper-launch path is reduced to a no-op so normal auto-start remains single-process.

Verification: `lua scripts/test-plugin-start-gate.lua`, `lua scripts/test-plugin-process-start-retries.lua`, `bun test launcher/mpv.test.ts launcher/commands/playback-command.test.ts launcher/config/args-normalizer.test.ts`, and `bun run typecheck`. Standalone launcher/app entrypoints for explicit `subminer texthooker` / plain `--texthooker` were left untouched.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stopped the mpv/plugin auto-start path from spawning a second standalone texthooker helper. Texthooker now rides on the main `--start` app instance for normal startup, with Lua regressions updated to require `--texthooker` on the main start command and reject separate helper subprocesses. Explicit standalone `subminer texthooker` / plain `--texthooker` entrypoints remain available.
<!-- SECTION:FINAL_SUMMARY:END -->
