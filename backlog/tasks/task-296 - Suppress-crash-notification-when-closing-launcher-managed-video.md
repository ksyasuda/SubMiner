---
id: TASK-296
title: Suppress crash notification when closing launcher-managed video
status: Done
assignee: []
created_date: '2026-04-25 23:12'
updated_date: '2026-04-26 02:44'
labels:
  - bug
  - launcher
  - mpv
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix regression where closing a running mpv video causes SubMiner/Electron service crash notification (`<html><tt>/SubMiner</tt> has encountered a fatal error and was closed.</html>`). Not present on origin/main/v0.12.0 path.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Closing a launcher-managed video stops the overlay/app without desktop crash notification.
- [x] #2 Regression test covers the shutdown path that caused the notification.
- [x] #3 Relevant launcher/app tests pass.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Confirm notification source from local crash metadata and mpv/app logs.
2. Add regression coverage for mpv quit/shutdown lifecycle helper spawning.
3. Update mpv Lua lifecycle to avoid Electron helper subprocesses during quit/shutdown while preserving normal end-file hide behavior.
4. Run plugin tests and changelog lint.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Crash records in ~/.cache/drkonqi show SIGBUS in Electron NetworkService utility process for SubMiner AppImage. mpv log shows shutdown starts two `/home/sudacode/.local/bin/SubMiner.AppImage --hide-visible-overlay` subprocesses and kills them during close. Root cause is mpv plugin spawning Electron control helpers during quit/shutdown.

Follow-up after retest: installed plugin matched source patch and no close-time hide command was spawned. New mpv log shows the initial `/home/sudacode/.local/bin/SubMiner.AppImage --start ...` subprocess remains owned by mpv for the whole playback and is killed when mpv quits. New DrKonqi crash at 2026-04-25 16:44 again shows SIGBUS in Electron NetworkService from that AppImage mount. Need detach the long-lived plugin-launched `--start` app process from mpv.

Second fix: plugin-launched `--start` now includes `--background`, using SubMiner's existing background relaunch path so mpv owns only a short-lived starter process rather than the long-running Electron app. Ran `make install-plugin` so ~/.config/mpv/scripts/subminer now contains both lifecycle and background-start fixes. Full `scripts/test-plugin-start-gate.lua` is currently blocked by an unrelated dirty-worktree primary subtitle bar binding test from TASK-297.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Changed the mpv Lua lifecycle so `shutdown` no longer spawns a `--hide-visible-overlay` helper, and `end-file` skips that helper when mpv reports `reason = "quit"`. Also changed plugin-launched `--start` to include `--background`, so mpv owns only SubMiner's short background launcher process instead of the long-running Electron/AppImage process. This addresses both observed crash sources: close-time helper commands and mpv killing the main SubMiner child process at quit. Installed the updated plugin into `~/.config/mpv/scripts/subminer` with `make install-plugin`, and the user confirmed the latest close no longer produced the notification.

Tests: `lua scripts/test-plugin-start-gate.lua` initially proved the shutdown regression failed before the lifecycle fix; full start-gate is currently affected by other dirty work in this file. Passing checks for this commit: `lua scripts/test-plugin-lua-compat.lua && lua scripts/test-plugin-binary-windows.lua`; `bun run changelog:lint`.
<!-- SECTION:FINAL_SUMMARY:END -->
