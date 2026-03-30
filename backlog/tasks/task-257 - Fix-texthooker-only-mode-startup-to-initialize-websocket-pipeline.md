---
id: TASK-257
title: Fix texthooker-only mode startup to initialize websocket pipeline
status: Done
assignee:
  - codex
created_date: '2026-03-30 06:15'
updated_date: '2026-03-30 06:17'
labels:
  - bug
  - texthooker
  - websocket
  - startup
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix `--texthooker` / `subminer texthooker` startup so it launches the texthooker page without the overlay window but still initializes the runtime pieces required for live subtitle delivery. Today texthooker-only mode serves the page yet skips mpv client and websocket startup, leaving the page pointed at `ws://127.0.0.1:6678` with no listener behind it.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `--texthooker` mode starts the texthooker page without opening the overlay window and still initializes the websocket path needed for live subtitle delivery.
- [x] #2 Texthooker-only startup creates the mpv/websocket runtime needed for the configured annotation or subtitle websocket feed.
- [x] #3 Regression coverage fails before the fix and passes after the fix for texthooker-only startup.
- [x] #4 Docs/help text remain accurate for texthooker-only behavior; update docs only if wording needs correction.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Replace the existing texthooker-only startup regression test so it asserts websocket/mpv startup still happens while overlay window initialization stays skipped.
2. Remove or narrow the early texthooker-only short-circuit in app-ready startup so runtime config, mpv client, subtitle websocket, and annotation websocket still initialize.
3. Run focused tests plus a local process check proving `--texthooker` now opens the websocket listener expected by the served page.
4. Update task notes/final summary with the live-process root cause (`--texthooker` serving HTML on 5174 with no 6678 listener).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Live-process repro on the user's machine: `ps` showed the active process as `/tmp/.mount_SubMin.../SubMiner --texthooker --port 5174`. `lsof` showed 5174 listening but no listener on 6678/6677, while `curl http://127.0.0.1:5174/` confirmed the served page was correctly bootstrapped to `ws://127.0.0.1:6678`. That proved the remaining failure was startup mode, not page injection.

Root cause: `runAppReadyRuntime(...)` had an early `texthookerOnlyMode` return that reloaded config and handled initial args, but skipped `createMpvClient()`, subtitle websocket startup, annotation websocket startup, subtitle timing tracker creation, and the later texthooker-only branch that only skips the overlay window.

Fix: removed the early texthooker-only short-circuit so texthooker-only mode now runs the normal startup pipeline, then falls through to the existing `Texthooker-only mode enabled; skipping overlay window.` branch.

Verification: `bun run typecheck`; focused Bun tests for app-ready startup, startup bootstrap, CLI texthooker startup, and CLI context wiring. Existing local live-binary repro still reflects the old mounted AppImage until rebuilt/restarted. Current-binary workaround is to launch normal startup / `--start --texthooker` instead of plain `--texthooker`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the second texthooker regression: plain `--texthooker` mode was serving the page but skipping mpv/websocket initialization, so the page pointed at `ws://127.0.0.1:6678` with no listener. Removed the early texthooker-only startup return, kept the later overlay-skip behavior, updated the startup regression test to require websocket/mpv initialization in texthooker-only mode, and re-verified with typecheck plus focused test coverage.
<!-- SECTION:FINAL_SUMMARY:END -->
