---
id: TASK-181
title: Add background-managed stats server lifecycle commands
status: Done
assignee:
  - codex
created_date: '2026-03-17 15:31'
updated_date: '2026-03-18 05:28'
labels:
  - cli
  - launcher
  - stats
milestone: m-1
dependencies: []
priority: medium
ordinal: 110500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a dedicated background stats server mode that can be started and stopped from the launcher without blocking normal SubMiner instances. Launcher UX: `subminer stats -b` starts the stats server in the background, `subminer stats -s` stops the background stats server only, and plain `subminer stats` preserves the existing foreground/open-browser flow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `subminer stats -b` starts a background stats server without blocking other SubMiner instances.
- [x] #2 `subminer stats -s` stops only the background stats server and succeeds cleanly when state is stale.
- [x] #3 Plain `subminer stats` preserves current dashboard-open behavior.
- [x] #4 Automated tests cover launcher parsing/dispatch and app-side start-stop lifecycle behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Extend launcher stats parsing so `subminer stats -b` maps to background-start and `subminer stats -s` maps to stop-only while preserving existing cleanup/rebuild parsing.
2. Add launcher execution branches: detached background start with startup acknowledgement wait, stop command forwarding with response wait, and preserve existing attached foreground behavior for plain `stats` and cleanup flows.
3. Extend app CLI args and stats command handler for background start/stop lifecycle responses, including already-running and stale-state handling.
4. Add a dedicated stats-daemon runtime/state-file path in the app and bypass the normal single-instance lock only for that mode.
5. Verify with focused tests first, then launcher/env lane, and update task acceptance criteria/final summary before handoff.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User approved option 2 design: dedicated app-side stats daemon, `subminer stats -b` to start, `subminer stats -s` to stop server only.

Implemented launcher `stats -b` and `stats -s` flows plus app-side `--stats-background` / `--stats-stop` handling.

Added background stats daemon state-file management and remote-daemon reuse so normal SubMiner instances do not try to bind a second stats server when the daemon is already running.

Verification: `bun test launcher/main.test.ts launcher/commands/command-modules.test.ts launcher/parse-args.test.ts src/main/runtime/stats-cli-command.test.ts src/main/early-single-instance.test.ts`, `bun run typecheck`, `bun run test:env`, `bun run test:fast`, `bun run build`, `bun run test:smoke:dist`, `bun run docs:test`, `bun run docs:build`, `bun run changelog:lint`.

Non-blocking note: `bun run test:launcher` still showed unrelated existing failures in `launcher/picker.test.ts` and an intermittent `launcher/smoke.e2e.test.ts` mpv-status check on this machine; the narrowed launcher suites covering the changed stats paths passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a dedicated background stats-daemon lifecycle for the launcher and app runtime. `subminer stats -b` now starts or reuses a detached stats server and returns after startup acknowledgement, while `subminer stats -s` stops that daemon without touching browser tabs. On the app side, new stats background/stop CLI flags bypass the normal single-instance lock only for daemon helper processes, write/read a daemon state file under user data, and reuse an already-running daemon instead of attempting a second local stats bind when another SubMiner instance needs stats access. Updated docs-site stats docs, added a changelog fragment, and covered the new flows with launcher parse/dispatch tests, app stats CLI handler tests, and single-instance bypass tests. Verification run: `bun run typecheck`, `bun run test:env`, `bun run test:fast`, `bun run build`, `bun run test:smoke:dist`, `bun run docs:test`, `bun run docs:build`, `bun run changelog:lint`, plus narrowed changed-path launcher/app test bundles.
<!-- SECTION:FINAL_SUMMARY:END -->
