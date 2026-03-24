---
id: TASK-177.3
title: Fix attached stats command flow and browser config
status: Done
assignee:
  - '@codex'
created_date: '2026-03-19 20:15'
updated_date: '2026-03-23 03:22'
labels:
  - launcher
  - stats
  - cli
dependencies: []
references:
  - launcher/commands/stats-command.ts
  - launcher/commands/command-modules.test.ts
  - launcher/main.test.ts
  - src/main/runtime/stats-cli-command.ts
  - src/main/runtime/stats-cli-command.test.ts
parent_task_id: TASK-177
priority: medium
ordinal: 129500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Make `subminer stats` stay attached to the foreground app process instead of routing through daemon startup, while keeping background/stop behavior on the daemon path. Ensure browser opening for stats respects only `stats.autoOpenBrowser` in the normal stats flow.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Default `subminer stats` forwards through the attached foreground stats command path instead of the daemon-start path
- [x] #2 `subminer stats --background` and `subminer stats --stop` continue using the daemon control path
- [x] #3 Normal stats launches do not open a browser when `stats.autoOpenBrowser` is false, and automated tests cover the launcher/runtime regressions
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing launcher tests first so default `stats` expects `--stats` forwarding while `--background` and `--stop` continue to expect daemon control flags.
2. Add/adjust runtime stats command tests to prove `stats.autoOpenBrowser=false` suppresses browser opening on the normal attached stats path.
3. Patch launcher forwarding logic in `launcher/commands/stats-command.ts` to choose foreground vs daemon flags correctly without changing cleanup handling.
4. Run targeted launcher and stats runtime tests, then record verification results and blockers.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Confirmed root cause: launcher default `stats` flow always forwarded `--stats-daemon-start` plus `--stats-daemon-open-browser`, which detached the terminal process and bypassed `stats.autoOpenBrowser` because browser opening happened in daemon control instead of the normal stats CLI handler.

Updated launcher forwarding so plain `subminer stats` now uses the attached `--stats` path, while explicit `--background` and `--stop` continue using daemon control flags.

Added launcher regression coverage for the attached/default path and preserved background/stop expectations; added runtime coverage proving `stats.autoOpenBrowser=false` suppresses browser opening on the normal stats path.

Verifier passed for `launcher-plugin` and `runtime-compat` lanes. Artifact: .tmp/skill-verification/subminer-verify-20260319-131703-ZaAaUV.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed `subminer stats` so the default command now forwards to the normal attached `--stats` app path instead of the daemon-start path. That keeps the foreground process attached to the terminal as expected, while `subminer stats --background` and `subminer stats --stop` still use daemon control. Because the normal stats CLI path already respects `config.stats.autoOpenBrowser`, this also fixes the unwanted browser-open behavior that previously bypassed config via `--stats-daemon-open-browser`.

Added launcher command and launcher integration regressions for the new forwarding behavior, plus a runtime stats CLI regression that asserts `stats.autoOpenBrowser=false` suppresses browser opening. Verification passed with targeted launcher tests, targeted runtime stats tests, and the SubMiner verifier `launcher-plugin` + `runtime-compat` lanes.
<!-- SECTION:FINAL_SUMMARY:END -->
