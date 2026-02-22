---
id: TASK-81
title: Refactor launcher into command modules and process adapters
status: Done
assignee:
  - opencode-task81-launcher-modules
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 07:49'
labels:
  - launcher
  - cli
  - architecture
dependencies:
  - TASK-70
  - TASK-73
  - TASK-74
priority: medium
ordinal: 74000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launcher code is still large and process-control heavy (`launcher/main.ts`, `launcher/mpv.ts`, `launcher/config.ts`). This task refactors launcher flows into command modules plus explicit process adapters to reduce branching complexity and improve testability.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Split by command domain (`doctor`, `config`, `mpv`, `jellyfin`, `play`).
- Isolate shell/process invocations behind adapter interfaces.
- Keep argv parsing independent from command execution side effects.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define launcher command dispatcher and command handler interfaces.
2. Extract command handlers from `launcher/main.ts`.
3. Move spawn/fs/net operations into process + platform adapter modules.
4. Keep CLI UX and exit-code behavior stable.
5. Add unit tests per command module with mocked adapters.
6. Update launcher docs for module structure and extension guidance.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher commands are implemented as focused modules
- [x] #2 Process side effects are isolated behind adapter interfaces
- [x] #3 Existing CLI behavior and exit codes remain compatible
- [x] #4 Launcher testability improves via mocked adapter tests
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Add `launcher/process-adapter.ts` and `launcher/commands/context.ts` so command handlers receive a shared context and explicit process/stdout/signal/exit seam instead of hardcoding `process.*` calls.
2) Extract early-return command branches from `launcher/main.ts` into focused modules under `launcher/commands/` (`config`, `doctor`, `mpv`, `app`, `texthooker`) while preserving existing outputs and exit-code behavior.
3) Extract `jellyfin` and default playback orchestration into dedicated command modules so `launcher/main.ts` becomes thin command dispatch + lifecycle wiring.
4) Add adapter-mocked command tests (`launcher/commands/command-modules.test.ts`) and keep integration regressions in `launcher/main.test.ts` to prove behavior parity.
5) Wire new command test file into launcher/core source test scripts, run required gates (`bun run test:launcher`, `bun run test:core:src`), then finalize TASK-81 notes/AC/DoD without commit.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Planning complete. Detailed execution plan saved at docs/plans/2026-02-22-task-81-launcher-command-modules-process-adapters.md. Proceeding directly per user request to execute via writing-plans + executing-plans flow without commit.

Implemented launcher refactor by extracting command handlers into `launcher/commands/*` (`config`, `doctor`, `mpv`, `app`, `jellyfin`, `playback`) and converting `launcher/main.ts` into thin context+dispatch orchestration.

Added explicit process side-effect seam in `launcher/process-adapter.ts` and routed command output/exit/signal behavior through adapter-aware command modules.

Added mocked-adapter regression tests in `launcher/commands/command-modules.test.ts` and wired that file into `test:launcher:src` + `test:core:src` scripts.

Verification: `bun test launcher/commands/command-modules.test.ts launcher/main.test.ts launcher/parse-args.test.ts` and `bun run test:launcher && bun run test:core:src` all passing.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refactored launcher command branching into focused command modules and introduced a dedicated process adapter interface to isolate output/exit/signal side effects from orchestration code. Preserved existing CLI behavior by keeping helper flows intact, added adapter-mocked command tests, and validated compatibility with full launcher and core source test gates.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Launcher tests and core test gate pass
- [x] #2 No regression in wrapper command behavior
<!-- DOD:END -->
