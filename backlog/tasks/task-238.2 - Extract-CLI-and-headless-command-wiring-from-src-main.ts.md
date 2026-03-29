---
id: TASK-238.2
title: Extract CLI and headless command wiring from src/main.ts
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - tech-debt
  - cli
  - runtime
  - maintainability
milestone: m-0
dependencies:
  - TASK-238.1
references:
  - src/main.ts
  - src/main/cli-runtime.ts
  - src/cli/args.ts
  - launcher
parent_task_id: TASK-238
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` still owns the headless-initial-command flow, argument handling, and a large amount of CLI/runtime bridging. That makes non-window startup paths difficult to reason about and keeps CLI behavior coupled to unrelated desktop boot logic. Extract the remaining CLI/headless orchestration into dedicated runtime services so the main entrypoint only decides which startup path to invoke.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 CLI parsing, initial-command dispatch, and headless command execution no longer live as large inline flows in `src/main.ts`.
- [ ] #2 The new modules make the desktop startup path and headless startup path visibly separate and easier to test.
- [ ] #3 Existing CLI behaviors remain unchanged, including help output and startup gating behavior.
- [ ] #4 Targeted CLI/runtime tests cover the extracted path, and `bun run typecheck` passes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Map the current `parseArgs` / `handleInitialArgs` / `runHeadlessInitialCommand` / `handleCliCommand` flow in `src/main.ts`.
2. Extract a small startup-path selector plus dedicated runtime services for headless execution and interactive startup dispatch.
3. Keep Electron app ownership in `src/main.ts`; move only CLI orchestration and context assembly.
4. Verify with CLI-focused tests plus `bun run typecheck`.
<!-- SECTION:PLAN:END -->
