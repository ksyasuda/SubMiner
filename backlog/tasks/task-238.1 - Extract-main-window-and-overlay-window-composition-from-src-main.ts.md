---
id: TASK-238.1
title: Extract main-window and overlay-window composition from src/main.ts
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - tech-debt
  - runtime
  - windows
  - maintainability
milestone: m-0
dependencies: []
references:
  - src/main.ts
  - src/main/runtime/composers
  - src/main/runtime/overlay-runtime-bootstrap.ts
  - docs/architecture/README.md
parent_task_id: TASK-238
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` still directly owns several `BrowserWindow` construction and window-lifecycle paths, including overlay-adjacent windows and setup flows. That keeps the composition root far larger than intended and makes window behavior hard to test in isolation. Extract the remaining window/bootstrap composition into named runtime modules so `src/main.ts` mostly wires dependencies and app lifecycle events together.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 At least the main overlay window path plus two other window/setup flows are extracted from direct `BrowserWindow` construction inside `src/main.ts`.
- [ ] #2 The extracted modules expose narrow factory/handler APIs that can be tested without booting the whole app.
- [ ] #3 `src/main.ts` becomes materially smaller and easier to scan, with window creation concentrated behind well-named runtime surfaces.
- [ ] #4 Relevant runtime/window tests pass, and new tests are added for any newly isolated window composition helpers.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Map the remaining direct `BrowserWindow` creation sites in `src/main.ts` and group them by shared lifecycle concerns.
2. Extract coherent modules for construction, preload/path resolution, and open/focus/reuse behavior rather than moving raw option objects wholesale.
3. Update the composition root to consume the new modules and keep side effects/app state ownership explicit.
4. Verify with focused runtime/window tests plus `bun run typecheck`.
<!-- SECTION:PLAN:END -->
