---
id: TASK-27.5
title: Split renderer positioning.ts into focused modules
status: To Do
assignee:
  - frontend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-13 21:17'
labels:
  - refactor
  - renderer
dependencies:
  - TASK-27.1
references:
  - src/renderer/renderer.ts
  - src/renderer/subtitle-render.ts
  - src/renderer/positioning.ts
  - src/renderer/handlers/keyboard.ts
  - src/renderer/handlers/mouse.ts
documentation:
  - docs/architecture.md
parent_task_id: TASK-27
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Split positioning.ts (513 LOC) — the only oversized file in the renderer — into focused modules. The rest of the renderer structure is already well-organized and does not need reorganization.

## Current Renderer State (already good)
- `renderer.ts` (241 lines) — pure composition, well-factored
- `state.ts` (132 lines), `context.ts` (14 lines), `subtitle-render.ts` (206 lines) — reasonable sizes
- `handlers/keyboard.ts` (238 lines), `handlers/mouse.ts` (271 lines) — focused
- `modals/` — 4 modal implementations, each self-contained
- Communication already uses explicit `ctx` pattern and function parameters, not globals

## What Actually Needs Work
`positioning.ts` mixes visible overlay positioning, invisible overlay positioning, MPV subtitle render metrics layout, and position persistence. These are distinct concerns that should be separate modules.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Split positioning.ts into at least 2 focused modules (e.g., visible-positioning and invisible-positioning, or by concern: layout, persistence, metrics).
- [ ] #2 No module exceeds 300 LOC.
- [ ] #3 Existing overlay behavior (subtitle positioning, drag, invisible layer metrics) unchanged.
- [ ] #4 renderer.ts imports stay clean — use an index re-export if needed.
- [ ] #5 Manual validation: subtitle positioning, drag/select, invisible layer alignment all work correctly.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Downscope Rationale

Original task proposed creating src/renderer/subtitles/, src/renderer/input/, src/renderer/state/ directories and introducing "explicit interfaces/events for keyboard/mouse/positioning/state updates to avoid global mutable coupling."

Review found:
1. The renderer already uses a `ctx` composition pattern — no global mutable coupling exists
2. Files are already organized by concern (handlers/, modals/, utils/)
3. Only positioning.ts (513 LOC) exceeds the 400 LOC threshold
4. Creating new directory structures for files under 300 lines adds churn without proportional benefit

Reduced scope to: split positioning.ts only. If future feature work (JLPT tagging, frequency highlighting) adds significant renderer complexity, a broader reorganization can be reconsidered then.
<!-- SECTION:NOTES:END -->
