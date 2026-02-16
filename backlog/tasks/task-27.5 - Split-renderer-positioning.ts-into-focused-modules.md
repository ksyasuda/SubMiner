---
id: TASK-27.5
title: Split renderer positioning.ts into focused modules
status: Done
assignee:
  - frontend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-15 23:59'
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
- [x] #1 Split positioning.ts into at least 2 focused modules (e.g., visible-positioning and invisible-positioning, or by concern: layout, persistence, metrics).
- [x] #2 No module exceeds 300 LOC.
- [x] #3 Existing overlay behavior (subtitle positioning, drag, invisible layer metrics) unchanged.
- [x] #4 renderer.ts imports stay clean — use an index re-export if needed.
- [x] #5 Manual validation: subtitle positioning, drag/select, invisible layer alignment all work correctly.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
TASK-11 decomposition is implemented as part of this task by moving the prior monolithic mpv-metrics function into dedicated helper modules.

Refactored renderer positioning into focused modules via a new controller barrel plus helpers: position-state, invisible-offset, invisible-layout, invisible-layout-metrics, and invisible-layout-helpers.

Split `applyInvisibleSubtitleLayoutFromMpvMetrics` into math/layout/style helper functions and ensured no module exceeds 300 LOC by extracting two metric/style files.

Validation done in this run: TypeScript build passes (`npm run build`). Manual behavior verification pending.

`src/renderer/positioning.ts` now re-exports `createPositioningController` from `./positioning/controller.js`.

Acceptance updates: checked #1 (at least two focused modules) and #2 (no module >300 LOC).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed the renderer positioning split. `src/renderer/positioning.ts` is now a thin re-export and logic is decomposed into focused modules (`controller.ts`, `position-state.ts`, `invisible-offset.ts`, `invisible-layout.ts`, `invisible-layout-helpers.ts`, `invisible-layout-metrics.ts`). Kept `renderer.ts` call-sites unchanged and preserved APIs via controller return shape. Verified by `npm run build` and user manual validation.
<!-- SECTION:FINAL_SUMMARY:END -->
