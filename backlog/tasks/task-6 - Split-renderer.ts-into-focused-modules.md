---
id: TASK-6
title: Split renderer.ts into focused modules
status: Done
assignee:
  - codex
created_date: '2026-02-11 08:20'
updated_date: '2026-02-18 04:11'
labels:
  - refactor
  - renderer
  - architecture
milestone: Codebase Clarity & Composability
dependencies:
  - TASK-5
references:
  - src/renderer/renderer.ts
priority: high
ordinal: 52000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
renderer.ts is 2,754 lines with 94 functions handling 6+ distinct concerns: subtitle rendering, invisible overlay positioning, 4 modal UIs (Jimaku, Kiku, RuntimeOptions, Subsync), event handlers, keyboard chord system, and platform-specific layout. 16+ module-level state variables track overlapping modal states.

Proposed structure:
```
src/renderer/
├── renderer.ts          (entry point, initialization, IPC listeners)
├── state.ts             (centralized state object replacing 16+ scattered lets)
├── subtitle-render.ts   (renderSubtitle, renderTokenized, renderCharLevel, renderPlain)
├── positioning.ts       (applyInvisibleSubtitleLayoutFromMpvMetrics + helpers)
├── modals/
│   ├── jimaku.ts        (Jimaku download modal - lines 1097-1518)
│   ├── kiku.ts          (Kiku field grouping modal - lines 1519-1702)
│   ├── runtime-options.ts (Runtime options modal - lines 1247-1364)
│   └── subsync.ts       (Subsync modal - lines 1387-1466)
├── handlers/
│   ├── keyboard.ts      (keydown handlers, chord system)
│   └── mouse.ts         (drag, hover, click handlers)
└── utils/
    ├── dom.ts           (DOM element access with validation)
    └── platform.ts      (isLinux/isMacOS detection, platform-specific helpers)
```

Note: The renderer runs in Electron's renderer process, so module bundling considerations (esbuild/webpack or Electron's native ESM) need to be evaluated.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 renderer.ts reduced to <400 lines (init + IPC wiring)
- [x] #2 Each modal UI in its own module
- [x] #3 Positioning logic extracted with helper functions replacing the 211-line mega function
- [x] #4 State centralized in a single object/module
- [x] #5 Platform-specific logic isolated behind abstractions
- [x] #6 All existing functionality preserved
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Create shared renderer infrastructure modules (`state.ts`, `platform.ts`, `dom.ts`) and a typed context for cross-module dependencies.
2. Extract subtitle render and secondary subtitle logic into `subtitle-render.ts` with behavior-preserving APIs.
3. Extract invisible/visible subtitle positioning and offset edit logic into `positioning.ts`, splitting the mega layout function into helper functions.
4. Extract each modal into separate modules: `modals/jimaku.ts`, `modals/kiku.ts`, `modals/runtime-options.ts`, `modals/subsync.ts`.
5. Extract input and UI interaction logic into `handlers/keyboard.ts` and `handlers/mouse.ts`.
6. Rewrite `renderer.ts` as entrypoint/orchestrator only (<400 lines), wire IPC listeners and module composition.
7. Run `pnpm run build` and targeted tests; update task notes and acceptance checklist to reflect completion status.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reviewed src/renderer/renderer.ts structure and build wiring (tsc CommonJS output loaded by Electron). Confirmed renderer module splitting can be done safely without introducing a new bundler in this task.

Implemented renderer modularization with a centralized `RendererState` and shared context (`src/renderer/state.ts`, `src/renderer/context.ts`).

Extracted platform and DOM abstractions into `src/renderer/utils/platform.ts` and `src/renderer/utils/dom.ts`.

Extracted subtitle render/style concerns to `src/renderer/subtitle-render.ts` and positioning/layout concerns to `src/renderer/positioning.ts`, including helperized invisible subtitle layout pipeline.

Split modal UIs into dedicated modules: `src/renderer/modals/jimaku.ts`, `src/renderer/modals/kiku.ts`, `src/renderer/modals/runtime-options.ts`, `src/renderer/modals/subsync.ts`.

Split interaction logic into `src/renderer/handlers/keyboard.ts` and `src/renderer/handlers/mouse.ts`.

Reduced `src/renderer/renderer.ts` to entrypoint/orchestration (225 lines) with IPC wiring and module composition only.

Validation: `pnpm run build` passed; `pnpm run test:core` passed (21/21).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refactored the Electron renderer implementation from a monolithic file into focused modules while preserving runtime behavior and IPC integration.

What changed:
- Replaced ad hoc renderer globals with a centralized mutable state container in `src/renderer/state.ts`, wired through a shared renderer context (`src/renderer/context.ts`).
- Isolated platform/environment detection and DOM element resolution into `src/renderer/utils/platform.ts` and `src/renderer/utils/dom.ts`.
- Extracted subtitle rendering and subtitle style/secondary subtitle behavior into `src/renderer/subtitle-render.ts`.
- Extracted subtitle positioning logic into `src/renderer/positioning.ts`, including breaking invisible subtitle layout into helper functions for scale, container layout, vertical alignment, and typography application.
- Split each modal into its own module:
  - `src/renderer/modals/jimaku.ts`
  - `src/renderer/modals/kiku.ts`
  - `src/renderer/modals/runtime-options.ts`
  - `src/renderer/modals/subsync.ts`
- Split user interaction concerns into handler modules:
  - `src/renderer/handlers/keyboard.ts`
  - `src/renderer/handlers/mouse.ts`
- Rewrote `src/renderer/renderer.ts` to an initialization/orchestration entrypoint (225 lines), retaining IPC listeners and module composition only.

Why:
- Addressed architectural and maintainability issues in a large mixed-concern renderer file by enforcing concern boundaries and explicit dependencies.
- Improved testability and future change safety by reducing hidden cross-function/module state coupling.

Validation:
- `pnpm run build` succeeded.
- `pnpm run test:core` succeeded (21 passing tests).
<!-- SECTION:FINAL_SUMMARY:END -->
