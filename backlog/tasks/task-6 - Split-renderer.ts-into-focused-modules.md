---
id: TASK-6
title: Split renderer.ts into focused modules
status: To Do
assignee: []
created_date: '2026-02-11 08:20'
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
- [ ] #1 renderer.ts reduced to <400 lines (init + IPC wiring)
- [ ] #2 Each modal UI in its own module
- [ ] #3 Positioning logic extracted with helper functions replacing the 211-line mega function
- [ ] #4 State centralized in a single object/module
- [ ] #5 Platform-specific logic isolated behind abstractions
- [ ] #6 All existing functionality preserved
<!-- AC:END -->
