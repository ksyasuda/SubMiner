---
id: TASK-15
title: Fix renderer module loading regression after task 6 split
status: Done
assignee: []
created_date: '2026-02-12 00:45'
updated_date: '2026-02-18 04:11'
labels:
  - regression
  - overlay
  - renderer
dependencies: []
priority: high
ordinal: 49000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Overlay renderer stopped initializing after renderer.ts was split into modules. The emitted JS now uses CommonJS require/exports in a browser context (nodeIntegration disabled), causing script load failure and a blank transparent overlay with missing subtitle interactions.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Renderer script loads successfully in overlay BrowserWindow without nodeIntegration.
- [x] #2 Visible overlay displays subtitles again on initial launch.
- [x] #3 Overlay keyboard/mouse interactions are functional again.
- [x] #4 Build output remains compatible with Electron main/preload while renderer runs as browser modules.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Fixed a renderer module-loading regression introduced by renderer modularization. Added a dedicated renderer TypeScript build target (`tsconfig.renderer.json`) that emits browser-compatible ES modules, updated build script to compile renderer with that config, switched overlay HTML to load `renderer.js` as a module, and updated renderer runtime imports to `.js` module specifiers. Verified that built renderer output no longer contains CommonJS `require(...)` and that core test suite passes (`pnpm run test:core`).

<!-- SECTION:FINAL_SUMMARY:END -->
