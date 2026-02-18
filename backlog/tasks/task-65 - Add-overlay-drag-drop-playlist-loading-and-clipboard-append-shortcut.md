---
id: TASK-65
title: Add overlay drag-drop playlist loading and clipboard append shortcut
status: Done
assignee: []
created_date: '2026-02-18 13:10'
updated_date: '2026-02-18 13:10'
labels: []
dependencies: []
---

## Description

Implement direct playlist control from the overlay:

- Drag/drop video files onto overlay:
  - default drop: replace current playback with dropped set (first `replace`, remainder `append`)
  - `Shift` + drop: append all dropped files
- `Ctrl/Cmd+A`: read clipboard text, if it resolves to a supported local video file path, append it to mpv playlist.

## Implementation Steps

- [x] Add TDD coverage for drop path parsing, command mode generation, and clipboard path parsing (`src/core/services/overlay-drop.test.ts`).
- [x] Implement drop/clipboard parser + mpv command-builder utility (`src/core/services/overlay-drop.ts`).
- [x] Wire renderer drag/drop handling and mpv command dispatch (`src/renderer/renderer.ts`).
- [x] Add IPC API for clipboard append flow (`src/types.ts`, `src/preload.ts`, `src/core/services/ipc.ts`, `src/main/dependencies.ts`).
- [x] Implement main-process clipboard validation + append behavior (`src/main.ts`).
- [x] Add fixed keyboard shortcut hook (`Ctrl/Cmd+A`) in renderer keyboard handler (`src/renderer/handlers/keyboard.ts`, `src/renderer/renderer.ts`).
- [x] Update docs for new interaction model (`docs/usage.md`, `docs/configuration.md`).

## Verification

- `bun run build`
- `node --test dist/core/services/overlay-drop.test.js dist/core/services/ipc.test.js`
