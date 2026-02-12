---
id: TASK-13
title: Fix macOS native window bounds for overlay binding
status: Done
assignee:
  - codex
created_date: '2026-02-11 15:45'
updated_date: '2026-02-11 16:20'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - src/window-trackers/macos-tracker.ts
  - scripts/get-mpv-window-macos.swift
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Overlay windows on macOS are not properly aligned to the mpv window after switching from AppleScript window discovery to native Swift/CoreGraphics bounds retrieval.

Implement a robust native bounds strategy that prefers Accessibility window geometry (matching app-window coordinates used previously) and falls back to filtered CoreGraphics windows when Accessibility data is unavailable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Overlay bounds track the active mpv window with correct position and size on macOS.
- [x] #2 Helper avoids selecting off-screen/non-primary mpv-related windows.
- [x] #3 Build succeeds with the updated macOS helper.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Follow-up in progress after packaged app runtime showed fullscreen fallback behavior:
- Added packaged-app helper path resolution in tracker (`process.resourcesPath/scripts/get-mpv-window-macos`).
- Added `.asar` helper materialization to temp path so child process execution is possible if candidate path resolves inside asar.
- Added throttled tracker logging for helper execution failures to expose runtime errors without log spam.
- Updated Electron builder `extraResources` to ship `dist/scripts/get-mpv-window-macos` outside asar at `resources/scripts/get-mpv-window-macos`.
- Added macOS-only invisible subtitle vertical nudge (`+4px`) in renderer layout to align interactive subtitles with mpv glyph baseline after bounds fix.
<!-- SECTION:NOTES:END -->
