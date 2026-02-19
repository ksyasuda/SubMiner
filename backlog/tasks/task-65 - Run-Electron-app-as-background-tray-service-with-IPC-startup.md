---
id: TASK-65
title: Run Electron app as background tray service with IPC startup
status: Done
assignee: []
created_date: '2026-02-18 08:48'
updated_date: '2026-02-19 21:50'
labels:
  - electron
  - tray
  - ipc
  - desktop-entry
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Allow launching the app from desktop/application entry so it starts in background with minimal logging, waits for IPC connection, and exposes tray icon for settings/configuration.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launching via app/desktop file starts app without foreground window by default
- [x] #2 Process runs in background with minimal startup logging
- [x] #3 Main process initializes IPC server/client wait loop for incoming connection
- [x] #4 Tray icon is visible and provides menu actions for settings/configuration and quit
- [x] #5 Behavior documented for desktop launch usage
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented `--background` CLI mode to keep app running without overlay windows, default to quieter logging, and preserve IPC client startup.

Added persistent tray runtime (icon + menu actions for overlay, Yomitan settings, runtime options, Jellyfin setup, AniList setup, quit).

Disabled window-all-closed auto-quit while in background mode; desktop Linux launcher now passes `--background` by default via electron-builder desktop Exec.

Updated usage/installation docs and CLI help; validated with `bun run build && bun run test:fast`.

Follow-up packaging fix: replaced invalid `build.linux.desktop.Exec` override with supported electron-builder `build.linux.executableArgs: ["--background"]` after AppImage schema validation failure on electron-builder 26.7.0.

Re-verified Linux packaging with `bun run build:appimage`; background launcher argument is now applied without config validation errors.

Correction: any earlier mention of `desktop Exec` is obsolete; final shipped config uses `build.linux.executableArgs` only.

Background launch now detaches from terminal via new `src/main-entry.ts` bootstrap: `--background` parent process spawns detached child and exits, so Ctrl+C no longer stops the running tray process.

Background detached child now suppresses Node runtime warnings (`NODE_NO_WARNINGS=1`) and strips `VK_INSTANCE_LAYERS` when it contains `lsfg` to reduce non-actionable startup noise in background mode.

Updated package entrypoint to `dist/main-entry.js` and docs usage note for detached background behavior.

macOS follow-up: tray icon handling now normalizes to status-bar-safe form (`18x18` resize + template image mode) to prevent oversized/non-interactive menu bar icons when running in `--background` mode.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added an always-on background startup mode for desktop launches by introducing `--background`, wiring startup/lifecycle state to keep the process alive without windows, and defaulting this mode to quieter logging unless explicitly overridden. Added a tray icon/menu for settings and runtime configuration entry points and updated Linux desktop packaging so launcher executions use background mode by default, with docs and tests updated accordingly.

Added detached background bootstrap behavior for `--background` launches so terminal invocations return immediately while the tray process continues independently, and reduced background startup noise by suppressing Node warnings and removing problematic lsfg Vulkan layer env in detached child startups.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Implementation verified locally with launch command or desktop entry simulation
<!-- DOD:END -->
