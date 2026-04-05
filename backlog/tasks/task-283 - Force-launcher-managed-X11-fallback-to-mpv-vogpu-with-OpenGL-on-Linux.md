---
id: TASK-283
title: Force launcher-managed X11 fallback to mpv vo=gpu with OpenGL on Linux
status: Done
assignee:
  - codex
created_date: '2026-04-05 21:19'
updated_date: '2026-04-05 21:20'
labels:
  - bug
  - linux
  - launcher
  - mpv
  - overlay
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Follow-up to explicit X11 gpu-context forcing: Plasma Wayland runtime logs still show mpv using `vo/gpu-next` and failing to initialize any suitable GPU context under launcher-managed X11 fallback. Update launcher-managed mpv X11 fallback mode to force a more compatible renderer stack: `--vo=gpu`, `--gpu-api=opengl`, and `--gpu-context=x11egl,x11`, while preserving supported native Wayland flows and allowing explicit user mpv args to override later on the command line.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher-managed mpv appends `--vo=gpu`, `--gpu-api=opengl`, and `--gpu-context=x11egl,x11` when explicit `--backend=x11` is used on Linux.
- [x] #2 Launcher-managed mpv appends the same renderer args for Linux auto-mode fallback on unsupported Wayland desktops.
- [x] #3 Supported Hyprland/Sway Wayland flows do not receive the forced X11 renderer args.
- [x] #4 Regression tests cover the forced-X11 renderer arg selection.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Tighten launcher tests so forced-X11 renderer args require `--vo=gpu`, `--gpu-api=opengl`, and `--gpu-context=x11egl,x11`.
2. Update the shared launcher X11-fallback helper to return the full renderer arg stack for explicit `x11` and unsupported Wayland auto fallback.
3. Re-run focused launcher env/backend tests and TypeScript verification.
4. Hand back with retry instructions and next debugging branch if runtime still fails.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Runtime log still showed `[vo/gpu-next] Failed initializing any suitable GPU context!`, which meant forcing only the context was not enough. Updated the fallback to force the classic OpenGL renderer path too: `--vo=gpu --gpu-api=opengl --gpu-context=x11egl,x11`.

Verification: `bun test launcher/mpv.test.ts --test-name-pattern "buildMpv(Env|BackendArgs)"` passed. `bun run tsc --noEmit` passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated launcher-managed X11 fallback mode to force a more compatible mpv renderer stack on Linux: `--vo=gpu`, `--gpu-api=opengl`, and `--gpu-context=x11egl,x11`. This applies both to explicit `--backend=x11` and to unsupported Wayland auto fallback, while supported Hyprland/Sway Wayland sessions still keep their native path. The renderer args are still inserted before user-supplied `--args`, so an explicit user override can win later on the command line if needed.

Adjusted regression coverage to require the full renderer stack for forced-X11 mode and verified the helper behavior with focused launcher tests plus TypeScript compilation. Verification completed with `bun test launcher/mpv.test.ts --test-name-pattern "buildMpv(Env|BackendArgs)"` and `bun run tsc --noEmit`.
<!-- SECTION:FINAL_SUMMARY:END -->
