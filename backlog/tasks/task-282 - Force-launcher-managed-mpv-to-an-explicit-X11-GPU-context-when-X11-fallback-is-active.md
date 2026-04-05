---
id: TASK-282
title: >-
  Force launcher-managed mpv to an explicit X11 GPU context when X11 fallback is
  active
status: Done
assignee:
  - codex
created_date: '2026-04-05 21:14'
updated_date: '2026-04-05 21:15'
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
Follow-up to the launcher X11 fallback work: on Plasma Wayland, stripping Wayland env vars alone is not sufficient for launcher-managed mpv. mpv can still fail GPU initialization under the default video output path (`vo/gpu-next`) unless an explicit X11 GPU context is selected. Update launcher-managed mpv startup so X11 fallback mode also appends explicit mpv options for an X11 GPU context, while preserving supported Hyprland/Sway Wayland flows.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher-managed mpv appends explicit X11 GPU context args when explicit `--backend=x11` is used on Linux.
- [x] #2 Launcher-managed mpv appends the same explicit X11 GPU context args for Linux auto-mode fallback on unsupported Wayland desktops.
- [x] #3 Supported Hyprland/Sway Wayland flows do not receive the forced X11 GPU context args.
- [x] #4 Regression tests cover the forced-X11 mpv arg selection.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused launcher tests for explicit X11 GPU-context arg selection, unsupported Wayland auto fallback, and Hyprland/Sway no-regression cases.
2. Introduce a shared launcher helper that decides when mpv should be forced onto X11 fallback mode.
3. Use that helper to append explicit mpv X11 GPU-context args in both normal and detached idle mpv launch paths.
4. Run focused launcher tests plus TypeScript verification, then record remaining runtime follow-up guidance.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User runtime log showed `vo/gpu-next` failing with `Failed initializing any suitable GPU context!` under forced-X11 playback, which indicates env forcing alone was insufficient. Selected `--gpu-context=x11egl,x11` as the explicit mpv fallback: prefer X11/EGL, with GLX as a compatibility fallback.

Verification: `bun test launcher/mpv.test.ts --test-name-pattern "buildMpv(Env|BackendArgs)"` passed. `bun run tsc --noEmit` passed.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added explicit mpv backend args for launcher-managed X11 fallback mode. The launcher now uses a shared `shouldForceX11MpvBackend` decision for both env rewriting and mpv arg selection, so explicit `--backend=x11` and unsupported Linux Wayland auto fallback both append `--gpu-context=x11egl,x11` while still stripping Wayland env hints. This preserves supported Hyprland/Sway native Wayland flows and makes the X11 fallback more explicit for mpv's GPU initialization path.

Wired the new X11 GPU-context args into both the normal playback launch path and the detached idle mpv launch path. Added focused regression coverage for explicit `x11`, Plasma-style unsupported Wayland auto fallback, and Hyprland/Sway no-regression behavior. Verification completed with `bun test launcher/mpv.test.ts --test-name-pattern "buildMpv(Env|BackendArgs)"` and `bun run tsc --noEmit`.
<!-- SECTION:FINAL_SUMMARY:END -->
