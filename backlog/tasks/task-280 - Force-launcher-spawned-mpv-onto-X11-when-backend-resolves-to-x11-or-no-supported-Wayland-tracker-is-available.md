---
id: TASK-280
title: >-
  Force launcher-spawned mpv onto X11 when backend resolves to x11 or no
  supported Wayland tracker is available
status: Done
assignee:
  - codex
created_date: '2026-04-05 21:01'
updated_date: '2026-04-05 21:05'
labels:
  - bug
  - linux
  - launcher
  - overlay
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
On Linux Plasma Wayland and similar sessions, `subminer --backend=x11` currently only changes SubMiner's window-tracker override. The launcher still spawns mpv without forcing an X11/XWayland backend, so the X11 tracker cannot find the mpv window and the overlay remains hidden. Update launcher-side mpv spawn behavior so launcher-managed mpv runs under X11 when backend resolves to `x11`, and also when auto detection cannot resolve to a supported Wayland tracker. Preserve existing Hyprland/Sway behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher-managed mpv is spawned with X11/XWayland-forcing environment/config when backend resolves to `x11`.
- [x] #2 Linux auto mode falls back to X11/XWayland-forced mpv when no supported Wayland tracker backend is detected.
- [x] #3 Hyprland and Sway launcher flows do not regress to forced X11 mpv.
- [x] #4 Regression tests cover launcher env/backend selection for these Linux cases.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused launcher tests that assert the mpv spawn environment forces X11 when backend resolves to `x11`, and when Linux auto mode cannot use a supported Wayland tracker.
2. Refactor launcher mpv spawn code to compute an mpv-specific environment without changing existing Hyprland/Sway flows.
3. Route all launcher-managed mpv spawns through the new environment helper.
4. Run focused launcher tests, then summarize behavior and any remaining verification gaps.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
User approved scope: force launcher-managed mpv to X11 for explicit `--backend=x11` and for unsupported Linux Wayland auto-detect fallback; preserve Hyprland/Sway behavior.

Implemented launcher-side `buildMpvEnv` to strip Wayland hints and force X11/XWayland for launcher-managed mpv when `--backend=x11`, and for Linux auto mode on unsupported Wayland desktops with an X11 display available. Wired both normal mpv launches and idle detached mpv launches through the helper.

Verification: `bun test launcher/mpv.test.ts --test-name-pattern "buildMpvEnv"` passed; `bun run tsc --noEmit` passed. A broader `bun test launcher/mpv.test.ts` run still hits a pre-existing sandbox-specific failure in `launchAppCommandDetached handles child process spawn errors` because this environment cannot write the default app log path under `/home/sudacode/.config/SubMiner/logs`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated the launcher so mpv gets an X11/XWayland-oriented spawn environment whenever the user explicitly requests `--backend=x11`, and when Linux auto mode is running under an unsupported Wayland desktop that still exposes an X11 display. The new helper reuses the launcher child-process base environment, strips Wayland-specific hints (`WAYLAND_DISPLAY`, Hyprland/Sway markers), and flips `XDG_SESSION_TYPE` to `x11` only for those fallback cases. Both foreground mpv launches and detached idle mpv launches now use the same helper so overlay-tracked playback stays consistent.

Added focused regression coverage in `launcher/mpv.test.ts` for three cases: explicit `x11` forcing, unsupported Wayland auto fallback (for example KDE Plasma Wayland), and preserving native Wayland env for supported Hyprland/Sway auto backends. Verification completed with `bun test launcher/mpv.test.ts --test-name-pattern "buildMpvEnv"` and `bun run tsc --noEmit`. A broader launcher mpv test run still shows an unrelated sandbox write failure for the default app log path in this environment.
<!-- SECTION:FINAL_SUMMARY:END -->
