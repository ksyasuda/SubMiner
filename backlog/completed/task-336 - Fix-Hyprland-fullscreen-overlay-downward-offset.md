---
id: TASK-336
title: Fix Hyprland fullscreen overlay downward offset
status: Done
assignee: []
created_date: '2026-05-04 05:42'
updated_date: '2026-05-04 06:10'
labels:
  - linux
  - hyprland
  - overlay
  - bug
dependencies: []
references:
  - src/window-trackers/hyprland-tracker.ts
  - src/core/services/overlay-window-bounds.ts
  - src/main/runtime/linux-mpv-fullscreen-overlay-refresh.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
SubMiner visible overlay is slightly below mpv when mpv is fullscreen on Linux Hyprland. Align overlay bounds with mpv fullscreen client/monitor bounds.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Hyprland fullscreen mpv overlay uses top-aligned geometry instead of inheriting a downward offset.
- [x] #2 Regression coverage captures the fullscreen Hyprland geometry case.
- [x] #3 Targeted tests pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added follow-up Hyprland placement handling after the fullscreenClient geometry fix. SubMiner overlay/stats windows now get stable titles and, on Hyprland, are resolved from `hyprctl -j clients` by current PID/title, then set floating before bounds are applied. The stats overlay reapplies bounds after showing because Hyprland cannot see the hidden window before it is mapped.

2026-05-04 follow-up: offset remains after removing pinning. User reports stats modal still has a top gap from mpv in Hyprland fullscreen. Need inspect exact stats overlay CSS/window bounds after float-only placement.

2026-05-04 follow-up fix: stats CSS already had zero body margin, so the remaining gap points at native Hyprland placement after float-only handling. Added exact `movewindowpixel`/`resizewindowpixel` Hyprland dispatches using the same tracked mpv bounds passed to Electron.

2026-05-04 second follow-up: live `hyprctl -j clients` showed the SubMiner client was already full monitor size at `[0,0]`, so the remaining visible top strip was inside Electron's transparent stats surface rather than compositor geometry. Made the stats overlay BrowserWindow opaque with the stats base background. Also prevented page titles from overwriting the stable SubMiner overlay/stats titles used for Hyprland client matching.

2026-05-04 third follow-up: user confirmed native overlay placement is correct and the remaining gap is stats-page-specific. Made stats overlay mode paint an opaque full-viewport root/background and constrained the stats app to `h-screen` with an internal scrolling main pane, so the overlay page itself covers the mpv frame from y=0.

2026-05-04 fourth follow-up: live Hyprland data showed mpv and SubMiner shared the same outer geometry while stats content still rendered lower. Stats window placement now compensates for Electron/Wayland content insets using `getContentBounds()` versus `getBounds()`, then sends the adjusted outer bounds to Hyprland exact placement so the content area, not just the native surface, aligns to mpv.

2026-05-04 fifth follow-up: user confirmed the offset is Hyprland-fullscreen-only and not present while mpv is windowed. Added Hyprland `setprop` decoration cleanup during exact overlay placement (`rounding 0`, `border_size 0`, `no_shadow 1`, `no_blur 1`, `decorate 0`) because fullscreen mpv has square fullscreen edges while a floating SubMiner stats window can retain Hyprland floating-window decoration.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Treated Hyprland `fullscreenClient` as a fullscreen signal when resolving mpv overlay geometry.
- Added Hyprland window placement handling so SubMiner overlay/stats windows are set floating before bounds are applied.
- Added exact Hyprland move/resize dispatches so floating overlay/stats windows are force-aligned to the tracked mpv bounds.
- Gave overlay/stats windows stable titles for Hyprland client matching, and reapplied stats bounds after show.
- Locked overlay/stats window titles against page title changes and made the stats overlay window opaque so mpv cannot show through transparent Electron insets.
- Made the stats overlay page paint an opaque full-viewport background and added CSS regression coverage for overlay mode.
- Compensated stats overlay outer placement for Electron/Wayland content insets.
- Disabled Hyprland floating-window decoration for exact overlay placement over fullscreen mpv.
- Added regression coverage for the 28px fullscreen geometry shape and Hyprland placement dispatches.
- Added a changelog fragment for the overlay fix.

Verification:
- `bun test src/core/services/hyprland-window-placement.test.ts src/core/services/overlay-window-config.test.ts src/core/services/stats-window.test.ts src/core/services/overlay-window-bounds.test.ts src/window-trackers/hyprland-tracker.test.ts`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun run test:fast`
- `bun test stats/src/styles/globals.test.ts stats/src/lib/api-client.test.ts src/core/services/stats-window.test.ts`
- `bun run build:stats`
- `bun test src/core/services/stats-window.test.ts src/core/services/hyprland-window-placement.test.ts stats/src/styles/globals.test.ts`
- `bun test src/core/services/hyprland-window-placement.test.ts src/core/services/stats-window.test.ts stats/src/styles/globals.test.ts`
<!-- SECTION:FINAL_SUMMARY:END -->
