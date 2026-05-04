---
id: TASK-336
title: Fix Hyprland fullscreen overlay downward offset
status: Done
assignee: []
created_date: '2026-05-04 05:42'
updated_date: '2026-05-04 05:56'
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
Added follow-up Hyprland placement handling after the fullscreenClient geometry fix. SubMiner overlay/stats windows now get stable titles and, on Hyprland, are resolved from `hyprctl -j clients` by current PID/title, then set floating and pinned before bounds are applied. The stats overlay reapplies bounds after showing because Hyprland cannot see the hidden window before it is mapped.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Treated Hyprland `fullscreenClient` as a fullscreen signal when resolving mpv overlay geometry.
- Added Hyprland window placement handling so SubMiner overlay/stats windows are set floating and pinned before bounds are applied.
- Gave overlay/stats windows stable titles for Hyprland client matching, and reapplied stats bounds after show.
- Added regression coverage for the 28px fullscreen geometry shape and Hyprland placement dispatches.
- Added a changelog fragment for the overlay fix.

Verification:
- `bun test src/core/services/hyprland-window-placement.test.ts src/core/services/overlay-window-config.test.ts src/core/services/stats-window.test.ts src/core/services/overlay-window-bounds.test.ts src/window-trackers/hyprland-tracker.test.ts`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun run test:fast`
<!-- SECTION:FINAL_SUMMARY:END -->
