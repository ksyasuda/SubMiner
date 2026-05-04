---
id: TASK-339
title: Stop pinning Hyprland overlay windows
status: Done
assignee: []
created_date: '2026-05-04 06:07'
updated_date: '2026-05-04 06:09'
labels:
  - linux
  - hyprland
  - overlay
  - bug
dependencies: []
references:
  - src/core/services/hyprland-window-placement.ts
  - src/core/services/overlay-window.ts
  - src/core/services/stats-window.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Recent Hyprland placement fix pins SubMiner overlay/stats windows, making them follow across workspaces instead of staying attached to mpv. Keep the float-for-bounds behavior, but never pin overlay windows.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Hyprland placement dispatches set floating state only and does not dispatch pin.
- [x] #2 Regression coverage proves pinned clients are unpinned or at least not re-pinned by SubMiner.
- [x] #3 Targeted tests and typecheck pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Changed Hyprland placement dispatch construction so unpinned overlay windows only get `setfloating`; pinned overlay windows get a single `pin` dispatch to toggle the bad prior pinned state off. This preserves floating placement for bounds while keeping overlay windows workspace-local with mpv.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Stopped re-pinning Hyprland overlay/stats windows during placement.
- Added cleanup behavior for previously pinned SubMiner windows by toggling pin only when Hyprland reports `pinned: true`.
- Updated regression coverage and added a changelog fragment.

Verification:
- `bun test src/core/services/hyprland-window-placement.test.ts src/core/services/overlay-window-config.test.ts src/core/services/stats-window.test.ts src/core/services/overlay-window-bounds.test.ts src/window-trackers/hyprland-tracker.test.ts`
- `bun run typecheck`
- `bun run changelog:lint`
- `bun run test:fast`
<!-- SECTION:FINAL_SUMMARY:END -->
