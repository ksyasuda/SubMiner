---
id: TASK-166
title: Replace OSD notifications with overlay notifications and runtime fallback
status: In Progress
assignee: []
created_date: '2026-03-16 00:00'
updated_date: '2026-03-16 00:00'
labels:
  - overlay
  - notifications
  - ux
dependencies: []
references:
  - /home/sudacode/.t3/worktrees/SubMiner/t3code-c5d9c2f7/src/main.ts
  - /home/sudacode/.t3/worktrees/SubMiner/t3code-c5d9c2f7/src/renderer/renderer.ts
  - /home/sudacode/.t3/worktrees/SubMiner/t3code-c5d9c2f7/plugin/subminer/process.lua
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Replace SubMiner's mpv OSD-first notification flow with overlay-native notifications. The visible overlay becomes the primary in-app notification surface, with fallback to mpv OSD only when the overlay is unavailable, hidden, disabled, or failed to load. Electron/system notifications remain unchanged. Existing notification copy and loading states should be refreshed to fit the richer overlay surface, especially spinner-driven loading states.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Main-process logical OSD notifications render on the overlay when the visible overlay is available and visible.
- [ ] #2 Logical OSD notifications fall back to actual mpv OSD when the overlay is hidden, disabled, unavailable, or fails to load.
- [ ] #3 Notification-type config semantics remain intact: `osd` uses the overlay-or-OSD channel, `system` remains Electron notifications, `both` uses both channels, and `none` suppresses notifications.
- [ ] #4 Overlay notifications support richer presentation for loading/progress states, including a spinner/persistent loading state.
- [ ] #5 Regression coverage verifies routing, renderer behavior, and fallback handling.

<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Plan:
1. Add failing tests for a pure main-process overlay notification router and a renderer overlay notification controller.
2. Add overlay notification IPC/types plus renderer DOM/controller/CSS for styled toast notifications and spinner handling.
3. Route existing logical OSD callsites through the new overlay-or-OSD fallback seam in main.
4. Trim plugin-side direct OSD usage down to true fallback/error cases.
5. Verify with targeted runtime-compat/core lanes, then escalate if runtime behavior claims remain unproven.

<!-- SECTION:NOTES:END -->
