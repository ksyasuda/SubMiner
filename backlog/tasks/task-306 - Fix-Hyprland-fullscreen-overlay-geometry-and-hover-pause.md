---
id: TASK-306
title: Fix Hyprland fullscreen overlay geometry and hover pause
status: Done
assignee: []
created_date: '2026-04-27 01:44'
labels:
  - linux
  - hyprland
  - overlay
  - bug
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Overlay should track mpv geometry through Hyprland fullscreen transitions, stay above fullscreen video, and keep primary subtitle hover pause working after fullscreen/toggle cycles.

Implemented by observing mpv fullscreen property changes in addition to Hyprland geometry events, then refreshing visible overlay bounds/layering on Linux.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Hyprland tracker reacts to fullscreen/window state changes with updated geometry.
- [x] #2 Visible overlay is re-layered above mpv after Hyprland fullscreen geometry updates.
- [x] #3 Primary subtitle hover pause remains active after overlay geometry changes or visible overlay toggle cycles.
<!-- AC:END -->
