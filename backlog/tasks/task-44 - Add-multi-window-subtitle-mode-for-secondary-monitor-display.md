---
id: TASK-44
title: Add multi-window subtitle mode for secondary monitor display
status: To Do
assignee: []
created_date: '2026-02-14 02:15'
labels:
  - feature
  - renderer
  - multi-monitor
  - quality-of-life
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Support rendering the subtitle overlay on a secondary display at a larger, more readable size while the video plays full-screen on the primary monitor.

## Motivation
For study-focused sessions, reading subtitles overlaid on video at normal size can cause eye strain. Users with multiple monitors would benefit from a dedicated subtitle display — larger font, clean background, with full tokenization and click-to-lookup functionality — while the video plays unobstructed on the main screen.

## Features
1. **Detached subtitle window**: A second Electron BrowserWindow that mirrors the current subtitle content
2. **Configurable placement**: User selects which monitor and position for the subtitle window
3. **Independent styling**: The detached window can have different font size, background opacity, and layout than the overlay
4. **Synchronized content**: Both windows show the same subtitle in real-time
5. **Lookup support**: Yomitan click-to-lookup works in the detached window
6. **Toggle**: Shortcut to quickly enable/disable the secondary display

## Technical considerations
- Electron supports multi-monitor via `screen.getAllDisplays()` and `BrowserWindow` bounds
- The detached window should receive subtitle updates via IPC from the main process (same path as the overlay)
- Overlay on the primary monitor could optionally be hidden when the detached window is active
- Window should remember its position across sessions (persist to config or state file)
- Consider whether the detached window should be always-on-top or normal
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 A detached subtitle window can be opened on a secondary monitor.
- [ ] #2 Subtitle content is synchronized in real-time between overlay and detached window.
- [ ] #3 Detached window has independently configurable font size and styling.
- [ ] #4 Yomitan word lookup works in the detached window.
- [ ] #5 Shortcut toggles the detached window on/off.
- [ ] #6 Window position is persisted across sessions.
<!-- AC:END -->
