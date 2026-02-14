---
id: TASK-39
title: Add hot-reload for non-destructive config changes
status: To Do
assignee: []
created_date: '2026-02-14 02:04'
labels:
  - config
  - developer-experience
  - quality-of-life
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Watch the config file for changes and apply non-destructive updates (colors, font sizes, subtitle modes, overlay opacity, keybindings) without requiring an app restart.

## Motivation
Currently all config is loaded at startup. Users tweaking visual settings (font size, colors, subtitle positioning) must restart the app after every change, which breaks their video session. Hot-reload for safe config values would dramatically improve the tuning experience.

## Scope
1. Watch the config file using `fs.watch` or similar
2. On change, re-parse and re-validate the config
3. Categorize config fields as hot-reloadable vs restart-required
4. Apply hot-reloadable changes immediately (push to renderer via IPC if needed)
5. For restart-required changes, show a notification that a restart is needed
6. Debounce file-change events (editors save multiple times rapidly)

## Hot-reloadable candidates
- Font family, size, weight, color
- Subtitle background opacity/color
- Secondary subtitle display mode
- Overlay opacity and positioning offsets
- Keybinding mappings
- AI translation provider settings

## Restart-required (NOT hot-reloadable)
- Anki field mappings (affects card creation pipeline)
- MeCab path / tokenizer settings
- MPV socket path
- Window tracker selection
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Config file changes are detected automatically via file watcher.
- [ ] #2 Hot-reloadable fields are applied immediately without restart.
- [ ] #3 Restart-required fields trigger a user-visible notification.
- [ ] #4 File change events are debounced to handle editor save patterns.
- [ ] #5 Invalid config changes are rejected with an error notification, keeping the previous valid config.
- [ ] #6 Renderer receives updated styles/settings via IPC without full page reload.
<!-- AC:END -->
