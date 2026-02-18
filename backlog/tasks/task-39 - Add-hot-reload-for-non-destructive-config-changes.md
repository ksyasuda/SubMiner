---
id: TASK-39
title: Add hot-reload for non-destructive config changes
status: Done
assignee:
  - '@sudacode'
created_date: '2026-02-14 02:04'
updated_date: '2026-02-18 09:29'
labels:
  - config
  - developer-experience
  - quality-of-life
dependencies: []
references:
  - docs/plans/2026-02-18-task-39-hot-reload-config.md
  - README.md
  - docs/configuration.md
  - docs/usage.md
  - config.example.jsonc
  - docs/public/config.example.jsonc
priority: low
ordinal: 59000
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
- [x] #1 Config file changes are detected automatically via file watcher.
- [x] #2 Hot-reloadable fields are applied immediately without restart.
- [x] #3 Restart-required fields trigger a user-visible notification.
- [x] #4 File change events are debounced to handle editor save patterns.
- [x] #5 Invalid config changes are rejected with an error notification, keeping the previous valid config.
- [x] #6 Renderer receives updated styles/settings via IPC without full page reload.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Implementation plan documented in `docs/plans/2026-02-18-task-39-hot-reload-config.md`.

Execution breakdown:
1. Add strict config reload API in `ConfigService` so invalid JSON/JSONC updates are rejected without mutating current runtime config.
2. Add testable hot-reload coordinator service with debounced file-change handling and diff classification (hot-reload vs restart-required fields).
3. Wire coordinator into main process watcher lifecycle; apply hot changes (style/keybindings/shortcuts/secondary mode/anki AI), emit restart-needed notification for non-hot changes, and emit invalid-config notification.
4. Add renderer/preload live update IPC hooks for style/keybinding refresh without page reload.
5. Run build + targeted tests (`test:fast`) and validate criteria with notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented strict config reload path in `ConfigService` (`reloadConfigStrict`) to reject invalid JSON/JSONC while preserving the last valid in-memory config.

Added `createConfigHotReloadRuntime` with debounced watcher-driven reloads, hot-vs-restart diff classification, and invalid-config callback handling.

Wired runtime into main process startup/teardown with `fs.watch`-based config path monitoring (file or config dir fallback), hot update application, restart-needed notifications, and invalid-config notifications.

Added renderer IPC hot-update channel (`config:hot-reload`) through preload/types, updating keybindings map, subtitle style, and secondary subtitle mode without page reload.

Updated core/config tests and test runner list to cover strict reload behavior and new hot-reload runtime service.

Updated user-facing docs to describe hot-reload behavior and restart-required notifications in `README.md`, `docs/configuration.md`, and `docs/usage.md`.

Regenerated `config.example.jsonc` and docs-served `docs/public/config.example.jsonc` from updated template metadata so hot-reload notes appear inline in config sections.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented TASK-39 by introducing a debounced config hot-reload pipeline that watches config changes, safely reloads with strict JSON/JSONC validation, and applies non-destructive runtime updates without requiring a full app restart.

### What changed
- Added `reloadConfigStrict()` to `ConfigService` so malformed config edits are rejected and the previous valid runtime config remains active.
- Added new core runtime service `createConfigHotReloadRuntime` to:
  - debounce rapid save bursts,
  - classify changed fields into hot-reloadable vs restart-required,
  - route invalid reload errors to user-visible notifications.
- Wired hot-reload runtime into `main.ts` app lifecycle:
  - starts after initial config load,
  - stops on app quit,
  - applies hot updates for subtitle styling, keybindings, shortcuts, secondary subtitle mode defaults, and Anki AI config patching.
- Added renderer live-update IPC channel (`config:hot-reload`) via `preload.ts` + `types.ts`, and applied updates in renderer without page reload.
- Added/updated tests:
  - strict reload behavior in `src/config/config.test.ts`,
  - hot-reload diff/debounce/invalid handling in `src/core/services/config-hot-reload.test.ts`,
  - included new test file in `test:core:dist` script.

### Validation
- `bun run build && node --test dist/config/config.test.js`
- `bun run build && node --test dist/core/services/config-hot-reload.test.js`
- `bun run test:fast`

All above commands passed.

Added follow-up documentation coverage for the shipped feature: README feature/quickstart note, configuration hot-reload behavior section, usage guide live-reload section, and regenerated config templates with per-section hot-reload notes.
<!-- SECTION:FINAL_SUMMARY:END -->
