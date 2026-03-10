---
id: TASK-78
title: 'Launcher + mpv plugin: auto-start visible overlay pause-until-ready and single-start guard'
status: Done
assignee: []
created_date: '2026-02-28 22:45'
updated_date: '2026-02-28 22:45'
labels: []
dependencies: []
priority: medium
ordinal: 9000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add startup gating behavior for wrapper + mpv plugin flow so playback starts paused when visible overlay auto-start is enabled, then auto-resumes only after subtitle tokenization is ready.

Scope:

- Plugin option `auto_start_pause_until_ready` (default `yes`).
- Launcher reads plugin runtime config and starts mpv paused when `auto_start=yes`, `auto_start_visible_overlay=yes`, and `auto_start_pause_until_ready=yes`.
- Main process signals readiness via mpv script message after tokenized subtitle delivery.
- Prevent duplicate auto-start attempts from showing `SubMiner already running` OSD.
- Keep startup/loading OSD messaging visible and update docs/tests.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Launcher reads `auto_start`, `auto_start_visible_overlay`, and `auto_start_pause_until_ready` from `subminer.conf` and starts mpv with `--pause=yes` when all are enabled.
- [x] #2 Plugin pauses on eligible auto-start and resumes only on readiness signal or timeout fallback.
- [x] #3 Main process emits `script-message subminer-autoplay-ready` after subtitle tokenization is ready.
- [x] #4 Auto-start duplicate triggers are idempotent (no duplicate `--start` behavior and no spurious `Already running` OSD for auto-start path).
- [x] #5 Docs and regression tests cover defaults, startup gating behavior, and duplicate-start suppression.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented startup pause gate across launcher/plugin/main runtime:

- Added plugin runtime config parsing in launcher (`auto_start`, `auto_start_visible_overlay`, `auto_start_pause_until_ready`) and mpv start-paused behavior for eligible runs.
- Added plugin auto-play gate state, timeout fallback, and readiness release via `subminer-autoplay-ready` script message.
- Added main-process readiness signaling after tokenization delivery, including unpause fallback command path.
- Split auto-start visibility control into separate control commands and added duplicate auto-start idempotency guard to suppress repeated auto-start `Already running` noise.
- Updated plugin defaults to enabled (`auto_start=yes`, `auto_start_visible_overlay=yes`, `auto_start_pause_until_ready=yes`) and refreshed docs (`README`, usage, launcher, installation, plugin/config docs).
- Added/updated regression coverage (`scripts/test-plugin-start-gate.lua`, launcher smoke/unit tests) validating paused startup, readiness resume, and duplicate-start suppression.

<!-- SECTION:FINAL_SUMMARY:END -->
