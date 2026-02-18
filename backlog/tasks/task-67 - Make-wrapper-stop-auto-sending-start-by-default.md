---
id: TASK-67
title: Make wrapper stop auto-sending --start by default
status: Done
assignee: []
created_date: '2026-02-18 09:47'
updated_date: '2026-02-18 10:02'
labels:
  - launcher
  - wrapper
  - cli
  - background-mode
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update the subminer wrapper flow so it no longer appends --start automatically when launching the app, requiring explicit start semantics and allowing background AppImage workflow to be primary.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Wrapper launch path does not append --start unless explicitly requested
- [x] #2 Existing explicit startup commands still work and are documented
- [x] #3 Tests updated for new wrapper behavior
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Launcher wrapper no longer auto-starts overlay from mpv plugin `auto_start` setting; only explicit `--start`/`--start-overlay` trigger wrapper-managed startup.

Simplified plugin runtime config parsing in launcher to consume `socket_path` only for wrapper behavior.

Updated docs examples and descriptions to make explicit startup flow clear (`subminer --start video.mkv`), and rebuilt bundled `subminer` script.

Validated with `bun run build && make build-launcher && bun run test:fast`.

Follow-up: updated mpv plugin command builder (`plugin/subminer.lua`) to stop forcing `--start` for toggle/show/hide actions; only explicit `start` action now sends start context. This avoids second-instance command failures when app is already running in background mode.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated the `subminer` wrapper to stop implicitly issuing app `--start` based on plugin auto-start settings, so background AppImage usage is the default and overlay startup happens only on explicit wrapper flags (`--start`/`--start-overlay`) or manual plugin commands. Also updated launcher docs/examples to reflect explicit startup semantics and regenerated the bundled `subminer` script.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Validated with launcher-related tests or command simulation
<!-- DOD:END -->
