---
id: TASK-273
title: >-
  Fix first-run setup false positive when canonical mpv plugin is already
  installed
status: Done
assignee:
  - Kyle Yasuda
created_date: '2026-04-03 23:26'
updated_date: '2026-04-04 00:31'
labels:
  - bug
  - macos
  - first-run-setup
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and fix launcher/app first-run setup gating so playback does not block when the SubMiner mpv plugin is already installed at the canonical mpv config path on macOS. Align mpv path resolution with the actual install location, keep plugin detection scoped to the canonical plugin entrypoint, and make launcher setup gating resilient to stale cancelled setup state.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `resolveDefaultMpvInstallPaths` resolves the canonical macOS mpv config path used by existing installs.
- [ ] #2 Playback launcher bypasses first-run setup when the canonical `scripts/subminer/main.lua` plugin entrypoint already exists, even if `setup-state.json` is stale.
- [ ] #3 Regression tests cover canonical plugin detection and launcher handling of stale cancelled setup state.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Root cause ended up split across path resolution and launcher gating. No automated test command was executed in this pass by request.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated macOS mpv install path resolution to use the canonical `~/.config/mpv` location so first-run plugin detection matches the actual installed plugin path.

Restricted plugin detection to the canonical `scripts/subminer/main.lua` entrypoint instead of config presence or legacy loader files.

Updated the launcher setup gate to bypass stale `setup-state.json` when the mpv plugin is already installed, and to ignore an initially stale `cancelled` state after spawning setup.

Added regression coverage for canonical macOS detection and launcher setup-gate bypass behavior. No automated test command was executed in this pass by request.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Manual verification with scenario: existing plugin installed in custom mpv config path does not open first-run setup.
<!-- DOD:END -->
