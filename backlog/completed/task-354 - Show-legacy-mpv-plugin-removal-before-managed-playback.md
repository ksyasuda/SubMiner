---
id: TASK-354
title: Show legacy mpv plugin removal before managed playback
status: Done
assignee: []
created_date: '2026-05-12 14:30'
updated_date: '2026-05-12 14:35'
labels:
  - launcher
  - mpv-plugin
  - setup
  - windows
dependencies: []
references:
  - app-managed-mpv-runtime-plugin-plan.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When SubMiner-managed playback detects a legacy global SubMiner mpv plugin, show the removal UI before mpv starts so users can optionally trash the legacy files and then launch with the bundled runtime plugin.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher playback opens first-run setup before mpv starts when legacy global plugin files are detected, even if setup is already completed.
- [x] #2 Launcher playback resumes with bundled runtime plugin after the legacy plugin is removed.
- [x] #3 Windows mpv shortcut/app launch prompts before spawning mpv and re-detects after removal so bundled injection is used.
- [x] #4 Users can continue without removal; removal remains optional.
- [x] #5 Regression coverage prevents bypassing the removal prompt due to completed setup or installed-plugin detection.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Managed launcher playback now checks for legacy global SubMiner mpv plugin files before choosing/loading a video, opens first-run setup even when setup is already complete, and waits for removal or explicit user continuation before starting mpv. Windows managed mpv launches now show a pre-launch removal dialog, move detected legacy files to the OS trash on confirmation, re-detect, and inject the bundled runtime plugin after successful removal.
<!-- SECTION:FINAL_SUMMARY:END -->
