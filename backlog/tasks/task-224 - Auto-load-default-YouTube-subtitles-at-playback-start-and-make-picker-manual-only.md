---
id: TASK-224
title: >-
  Auto-load default YouTube subtitles at playback start and make picker
  manual-only
status: Done
assignee:
  - Codex
created_date: '2026-03-23 18:51'
updated_date: '2026-03-23 19:14'
labels:
  - youtube
  - mpv
  - overlay
  - keybindings
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/youtube-flow.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/youtube-track-picker.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/config/definitions/shared.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the mandatory YouTube subtitle picker startup flow with automatic default-track loading. On YouTube playback start, attempt to load the default primary subtitle and best-effort secondary subtitle without prompting. Gate playback only on primary subtitle load/tokenization readiness. If primary subtitle probing/download/loading fails, resume playback and report the failure through the configured notification/output path. Keep the YouTube subtitle picker as a regular overlay modal opened by a new default keybinding during active YouTube playback.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Opening a YouTube URL auto-selects and attempts to load the default primary subtitle without opening the picker modal.
- [x] #2 Opening a YouTube URL also attempts to load the default secondary subtitle when available, but playback never waits on secondary success.
- [x] #3 Playback remains gated only until the primary subtitle is loaded and tokenization is ready; primary failure resumes playback immediately.
- [x] #4 Primary auto-load failures report through the existing configured notification/output path and keep playback running.
- [x] #5 The YouTube subtitle picker can be opened manually during active YouTube playback via a new default keybinding.
- [x] #6 Regression tests cover startup auto-load success, primary failure fallback, and the manual picker keybinding flow.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for YouTube startup auto-load success, primary failure fallback, and manual picker keybinding flow.
2. Refactor the YouTube runtime to auto-select default tracks on startup, gate playback only on primary subtitle/tokenization readiness, and route failures through the configured notification/output path.
3. Add a new default keybinding and command path to open the YouTube picker manually during active YouTube playback.
4. Run targeted tests, then SubMiner verification lanes for launcher/runtime changes; update docs/changelog if required by the final behavior change.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verification blocker outside this change: `bun run test:fast` still fails at `scripts/update-aur-package.test.ts` on macOS because `scripts/update-aur-package.sh` uses `mapfile`, which is unavailable in the system Bash 3.x environment used here.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Reworked app-owned YouTube playback to auto-load the default primary subtitle plus a best-effort secondary subtitle at startup instead of forcing the picker modal first. Playback now waits only on primary subtitle load/tokenization readiness, routes startup primary-failure messaging through the configured notification output path, and keeps the YouTube subtitle picker available on demand via a new default `Ctrl+Shift+J` keybinding during active YouTube playback. Updated the runtime/IPC/config plumbing, user-facing help/docs, and added regression coverage for startup auto-load, primary-failure fallback, and manual picker invocation.
<!-- SECTION:FINAL_SUMMARY:END -->
