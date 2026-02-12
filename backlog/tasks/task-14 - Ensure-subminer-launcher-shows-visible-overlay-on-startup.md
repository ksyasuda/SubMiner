---
id: TASK-14
title: Ensure subminer launcher shows visible overlay on startup
status: Done
assignee: []
created_date: '2026-02-12 00:22'
updated_date: '2026-02-12 00:23'
labels:
  - bugfix
  - launcher
  - overlay
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The `subminer` launcher starts SubMiner with `--start` but can leave the visible overlay hidden when runtime config defers auto-show (`auto_start_overlay=false`). Update launcher command args to explicitly request visible overlay at startup so script-mode behavior matches user expectations.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Running `subminer <video>` starts SubMiner with startup args that include visible-overlay show intent.
- [x] #2 Launcher startup remains compatible with texthooker-enabled startup and backend/socket args.
- [x] #3 No regressions in existing startup argument construction for texthooker-only mode.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated `subminer` launcher startup args in `startOverlay()` to include `--show-visible-overlay` alongside `--start`.

This makes script-mode startup idempotently request visible overlay presentation instead of depending on runtime config auto-start visibility flags, while preserving existing backend/socket and optional texthooker args.

Scope:
- `subminer` script only.
- No changes to AppImage internal CLI parsing or runtime services.

Validation:
- Verified argument block in `startOverlay()` now includes `--show-visible-overlay` and preserves existing flags.
- Confirmed texthooker-only path (`launchTexthookerOnly`) is unchanged.
<!-- SECTION:FINAL_SUMMARY:END -->
