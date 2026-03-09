---
id: TASK-148
title: Fix Windows plugin env binary override resolution
status: Done
assignee:
  - codex
created_date: '2026-03-09 00:00'
updated_date: '2026-03-09 00:00'
labels:
  - windows
  - plugin
  - regression
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Fix the mpv plugin's Windows binary override lookup so `SUBMINER_BINARY_PATH` still resolves when `SUBMINER_APPIMAGE_PATH` is unset. The current Lua resolver builds an array with a leading `nil`, which causes `ipairs` iteration to stop before the later Windows override candidate.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 `scripts/test-plugin-binary-windows.lua` passes the env override regression that expects `.exe` suffix resolution from `SUBMINER_BINARY_PATH`.
- [x] #2 Existing plugin start/binary test gate stays green after the fix.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Updated `plugin/subminer/binary.lua` so env override lookup checks `SUBMINER_APPIMAGE_PATH` and `SUBMINER_BINARY_PATH` sequentially instead of via a Lua array literal that truncates at the first `nil`. This restores Windows `.exe` suffix resolution for `SUBMINER_BINARY_PATH` when the AppImage env var is unset.

Verification:

- `lua scripts/test-plugin-binary-windows.lua`
- `bun run test:plugin:src`

<!-- SECTION:FINAL_SUMMARY:END -->
