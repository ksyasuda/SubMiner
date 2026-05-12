---
id: TASK-324
title: Fix mpv playlist changes re-running app warmups
status: Done
assignee: []
created_date: '2026-05-03 07:48'
updated_date: '2026-05-03 07:52'
labels:
  - bug
  - mpv
  - overlay
dependencies: []
references:
  - launcher/
  - src/core/services/mpv.ts
  - src/main/runtime/
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When moving to the next or previous mpv playlist entry, SubMiner should reconnect the existing app/runtime to mpv instead of treating the new video like a fresh app startup. Re-running startup warmups or creating another app session after the first video can interfere with overlay behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Changing to next or previous mpv playlist item reuses the existing app/runtime instead of launching a new app session.
- [x] #2 Startup warmups are not repeated for playlist item changes after the first app startup.
- [x] #3 Overlay behavior remains available after playlist navigation.
- [x] #4 Regression test covers the playlist-change/reconnect path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce the plugin auto-start regression with a failing Lua start-gate test.
2. Update mpv plugin auto-start handling so playlist/file changes with an already-running overlay reuse the existing app path and do not re-arm pause-until-ready warmup.
3. Add changelog fragment and run plugin/launcher verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
RED: `lua scripts/test-plugin-start-gate.lua` failed after changing the duplicate pause-until-ready auto-start expectations; it showed the loading gate was armed twice while overlay was already running.

GREEN: `plugin/subminer/process.lua` now disarms any old ready gate and only reasserts visible overlay state when auto-start fires while `state.overlay_running` is already true.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Updated the mpv Lua plugin auto-start reuse path so a file/playlist load with an already-running overlay no longer re-arms the pause-until-ready tokenization gate.
- Kept the existing app/control command reuse behavior: subsequent auto-starts reassert visible/hidden overlay state without issuing another `--start` subprocess.
- Added a changelog fragment for the mpv playlist overlay reuse fix.

Tests:
- `lua scripts/test-plugin-start-gate.lua` (red before fix, green after)
- `bun run test:plugin:src`
- `bun run changelog:lint`
- `bun run test:launcher:env:src`
<!-- SECTION:FINAL_SUMMARY:END -->
