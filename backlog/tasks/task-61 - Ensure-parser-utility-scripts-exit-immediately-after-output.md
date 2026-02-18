---
id: TASK-61
title: Ensure parser utility scripts exit immediately after output
status: Done
assignee: []
created_date: '2026-02-16 22:35'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
ordinal: 24000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update `scripts/test-yomitan-parser.ts` and `scripts/get_frequency.ts` so they clean up Electron parser resources and terminate immediately after producing results, avoiding hangs.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `scripts/test-yomitan-parser.ts` exits promptly after printing output.
- [x] #2 `scripts/get_frequency.ts` exits promptly after printing output.
- [x] #3 Electron-related resources (parser window/app loop) are cleaned up on both success and error paths.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added deterministic shutdown to both utility scripts. `scripts/get_frequency.ts` now destroys parser windows in a `finally` block, calls `app.quit()` when Electron is loaded, and uses explicit `.then/.catch` exits so the process terminates immediately after output with correct exit codes. `scripts/test-yomitan-parser.ts` now mirrors this pattern with runtime cleanup (`shutdownYomitanRuntime`) and explicit process exit handling.
<!-- SECTION:FINAL_SUMMARY:END -->
