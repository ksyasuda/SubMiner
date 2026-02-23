---
id: TASK-117
title: Fix mpv plugin overlay start gate when SubMiner is not running
status: Done
assignee:
  - codex
created_date: '2026-02-23 04:49'
updated_date: '2026-02-23 04:50'
labels:
  - bug
  - plugin
  - regression
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Recent `plugin/subminer.lua` change added an IPC readiness gate in `start_overlay()` that rejects start when no SubMiner process exists yet. This blocks manual start (`y-s`) and script-message start even when binary path overrides are valid (`SUBMINER_APPIMAGE_PATH`, `SUBMINER_BINARY_PATH`, `binary_path`).

Need to restore expected behavior:
- allow cold start when process is not running,
- preserve mismatch/misconfiguration checks when process is already running,
- add regression coverage for the new start gate behavior.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add failing regression test covering cold-start path from plugin start gate.
2. Patch `start_overlay()` gate logic so process-not-running does not block initial start.
3. Keep existing IPC mismatch/missing-socket checks for running-process path.
4. Re-run plugin validation and focused tests.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `y-s` / `subminer-start` can launch overlay when SubMiner is not yet running and binary is available.
- [x] #2 Plugin still surfaces clear refusal for true IPC/socket mismatches after process is running.
- [x] #3 Regression test exists and fails before fix / passes after fix.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
- Patched `plugin/subminer.lua` start gate: `start_overlay()` now treats `SubMiner process not running` as a valid cold-start condition and only refuses start for other IPC readiness failures.
- Added regression harness `scripts/test-plugin-start-gate.lua` to load plugin with mocked mpv APIs and assert `subminer-start` triggers `--start` command for cold-start path.
- Validation:
  - red: `SUBMINER_APPIMAGE_PATH=/tmp/subminer-binary lua scripts/test-plugin-start-gate.lua` (failed before patch with expected cold-start assertion),
  - green: same command passes after patch,
  - `luac -p plugin/subminer.lua`,
  - `bun test launcher/mpv.test.ts`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed mpv plugin start regression by allowing cold-start launch when SubMiner process is not yet running, while preserving existing refusal path for non-cold-start IPC readiness errors. Added an executable Lua regression harness for this path and validated with syntax + focused launcher MPV tests.
<!-- SECTION:FINAL_SUMMARY:END -->
