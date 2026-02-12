---
id: TASK-16
title: Revert overlay startup experiment changes and keep renderer fix
status: Done
assignee: []
created_date: '2026-02-12 01:45'
updated_date: '2026-02-12 01:46'
labels:
  - regression
  - overlay
  - launcher
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
User confirmed renderer module-loading fix resolved the broken overlay, but startup experiment changes introduced side effects (e.g., y-s start path re-launch behavior). Revert non-essential auto-start/debugging changes in launcher/plugin/CLI startup flow while preserving renderer ESM fix.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Remove wrapper/plugin auto-start experiment changes that were added during debugging.
- [x] #2 Restore previous y-s start behavior without relaunching a new overlay session from wrapper-managed startup side effects.
- [x] #3 Keep renderer ESM/module-loading fix intact.
- [x] #4 Build and core tests pass after reversion.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Reverted startup experiment changes while preserving the renderer ESM fix. Removed wrapper-forced visible overlay startup and wrapper-managed mpv script opts from `subminer`, restored plugin defaults/behavior (`auto_start=true`) and removed `wrapper_managed` handling from `plugin/subminer.lua` + `plugin/subminer.conf`, and reverted CLI/bootstrap debug-path changes in `src/core/services/cli-command-service.ts` and `src/core/services/startup-service.ts` with matching test updates. Verified `pnpm run build` and full `pnpm run test:core` pass.
<!-- SECTION:FINAL_SUMMARY:END -->
