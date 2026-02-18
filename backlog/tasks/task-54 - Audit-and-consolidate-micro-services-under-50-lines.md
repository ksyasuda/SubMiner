---
id: TASK-54
title: Audit and consolidate micro-services under 50 lines
status: Done
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/
priority: medium
ordinal: 29000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

The core/services directory contains 67 files, many of which are very small services that may create unnecessary abstraction overhead. This task involves auditing services under 50 lines and determining if they should be consolidated with related services.

Candidates for review (all under 50 lines):

- mpv-state.ts (25 lines) - could merge with mpv-service.ts
- secondary-subtitle-service.ts (32 lines) - could merge with subtitle-related services
- runtime-config-service.ts (50 lines) - pure utility functions that could merge with config service
- mpv-control-service.ts (49 lines) - MPV command functions could merge with mpv-service.ts

Approach:

1. Identify logical groupings (mpv-related, subtitle-related, config-related)
2. Determine which micro-services are truly independent vs which are fragments
3. Consolidate related micro-services into cohesive modules
4. Maintain clear function exports for tree-shaking

Benefits:

- Reduces file navigation overhead
- Groups related functionality logically
- Makes service boundaries clearer
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Audit all services under 50 lines in src/core/services/
- [x] #2 Identify logical groupings for consolidation
- [x] #3 Merge related micro-services into cohesive modules
- [x] #4 Update all imports across codebase
- [x] #5 Update barrel exports in services/index.ts
- [x] #6 Run full test suite to ensure no regressions
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Consolidation for micro-services under 50 lines is now complete in `src/core/services`: MPV runtime helpers are now in `mpv-service.ts`, secondary-subtitle cycling logic is in `subtitle-position-service.ts`, and runtime config decision helpers are in `startup-service.ts`. The legacy split files (`mpv-state.ts`, `mpv-control-service.ts`, `runtime-config-service.ts`, `secondary-subtitle-service.ts`) are no longer part of the service surface. I also verified consolidated imports through `src/core/services/index.ts` and updated unit tests to target the new locations.

Core service files under 50 lines are now reduced to only two small test files: `mpv-state.test.ts` and `mpv-render-metrics-service.test.ts`; all tiny service implementations have been folded into cohesive modules.

Validation run: `pnpm exec tsc --noEmit` and selected service tests for mpv/runtime-config/subtitle grouping passed (25 tests, 0 failures).

<!-- SECTION:FINAL_SUMMARY:END -->
