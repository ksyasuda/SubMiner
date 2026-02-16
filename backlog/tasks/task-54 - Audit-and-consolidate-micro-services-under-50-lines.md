---
id: TASK-54
title: Audit and consolidate micro-services under 50 lines
status: In Progress
assignee: []
created_date: '2026-02-16 04:47'
updated_date: '2026-02-16 04:59'
labels: []
dependencies: []
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/
priority: medium
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
- [ ] #1 Audit all services under 50 lines in src/core/services/
- [ ] #2 Identify logical groupings for consolidation
- [ ] #3 Merge related micro-services into cohesive modules
- [ ] #4 Update all imports across codebase
- [ ] #5 Update barrel exports in services/index.ts
- [ ] #6 Run full test suite to ensure no regressions
<!-- AC:END -->
