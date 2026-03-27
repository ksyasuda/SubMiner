---
id: TASK-239.1
title: Add profile-aware config foundations and profile selection flow
status: To Do
assignee: []
created_date: '2026-03-26 20:49'
labels:
  - feature
  - config
  - launcher
  - ux
milestone: m-2
dependencies: []
references:
  - src/config/service.ts
  - src/config/load.ts
  - launcher/config.ts
  - src/main.ts
parent_task_id: TASK-239
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Introduce the foundation for local multi-profile use so users can keep separate setups for different workflows without hand-editing or swapping config files manually. Keep the first slice intentionally narrow: named local profiles, explicit selection, separate config/data paths, and safe migration from the current single-profile setup. Do not couple this task to cloud sync or remote profile sharing.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 Users can create/select a named local profile and launch SubMiner against that profile explicitly.
- [ ] #2 Each profile uses separate config and data storage paths for settings and profile-scoped runtime state that should not bleed across workflows.
- [ ] #3 Existing single-profile users migrate safely to a default profile without losing settings.
- [ ] #4 The active profile is visible in the launcher/app surface where it materially affects user behavior.
- [ ] #5 Tests cover profile resolution, migration/defaulting behavior, and at least one end-to-end selection path.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Design a minimal profile storage layout and resolution strategy that works for launcher and desktop runtime entrypoints.
2. Add profile selection plumbing before changing feature behavior inside individual services.
3. Migrate config/data-path resolution to be profile-aware while preserving a safe default-profile fallback.
4. Verify with config/launcher tests plus targeted runtime coverage.
<!-- SECTION:PLAN:END -->
