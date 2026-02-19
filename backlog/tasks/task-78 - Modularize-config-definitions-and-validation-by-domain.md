---
id: TASK-78
title: Modularize config definitions and validation by domain
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - config
  - refactor
  - maintainability
dependencies:
  - TASK-69
  - TASK-72
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Config defaults and resolution are still concentrated in large files (`src/config/definitions.ts`, `src/config/service.ts`). This task splits config schema/defaults/validation into domain modules while preserving a single composed public config API.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Organize by domain (`anki`, `subtitle`, `jellyfin`, `anilist`, `launcher`, `runtime`).
- Keep one composition root that merges domain validators and defaults.
- Prefer small reusable validator helpers over per-field inline logic.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define domain module boundaries for defaults + validation rules.
2. Extract domain validators with shared primitives and warning builder.
3. Compose resolved config in a central aggregator module.
4. Keep `ConfigService` external API stable.
5. Expand tests for domain-level validation and composed config output.
6. Update config docs to reflect module ownership and extension workflow.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Config definitions/validation are split by domain with clear ownership
- [ ] #2 `ConfigService` API remains backward-compatible
- [ ] #3 Validation behavior remains stable under existing tests
- [ ] #4 New domain-level tests prevent regression in future config additions
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Config and core tests pass
- [ ] #2 Documentation updated for config extension workflow
<!-- DOD:END -->

