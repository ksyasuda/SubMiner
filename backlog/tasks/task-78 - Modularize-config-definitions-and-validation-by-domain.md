---
id: TASK-78
title: Modularize config definitions and validation by domain
status: Done
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-22 07:49'
labels:
  - config
  - refactor
  - maintainability
dependencies:
  - TASK-69
  - TASK-72
priority: high
ordinal: 78000
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
- [x] #1 Config definitions/validation are split by domain with clear ownership
- [x] #2 `ConfigService` API remains backward-compatible
- [x] #3 Validation behavior remains stable under existing tests
- [x] #4 New domain-level tests prevent regression in future config additions
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: Started execution via writing-plans/executing-plans workflow in opencode session `opencode-task78-config-domain-20260221T235604Z-p9x2`.

2026-02-22: Refactored `src/config/definitions.ts` into a composed facade over domain modules under `src/config/definitions/` (`defaults-*.ts`, `options-*.ts`, `runtime-options.ts`, `template-sections.ts`, `shared.ts`) while preserving exported API names and `ConfigService` behavior.

Added domain-level regression tests in `src/config/definitions/domain-registry.test.ts` for critical registry paths, duplicate-path guard, template section key uniqueness, and per-domain contributor coverage.

Updated contributor docs (`docs/development.md`) with config extension workflow by domain ownership and composition entrypoint guidance.

Verification: `bun test src/config/definitions/domain-registry.test.ts` (3 pass), `bun run test:config:src` (49 pass), `bun run test:core:src` (219 pass, 6 skip).

`make generate-config` currently blocked by pre-existing TypeScript errors in `src/main/state.test.ts` missing exports from `src/main/state` (outside TASK-78 scope).

Follow-up: wired `src/config/definitions/domain-registry.test.ts` into `package.json` config test lanes (`test:config:src` + `test:config:dist`). Re-ran `bun run test:config:src` => 52 pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Modularized config definition ownership by introducing domain-specific defaults, option registry builders, runtime-option metadata, and template-section modules under `src/config/definitions/`, with `src/config/definitions.ts` preserved as the single composed public API. Added domain-level registry tests and updated contributor docs for the new config extension workflow; config/core source test lanes pass.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Config and core tests pass
- [x] #2 Documentation updated for config extension workflow
<!-- DOD:END -->
