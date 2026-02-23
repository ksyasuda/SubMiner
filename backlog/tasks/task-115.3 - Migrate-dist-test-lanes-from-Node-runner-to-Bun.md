---
id: TASK-115.3
title: Migrate dist test lanes from Node runner to Bun
status: Done
assignee: []
created_date: '2026-02-23 04:27'
updated_date: '2026-02-23 04:36'
labels:
  - test
  - tooling
dependencies:
  - TASK-115.2
references:
  - package.json
  - docs/development.md
  - docs/installation.md
documentation:
  - docs/plans/2026-02-23-bun-only-toolchain-migration.md
parent_task_id: TASK-115
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace Node-runner dist test commands with Bun-based dist verification while preserving existing test scope and confidence in compiled artifact behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Dist test scripts execute via Bun instead of node --test while preserving the current dist test file coverage.
- [x] #2 Dist smoke and full dist verification lanes pass under Bun-based commands.
- [x] #3 Documentation for dist verification commands reflects the updated Bun-based workflow.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated `package.json` dist lanes from `node --test` to `bun test` (`test:config:dist`, `test:config:smoke:dist`, `test:core:dist`, `test:core:smoke:dist`).

Resolved Bun dist compatibility gap by injecting `SafeStorageLike` into AniList token store and replacing brittle Electron global monkey-patching in `anilist-token-store` tests.

Validation: `bun run build && bun run test:config:dist && bun run test:core:dist && bun run test:smoke:dist` all pass.
<!-- SECTION:NOTES:END -->
