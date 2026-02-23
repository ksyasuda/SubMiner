---
id: TASK-115.4
title: Migrate Node-invoked utility script commands to Bun
status: Done
assignee: []
created_date: '2026-02-23 04:27'
updated_date: '2026-02-23 04:36'
labels:
  - tooling
  - scripts
dependencies:
  - TASK-115.2
references:
  - package.json
  - scripts/get_frequency.ts
  - scripts/test-yomitan-parser.ts
documentation:
  - docs/plans/2026-02-23-bun-only-toolchain-migration.md
parent_task_id: TASK-115
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Update utility command entrypoints that still rely on Node invocation so maintenance and diagnostics workflows can be run entirely through Bun.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Utility script commands that currently invoke Node have Bun-based equivalents as the default project commands.
- [x] #2 Electron-targeted utility flows remain functional after command migration.
- [x] #3 Contributor docs reflect the updated Bun-based utility command usage.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated `generate:config-example` command from `node dist/generate-config-example.js` to `bun dist/generate-config-example.js`.

Updated launcher smoke fixture executables from `#!/usr/bin/env node` to `#!/usr/bin/env bun` so smoke tests do not assume Node presence.

Validation: `bun run test:launcher:smoke:src` and `bun run generate:config-example` pass with Bun-first command path.
<!-- SECTION:NOTES:END -->
