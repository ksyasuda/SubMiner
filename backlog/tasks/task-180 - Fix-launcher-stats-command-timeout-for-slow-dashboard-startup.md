---
id: TASK-180
title: Fix launcher stats command timeout for slow dashboard startup
status: Done
assignee:
  - codex
created_date: '2026-03-17 15:16'
updated_date: '2026-03-18 05:28'
labels:
  - launcher
  - stats
  - tests
milestone: m-1
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/commands/stats-command.ts
  - /Users/sudacode/projects/japanese/SubMiner/launcher/main.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/launcher/commands/command-modules.test.ts
priority: medium
ordinal: 112500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Address the failing launcher stats startup path so the CLI tolerates the intended slow dashboard startup window instead of timing out early.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 The launcher stats command no longer times out before the intended slow-start window used by tests
- [x] #2 Regression coverage verifies the slower stats startup path succeeds
- [x] #3 The failing launcher stats startup test passes locally
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused launcher command regression that simulates a stats response arriving after the current timeout boundary and expects success.
2. Adjust the stats startup wait timeout in launcher/commands/stats-command.ts to match the intended slow-start tolerance.
3. Re-run the targeted command test, the previously failing launcher/main.test.ts case, and then the full launcher/main.test.ts file.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verified the failing launcher path: launcher/main.test.ts timed out because launcher/commands/stats-command.ts only waited 8000ms for the stats startup response while the supported slow-start test writes the response after 9s.

Raised the stats startup response timeout to 12000ms so attached stats startup tolerates the existing slow cold-start window without changing command flow.

Verification passed: bun test launcher/commands/command-modules.test.ts --test-name-pattern "stats command launches attached app command with response path|stats command returns after startup response even if app process stays running|stats command throws when stats response reports an error"; bun test launcher/main.test.ts --test-name-pattern "stats command tolerates slower dashboard startup before timing out"; bun test launcher/main.test.ts; bun run test:fast.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the launcher stats startup timeout by extending the response-file wait window in launcher/commands/stats-command.ts from 8s to 12s. The command flow was left unchanged; the launcher now simply gives the stats dashboard enough time to report readiness during slower cold starts, which matches the existing supported behavior exercised by launcher/main.test.ts.

Verification passed with the targeted launcher command tests, the previously failing slow-start launcher/main.test.ts case, the full launcher/main.test.ts file, and the full `bun run test:fast` gate.
<!-- SECTION:FINAL_SUMMARY:END -->
