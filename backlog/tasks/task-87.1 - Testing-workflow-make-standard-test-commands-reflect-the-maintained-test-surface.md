---
id: TASK-87.1
title: >-
  Testing workflow: make standard test commands reflect the maintained test
  surface
status: To Do
assignee: []
created_date: '2026-03-06 03:19'
updated_date: '2026-03-06 03:21'
labels:
  - tests
  - maintainability
milestone: m-0
dependencies: []
references:
  - package.json
  - src/main-entry-runtime.test.ts
  - src/anki-integration/anki-connect-proxy.test.ts
  - src/main/runtime/jellyfin-remote-playback.test.ts
  - src/main/runtime/registry.test.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The current package scripts hand-enumerate a small subset of test files, which leaves the standard green signal misleading. A local audit found 241 test/type-test files under src/ and launcher/, but only 53 unique files referenced by the standard package.json test scripts. This task should redesign the runnable test matrix so maintained tests are either executed by the standard commands or intentionally excluded through a documented rule, instead of silently drifting out of coverage.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The repository has a documented and reproducible test matrix for standard development commands, including which suites belong in the default lane versus slower or environment-specific lanes.
- [ ] #2 The standard test entrypoints stop relying on a brittle hand-maintained allowlist for the currently covered unit and integration suites, or an explicit documented mechanism exists that prevents silent omission of new tests.
- [ ] #3 Representative tests that were previously outside the standard lane from src/main/runtime, src/anki-integration, and entry/runtime surfaces are executed by an automated command and included in the documented matrix.
- [ ] #4 Documentation for contributors explains which command to run for fast verification, full verification, and environment-specific verification.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inventory the current test surface under src/ and launcher/ and compare it to package.json scripts to classify fast, full, slow, and environment-specific suites.
2. Replace or reduce the brittle hand-maintained allowlist so new maintained tests do not silently miss the standard matrix.
3. Update contributor docs with the intended fast/full/environment-specific commands.
4. Verify the new matrix by running the relevant commands and by demonstrating at least one previously omitted runtime/Anki/entry test now belongs to an automated lane.
<!-- SECTION:PLAN:END -->
