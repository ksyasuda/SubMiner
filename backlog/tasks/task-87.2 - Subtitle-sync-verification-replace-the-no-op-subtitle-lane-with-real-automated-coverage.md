---
id: TASK-87.2
title: >-
  Subtitle sync verification: replace the no-op subtitle lane with real
  automated coverage
status: To Do
assignee: []
created_date: '2026-03-06 03:19'
updated_date: '2026-03-06 03:21'
labels:
  - tests
  - subsync
milestone: m-0
dependencies: []
references:
  - package.json
  - README.md
  - src/core/services/subsync.ts
  - src/core/services/subsync.test.ts
  - src/subsync/utils.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
SubMiner advertises subtitle syncing with alass and ffsubsync, but the dedicated test:subtitle command currently does not run any tests. There is already lower-level coverage in src/core/services/subsync.test.ts, but the test matrix and contributor-facing commands do not reflect that reality. This task should replace the no-op lane with real verification, align scripts with the existing subsync test surface, and make the user-facing docs honest about how subtitle sync is verified.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The test:subtitle entrypoint runs real automated verification instead of echoing a placeholder message.
- [ ] #2 The subtitle verification lane covers both alass and ffsubsync behavior, including at least one non-happy-path scenario relevant to current functionality.
- [ ] #3 Contributor-facing documentation points to the real subtitle verification command and no longer implies a dedicated test lane exists when it does not.
- [ ] #4 The resulting verification strategy integrates cleanly with the repository-wide test matrix without duplicating or hiding existing subsync coverage.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Audit the existing subtitle-sync test surface, especially src/core/services/subsync.test.ts, and decide whether test:subtitle should reuse or regroup that coverage.
2. Replace the placeholder script with a real automated command and keep the matrix legible alongside TASK-87.1 work.
3. Update README or related docs so the advertised subtitle verification path matches reality.
4. Verify both alass and ffsubsync behavior remain covered by the resulting lane.
<!-- SECTION:PLAN:END -->
