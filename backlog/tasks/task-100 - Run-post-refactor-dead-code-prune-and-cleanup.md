---
id: TASK-100
title: Run post-refactor dead code prune and cleanup
status: To Do
assignee: []
created_date: '2026-02-21 07:15'
updated_date: '2026-02-21 07:15'
labels:
  - cleanup
  - maintainability
  - refactor
dependencies:
  - TASK-96
  - TASK-97
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Major refactors likely left unused exports/helpers and stale compatibility code. Perform a deliberate dead-code sweep with regression safety.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Run unused-export scans (`ts-prune` or equivalent) and collect candidate list.
2. Manually verify each candidate to avoid false positives from dynamic loading/IPC wiring.
3. Remove or inline confirmed dead code; simplify import surfaces/barrels accordingly.
4. Delete stale helper paths retained only for pre-composer wiring.
5. Add or update regression tests where removal could alter behavior.
6. Run verification gate: `bun run build`, `bun run test:config:dist`, `bun run test:core:dist`.
7. Publish cleanup report: removed files/exports, kept exceptions, and rationale.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Confirmed dead code removed without behavior regressions.
- [ ] #2 Remaining flagged candidates either resolved or documented with justification.
- [ ] #3 Build and core/config suites pass after cleanup.
- [ ] #4 Import graph complexity reduced (fewer exports/entrypoints where applicable).
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Dead-code report attached in task notes.
- [ ] #2 Regression test updates included for risky removals.
- [ ] #3 Verification gate commands complete successfully.
<!-- DOD:END -->

