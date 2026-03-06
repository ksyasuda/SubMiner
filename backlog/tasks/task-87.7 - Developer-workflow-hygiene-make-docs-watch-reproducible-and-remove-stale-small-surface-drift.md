---
id: TASK-87.7
title: >-
  Developer workflow hygiene: make docs watch reproducible and remove stale
  small-surface drift
status: To Do
assignee: []
created_date: '2026-03-06 03:20'
updated_date: '2026-03-06 03:21'
labels:
  - tooling
  - tech-debt
milestone: m-0
dependencies: []
references:
  - package.json
  - bun.lock
  - src/anki-integration/field-grouping-workflow.ts
documentation:
  - docs/reports/2026-02-22-task-100-dead-code-report.md
parent_task_id: TASK-87
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
The review found a few low-risk but recurring hygiene issues: docs:watch depends on bunx concurrently even though concurrently is not declared in package metadata, and small stale API surface remains after recent refactors, such as unused parameters in field-grouping workflow code. This task should make the developer workflow reproducible and clean up low-risk stale symbols that do not warrant a dedicated architecture task.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 The docs:watch workflow runs through declared project tooling or is rewritten to avoid undeclared dependencies.
- [ ] #2 Small stale symbols or parameters identified during the review outside the main composition-root cleanup are removed without behavior changes.
- [ ] #3 Any contributor-facing command changes are reflected in repository documentation.
- [ ] #4 The cleanup remains scoped to low-risk workflow and hygiene fixes rather than expanding into large architectural refactors.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Fix the docs:watch workflow so it relies on declared project tooling or an equivalent checked-in command path.
2. Clean up low-risk stale symbols surfaced by the review outside the main.ts architecture task, such as unused parameters left behind by refactors.
3. Keep the task scoped: avoid pulling in main composition-root cleanup or larger Anki/runtime refactors.
4. Verify the affected developer commands still work and document any usage changes.
<!-- SECTION:PLAN:END -->
