---
id: TASK-27.1
title: Map component ownership boundaries and migration path before file splitting
status: Done
assignee:
  - backend
created_date: '2026-02-13 17:13'
updated_date: '2026-02-18 04:11'
labels:
  - refactor
  - documentation
dependencies: []
references:
  - docs/architecture.md
documentation:
  - docs/structure-roadmap.md
parent_task_id: TASK-27
priority: high
ordinal: 42000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Create a lightweight inventory of files needing refactoring and their API contracts to prevent accidental coupling regression during splits.

This is a documentation-only task — no code changes. Its output (docs/structure-roadmap.md) is the prerequisite gate for all Phase 2 subtasks.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Inventory source files with >400 LOC and categorize by concern (bootstrap, Anki integration, MPV protocol, renderer, config, services).
- [ ] #2 Document exported API surface for each target file (entry points, exported types, event names, primary callers).
- [ ] #3 List the split sequence from the parent task plan and note any known risks per step.
- [ ] #4 Store results in docs/structure-roadmap.md and link from this task.
- [ ] #5 Include the global smoke test checklist from the parent TASK-27 plan.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

## Simplification Notes

Original task called for named maintainer owners per component, risk-gated migration sequences, and success metrics per slice. This is heavyweight for a solo project where the developer already knows the codebase intimately.

Reduced to: file inventory, API contracts, sequence + risks, and a shared smoke test checklist. The review analysis (in TASK-27 notes) already covers much of what this task would produce — this task captures it in a durable docs file.

Generated docs/structure-roadmap.md with file inventory, task-specific API contracts, split sequence, known risks, and smoke checklist to unlock TASK-27 subtasks.

<!-- SECTION:NOTES:END -->
