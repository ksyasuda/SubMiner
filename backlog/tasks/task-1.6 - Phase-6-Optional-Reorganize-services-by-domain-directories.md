---
id: TASK-1.6
title: 'Phase 6 (Optional): Reorganize services by domain directories'
status: Done
assignee: []
created_date: '2026-02-10 18:46'
updated_date: '2026-02-18 04:11'
labels: []
dependencies:
  - TASK-1.5
references:
  - plan.md
parent_task_id: TASK-1
ordinal: 4000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

If service flattening remains hard to navigate after Phases 1-5, optionally move modules into domain-based folders and update imports.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 A clear go/no-go decision for domain restructuring is documented based on post-phase-5 codebase state.
- [ ] #2 If executed, service modules are reorganized into domain folders with no import or runtime breakage.
- [x] #3 Build and core test commands pass after any directory reorganization.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Assess post-phase-5 directory complexity and determine whether domain reorganization is still justified.
2. If complexity remains acceptable, record a no-go decision and keep current structure stable.
3. If complexity is still problematic, perform import-safe domain reorganization and re-run build/tests.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Decision: no-go on Phase 6 directory reorganization for now. After Phases 1-5, service/module consolidation and test expansion have improved maintainability without introducing a high-risk import churn.

Rationale: preserving path stability now reduces regression risk while Phase 4 smoke validation remains open and large refactor commits are still stabilizing.

Verification baseline remains green (`pnpm run test:core`) with current structure.

<!-- SECTION:NOTES:END -->
