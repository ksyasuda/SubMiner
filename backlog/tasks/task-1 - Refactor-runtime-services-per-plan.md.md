---
id: TASK-1
title: Refactor runtime services per plan.md
status: Done
assignee: []
created_date: '2026-02-10 18:46'
updated_date: '2026-02-11 03:35'
labels: []
dependencies: []
references:
  - plan.md
ordinal: 1000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Execute the SubMiner refactoring initiative documented in plan.md to reduce thin abstractions, consolidate service boundaries, fix known quality issues, and increase test coverage while preserving current behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Phase-based execution tasks are created and linked under this initiative.
- [x] #2 Each phase task includes clear, testable outcomes aligned with plan.md.
- [x] #3 Implementation proceeds with build/test verification checkpoints after each completed phase.
- [x] #4 Main behavior remains stable for startup, overlay, IPC, CLI, and tokenizer flows throughout refactor.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Created initiative subtasks TASK-1.1 through TASK-1.6 with phase-aligned acceptance criteria and sequential dependencies.

Completed TASK-1.1 (Phase 1 thin-wrapper removal) with green build/core tests.

Completed TASK-1.2 (Phase 2 DI adapter consolidation) with successful build and core test verification checkpoint.

Completed TASK-1.5 (critical behavior tests) with expanded tokenizer/mpv/subsync/CLI coverage and green core test suite.

Completed TASK-1.6 with documented no-go decision for optional domain-directory reorganization (kept current structure; tests remain green).

TASK-1.4 remains the only open phase, blocked on interactive desktop smoke checks that cannot be fully validated in this headless environment.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed the plan.md refactor initiative across Phases 1-5 and optional Phase 6 decisioning: removed thin wrappers, consolidated DI adapters and related services, fixed targeted runtime correctness issues, expanded critical behavior test coverage, and kept build/core tests green throughout. Final runtime smoke checks (start/toggle/trigger-field-grouping/stop) passed in this headless environment, with known limitation that visual overlay rendering itself was not directly inspectable.
<!-- SECTION:FINAL_SUMMARY:END -->
