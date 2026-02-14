---
id: TASK-27
title: >-
  Refactor project structure to reduce architectural complexity and split
  oversized modules
status: To Do
assignee: []
created_date: '2026-02-13 17:13'
updated_date: '2026-02-13 21:07'
labels:
  - 'owner:architect'
  - 'owner:backend'
  - 'owner:frontend'
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Create a phased backlog-backed restructuring plan that keeps current service-oriented architecture while reducing cognitive load from oversized modules and tightening module ownership boundaries.

This initiative should make future feature work easier by splitting high-complexity files, reducing tightly-coupled orchestration, and introducing measurable structural guardrails.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 A phased decomposition plan is defined in task links and references the following target files: src/main.ts, src/anki-integration.ts, src/core/services/mpv-service.ts, src/renderer/*, src/config/*, and src/core/services/*.
- [ ] #2 Tasks are assigned with clear owners and include explicit dependencies so execution can proceed in parallel where safe.
- [ ] #3 Changes are constrained to structural refactors first (no behavior changes until foundational splits are in place).
- [ ] #4 Each subtask includes test/verification expectations (manual or automated) and a rollback-safe checkpoint.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
## Revised Execution Sequence

### Phase 0 — Prerequisites (outside TASK-27 tree)
- **TASK-7** — Extract main.ts global state into AppState container (required before TASK-27.2)
- **TASK-9** — Remove trivial wrapper functions from main.ts (depends on TASK-7; recommended before TASK-27.2 but not blocking)

### Phase 1 — Lightweight Inventory
- **TASK-27.1** — Inventory files >400 LOC, document contracts, define smoke test checklist

### Phase 2 — Sequential Split Wave
Order matters to avoid merge conflicts:
1. **TASK-27.3** — anki-integration.ts split (self-contained, doesn't affect main.ts wiring until facade is stable)
2. **TASK-27.2** — main.ts split (after TASK-7 provides AppState container and 27.3 stabilizes the Anki facade)
3. **TASK-27.4** — mpv-service.ts split (absorbs TASK-8 scope; blocked until 27.1 is done)
4. **TASK-27.5** — renderer positioning.ts split (downscoped; after 27.2 to avoid import-path conflicts)

### Phase 3 — Stabilization
- **TASK-27.6** — Quality gates and CI enforcement

## Smoke Test Checklist (applies to all subtasks)
Every subtask must verify before merging:
- [ ] App starts and connects to MPV
- [ ] Subtitle text appears in overlay
- [ ] Card mining creates a note in Anki
- [ ] Field grouping modal opens and resolves
- [ ] Global shortcuts work (mine, toggle overlay, copy subtitle)
- [ ] Secondary subtitle display works
- [ ] TypeScript compiles with no new errors
- [ ] All existing tests pass (`pnpm test:core && pnpm test:config`)
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
## Review Findings (2026-02-13)

### Key changes from original plan:
1. **Dropped parallel execution of Phase 2** — TASK-27.2 and 27.5 share import paths; 27.2 and 27.3 share main.ts wiring. Sequential order prevents merge conflicts.
2. **Added TASK-7 as external prerequisite** — main.ts has 30+ module-level `let` declarations. Splitting files without a state container first just scatters mutable state.
3. **TASK-8 absorbed into TASK-27.4** — TASK-8 (separate protocol from app logic) and TASK-27.4 (physical file split) overlap significantly. TASK-27.4 now covers both.
4. **TASK-27.5 downscoped** — Renderer is already well-organized (241-line orchestrator, handlers/, modals/, utils/ directories). Only positioning.ts (513 LOC) needs splitting.
5. **Simplified ownership model** — Removed multi-owner ceremony since this is effectively a solo project. Kept labels for categorical tracking only.
6. **Added global smoke test checklist** — No end-to-end or renderer tests exist, so manual verification is the safety net for every subtask.
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Plan task links and ordering are recorded in backlog descriptions.
- [ ] #2 At least 2 independent owners are assigned with explicit labels in subtasks.
<!-- DOD:END -->
