---
id: TASK-93
title: Synchronize TASK-85 closure tracking and child-task status
status: To Do
assignee: []
created_date: '2026-02-20 12:06'
updated_date: '2026-02-20 12:06'
labels:
  - process
  - refactor
  - maintainability
dependencies:
  - TASK-85
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`TASK-85` progress and closure metadata are out of sync with actual completed work, making prioritization noisy. Normalize `TASK-85` progress notes, acceptance checklist state, and child-task linkage so remaining scope is explicit and actionable.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Audit completed work against `TASK-85` Action Steps/AC/DoD and map each item to evidence (files/tests/docs).
2. Add explicit child-task linkage in `TASK-85` progress notes for open scopes (`TASK-94`, `TASK-95` + existing related tasks).
3. Update `TASK-85` AC checkboxes only where evidence is complete; leave unresolved items unchecked with short reason.
4. Update `TASK-85` DoD checkboxes only where evidence is complete; leave unresolved items unchecked with short reason.
5. Add a “Remaining Scope” section in `TASK-85` with ordered next actions and owning ticket IDs.
6. Validate consistency by re-reading `TASK-85` plus linked child tickets in one pass.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `TASK-85` status metadata and updated date reflect latest reality.
- [ ] #2 Each unresolved `TASK-85` AC/DoD item has explicit owning child ticket(s).
- [ ] #3 `TASK-85` has a clear, ordered “Remaining Scope” list with next execution order.
- [ ] #4 No stale or contradictory progress statements remain in `TASK-85`.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `TASK-85` reads as a reliable source-of-truth for what is done vs pending.
- [ ] #2 Child-ticket mapping for pending work is complete and unambiguous.
<!-- DOD:END -->
