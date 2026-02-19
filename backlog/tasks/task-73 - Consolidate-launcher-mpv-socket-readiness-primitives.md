---
id: TASK-73
title: Consolidate launcher mpv socket readiness primitives
status: To Do
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-18 11:35'
labels:
  - launcher
  - mpv
  - refactor
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launcher contains multiple overlapping MPV readiness/polling helpers (`waitForSocket`, `waitForPathExists`, `waitForUnixSocketReady`) with mixed timeout/poll semantics. Consolidation needed for clarity and lower maintenance.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define one canonical socket readiness contract (existence + connectability + timeout).
2. Replace duplicate polling helpers with a single implementation.
3. Update all callsites (`mpv status`, idle launch flow, subtitle loading dependencies).
4. Keep user-visible behavior and exit codes stable.
5. Add focused tests for timeout and success/failure edge cases.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Single canonical MPV socket readiness helper remains
- [ ] #2 All launcher callsites use unified helper
- [ ] #3 Behavior/exit code compatibility maintained for CLI flows
- [ ] #4 Test coverage added for readiness timeout behavior
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Launcher tests pass
- [ ] #2 No dead helper functions remain
<!-- DOD:END -->

