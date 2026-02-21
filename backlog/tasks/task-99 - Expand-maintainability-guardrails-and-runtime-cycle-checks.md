---
id: TASK-99
title: Expand maintainability guardrails and runtime cycle checks
status: To Do
assignee: []
created_date: '2026-02-21 07:15'
updated_date: '2026-02-21 07:15'
labels:
  - quality
  - architecture
  - ci
dependencies:
  - TASK-96
  - TASK-97
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current guardrails cover `main.ts` fan-in and broad file budgets. Add targeted checks for newly extracted hotspots and runtime import-cycle detection.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Rebaseline file-size budgets for known hotspots (including post-TASK-96 files).
2. Extend `check:file-budgets` config to enforce thresholds on new modules and key test files.
3. Add runtime dependency cycle detection for `src/main/runtime/**` (script + CI hook).
4. Wire new checks into local gate commands and CI workflow.
5. Add failure guidance output so contributors can remediate quickly.
6. Run gates on current branch and capture pass/fail evidence in task notes.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Budget thresholds include new hotspot modules and prevent silent growth.
- [ ] #2 Runtime cycle check detects at least one injected fixture cycle in tests.
- [ ] #3 CI runs both checks and fails fast on violations.
- [ ] #4 Contributor guidance exists for fixing guardrail failures.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Guardrail scripts and CI wiring merged with passing checks.
- [ ] #2 Documentation updated with local commands and expected outputs.
- [ ] #3 Baseline numbers recorded to justify thresholds.
<!-- DOD:END -->

