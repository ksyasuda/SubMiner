---
id: TASK-262
title: Create overlay UI bootstrap input helper
status: To Do
assignee: []
created_date: '2026-03-31 17:04'
labels:
  - refactor
  - main
  - overlay-ui
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a coarse input-builder/helper module to reduce the large createOverlayUiRuntime(...) callsite in src/main.ts without changing runtime behavior. Do not edit src/main.ts in this task.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 New helper module(s) live under src/main/
- [ ] #2 Helper accepts grouped overlay UI/domain inputs instead of giant inline literals
- [ ] #3 Helper keeps files under 500 LOC
- [ ] #4 Optional focused tests added if useful
- [ ] #5 No runtime behavior changes
<!-- AC:END -->
