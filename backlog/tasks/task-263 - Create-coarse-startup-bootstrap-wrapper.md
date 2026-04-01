---
id: TASK-263
title: Create coarse startup bootstrap wrapper
status: To Do
assignee: []
created_date: '2026-03-31 17:21'
labels:
  - refactor
  - main
  - startup
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Move the large createMainStartupRuntime construction and its self-reference handling out of src/main.ts into a coarse startup bootstrap wrapper. Keep behavior identical and shrink the startup section materially.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 New wrapper module(s) live under src/main/
- [ ] #2 Wrapper owns createMainStartupRuntime construction and self-reference handling
- [ ] #3 src/main.ts startup section becomes materially smaller
- [ ] #4 Files stay under 500 LOC
- [ ] #5 Focused tests cover the wrapper if useful
- [ ] #6 No runtime behavior changes
<!-- AC:END -->
