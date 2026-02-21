---
id: TASK-74
title: Add launcher regression tests for config discovery and command branching
status: Done
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-21 20:21'
labels:
  - launcher
  - tests
  - regression
dependencies:
  - TASK-70
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Launcher currently has no direct test coverage for config discovery behavior and key command branching paths. This leaves wrapper flows vulnerable to regressions when refactoring config or process-control logic.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add launcher-focused test harness (unit/integration-lite) under existing project test runner.
2. Cover config discovery matrix:
   - `XDG_CONFIG_HOME` overrides
   - `SubMiner` vs `subminer`
   - `.jsonc` vs `.json` preference
3. Cover CLI branch behavior (`doctor`, `config path/show`, `mpv status/socket`, jellyfin command routing).
4. Mock fs/process/spawn boundaries to keep tests deterministic.
5. Wire tests into existing `test:core:dist` or dedicated launcher test command.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Launcher config discovery paths are regression-tested
- [x] #2 Core launcher command branches are regression-tested
- [x] #3 Tests run in CI/local test gate without external process dependencies
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Launcher tests included in standard test workflow
- [x] #2 Documented how to run launcher tests locally
<!-- DOD:END -->
