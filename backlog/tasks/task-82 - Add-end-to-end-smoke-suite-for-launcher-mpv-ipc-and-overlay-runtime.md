---
id: TASK-82
title: Add end-to-end smoke suite for launcher mpv ipc and overlay runtime
status: To Do
assignee: []
created_date: '2026-02-18 11:43'
updated_date: '2026-02-18 11:43'
labels:
  - testing
  - e2e
  - launcher
  - mpv
dependencies:
  - TASK-74
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Current coverage is strong at unit/service level but thin on end-to-end behavior across launcher, mpv socket lifecycle, IPC wiring, and overlay runtime initialization. This task introduces an observable e2e smoke suite to reduce refactor risk.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Start with deterministic smoke flows, not full UI automation.
- Use minimal fixtures and fakeable mpv socket harness where possible.
- Treat this as release gate signal, with clear pass/fail logs.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Define smoke scenarios:
   - launcher command branch sanity
   - mpv socket ready/not-ready handling
   - IPC channel registration + key command dispatch
   - overlay runtime bootstrap lifecycle
2. Build test harness that can run headless in CI/local environments.
3. Add stable fixtures/mocks for mpv and external tool boundaries.
4. Produce concise logs/artifacts for failures.
5. Integrate smoke suite into pre-release validation flow.
6. Document smoke test usage and troubleshooting.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 E2E smoke suite exists and runs in automated workflow
- [ ] #2 Core launcher/mpv/ipc/overlay startup scenarios are covered
- [ ] #3 Failures produce actionable logs/artifacts
- [ ] #4 Smoke suite documented in development/release docs
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Smoke suite passes on baseline branch
- [ ] #2 Release checklist references smoke suite
<!-- DOD:END -->

