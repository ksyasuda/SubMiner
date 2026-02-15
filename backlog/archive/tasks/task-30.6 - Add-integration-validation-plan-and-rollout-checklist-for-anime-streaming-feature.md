---
id: TASK-30.6
title: >-
  Add integration validation plan and rollout checklist for anime streaming
  feature
status: To Do
assignee: []
created_date: '2026-02-13 18:34'
updated_date: '2026-02-13 18:34'
labels: []
dependencies: []
parent_task_id: TASK-30
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Create a concrete validation task that defines end-to-end acceptance checks for config loading, resolver behavior, IPC contract correctness, UI flow, and mpv-triggered launch. The checklist should be actionable and align with existing project conventions so completion can be verified without guesswork.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Define test scenarios for config success/failure cases, including invalid provider config and feature disabled mode.
- [ ] #2 Define search/list/resolve API contract tests and error-code assertions (empty, timeout, auth error, no playable URL).
- [ ] #3 Define renderer UX checks for modal state transitions, loading indicators, empty results, selection, and play invocation.
- [ ] #4 Define in-mpv trigger checks for command/message pathway and fallback behavior when modal disabled/unavailable.
- [ ] #5 Define manual smoke steps for end-to-end play from query to mpv playback using at least one configured source.
- [ ] #6 Document expected logs/telemetry markers for troubleshooting and rollback.
- [ ] #7 Define rollback criteria and what constitutes safe partial completion.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Phase 6 — Validation: rollout, smoke tests, and release readiness checklist
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Checklist covers happy-path and failure-path for each task dependency.
- [ ] #2 Verification steps are executable without external tooling assumptions.
- [ ] #3 No task can be marked done without explicit evidence fields filled in.
- [ ] #4 Rollout and fallback behavior are documented per deployment/release phase.
<!-- DOD:END -->
