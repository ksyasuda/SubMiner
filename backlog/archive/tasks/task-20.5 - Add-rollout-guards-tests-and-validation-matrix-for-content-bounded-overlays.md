---
id: TASK-20.5
title: 'Add rollout guards, tests, and validation matrix for content-bounded overlays'
status: To Do
assignee: []
created_date: '2026-02-12 09:43'
labels: []
dependencies: []
parent_task_id: TASK-20
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add safety controls and verification coverage for the new content-bounded overlay architecture and secondary top-bar window.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Feature flag or equivalent rollout guard exists for switching to new sizing/window behavior.
- [ ] #2 Service-level/unit tests cover bounds clamping, jitter thresholding, invalid measurement fallback, and per-layer updates.
- [ ] #3 Manual validation checklist documents and verifies wrap/no-wrap, style changes, monitor moves, tracker churn, modal interactions, and simultaneous overlay states.
- [ ] #4 Regression checks confirm existing single-layer and startup/shutdown behavior remain stable.
- [ ] #5 Task includes explicit pass/fail notes from validation run(s).
<!-- AC:END -->
