---
id: TASK-347
title: Address PR 57 CodeRabbit review round after stats session fix
status: In Progress
assignee:
  - codex
created_date: '2026-05-12 07:02'
updated_date: '2026-05-12 07:02'
labels:
  - pr-review
  - coderabbit
  - ci
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess and address the 2026-05-12 CodeRabbit review on PR #57 plus the current red GitHub Actions check. Latest comments cover stats session detail token aggregation, Linux fullscreen overlay refresh scheduling, Hyprland title-event polling, malformed Hyprland monitor JSON handling, and JLPT-lock test coverage for name matches.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Still-valid latest CodeRabbit findings on PR #57 are fixed or documented as skipped with rationale.
- [ ] #2 CI failure context is inspected and any repo-relevant failing tests or formatting issues are fixed.
- [ ] #3 Regression coverage is added for behavior changes where practical before production edits.
- [ ] #4 Relevant local verification passes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect failing GitHub Actions log and current code around each latest CodeRabbit finding.
2. Add or update focused regression tests first for behavior changes: stats token aggregation, fullscreen refresh exit cancellation, Hyprland monitor parse failure, and title-only event filtering.
3. Apply minimal production fixes for still-valid findings, plus the subtitle-render duplicate test coverage item.
4. Run targeted tests first, then format/typecheck and broader relevant gates; update the task with results.
<!-- SECTION:PLAN:END -->
