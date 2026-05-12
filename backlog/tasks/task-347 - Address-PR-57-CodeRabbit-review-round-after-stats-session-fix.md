---
id: TASK-347
title: Address PR 57 CodeRabbit review round after stats session fix
status: Done
assignee:
  - codex
created_date: '2026-05-12 07:02'
updated_date: '2026-05-12 09:48'
labels:
  - pr-review
  - coderabbit
  - ci
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/57'
modified_files:
  - src/core/services/overlay-window-input.ts
  - src/core/services/overlay-window.test.ts
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Assess and address the 2026-05-12 CodeRabbit review on PR #57 plus the current red GitHub Actions check. Latest comments cover stats session detail token aggregation, Linux fullscreen overlay refresh scheduling, Hyprland title-event polling, malformed Hyprland monitor JSON handling, and JLPT-lock test coverage for name matches.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Still-valid latest CodeRabbit findings on PR #57 are fixed or documented as skipped with rationale.
- [x] #2 CI failure context is inspected and any repo-relevant failing tests or formatting issues are fixed.
- [x] #3 Regression coverage is added for behavior changes where practical before production edits.
- [x] #4 Relevant local verification passes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect failing GitHub Actions log and current code around each latest CodeRabbit finding.
2. Add or update focused regression tests first for behavior changes: stats token aggregation, fullscreen refresh exit cancellation, Hyprland monitor parse failure, and title-only event filtering.
3. Apply minimal production fixes for still-valid findings, plus the subtitle-render duplicate test coverage item.
4. Run targeted tests first, then format/typecheck and broader relevant gates; update the task with results.

Address latest CodeRabbit callback-scope finding: add regression coverage that macOS visible-overlay blur does not invoke onWindowsVisibleOverlayBlur, then split win32/darwin blur handling in src/core/services/overlay-window-input.ts.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-05-12 09:47: Assessed latest PR #57 CodeRabbit comments via gh. Prior review findings in comments are marked addressed by CodeRabbit; latest still-actionable item was macOS visible-overlay blur invoking the Windows-only blur callback in overlay-window-input.ts.

Added regression coverage showing macOS visible-overlay blur leaves onWindowsVisibleOverlayBlur inactive; test failed before production change and passed after splitting win32/darwin handling.

Verification: bun test src/core/services/overlay-window.test.ts; bun run typecheck; git diff --check. PR checks: build-test-audit pass, GitGuardian pass, CodeRabbit pass/skipped, Claude skipped.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed the latest CodeRabbit callback-scope review on PR #57. Windows visible-overlay blur still invokes onWindowsVisibleOverlayBlur and returns without restacking; macOS visible-overlay blur now returns without restacking and without invoking the Windows-only callback. Added focused regression coverage and verified targeted tests, typecheck, diff whitespace, and current PR checks.
<!-- SECTION:FINAL_SUMMARY:END -->
