---
id: TASK-229
title: 'Address PR #31 final CodeRabbit picker test follow-up'
status: Done
assignee:
  - codex
created_date: '2026-03-24 04:27'
updated_date: '2026-03-24 06:41'
labels:
  - pr-review
  - coderabbit
dependencies: []
references:
  - >-
    PR #31 feat: add app-owned YouTube subtitle flow with absPlayer-style
    parsing
  - >-
    CodeRabbit comment on src/renderer/modals/youtube-track-picker.test.ts
    global restoration / harness duplication
priority: medium
ordinal: 145500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the remaining CodeRabbit comment on the YouTube picker test file by restoring absent globals correctly and reducing repeated test harness setup so global stubbing is consistent and isolated.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Picker tests restore `window`, `document`, and `CustomEvent` without leaving undefined-valued globals behind.
- [x] #2 Repeated picker test setup is consolidated enough to remove the current review complaint.
- [x] #3 Relevant picker tests pass and PR thread is updated.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression around global restoration semantics in the YouTube picker test harness.
2. Extract shared DOM/environment helpers and restore logic using delete when globals were originally absent.
3. Re-run focused tests and typecheck, then commit/push and reply on the PR thread.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Latest CodeRabbit comment targets youtube-track-picker.test.ts harness cleanup and correct restoration of global properties.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed the last PR #31 CodeRabbit comment by refactoring the YouTube picker test harness to use shared DOM/env helpers, restoring absent globals via delete semantics, adding a regression for cleanup behavior, and pushing commit 039e2f56 with focused picker tests plus typecheck passing.
<!-- SECTION:FINAL_SUMMARY:END -->
