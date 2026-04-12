---
id: TASK-286.2
title: 'Assess and address PR #49 next CodeRabbit review round'
status: Done
assignee: []
created_date: '2026-04-12 02:50'
updated_date: '2026-04-12 02:52'
labels:
  - bug
  - code-review
  - release
  - testing
dependencies: []
references:
  - .github/workflows/prerelease.yml
  - src/prerelease-workflow.test.ts
  - src/core/services/overlay-shortcut.ts
parent_task_id: TASK-286
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Track the next unresolved CodeRabbit review threads on PR #49 after commit 62ad77dc and resolve the still-valid follow-up issues while documenting stale repeats.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 All still-actionable CodeRabbit comments in the latest PR #49 round are fixed or explicitly shown stale with evidence.
- [x] #2 Regression coverage is updated for any workflow or test changes made in this round.
- [x] #3 Relevant verification passes for the touched workflow and prerelease test changes.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Assessed latest unresolved CodeRabbit round on PR #49. `src/core/services/overlay-shortcut.ts` comment is stale: `registerOverlayShortcuts()` returns `hasConfiguredOverlayShortcuts(shortcuts)`, so runtime registration is not hard-coded false.

Added exact, line-ending-agnostic prerelease tag trigger assertions in `src/prerelease-workflow.test.ts` and a regression asserting `bun run test:env` sits in the prerelease quality gate before source coverage.

Updated `.github/workflows/prerelease.yml` quality-gate to run `bun run test:env` after `bun run test:fast`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Assessed the latest CodeRabbit round for PR #49. Left the `overlay-shortcut.ts` thread open as stale with code evidence, tightened prerelease workflow trigger coverage, and added the missing `test:env` step to the prerelease quality gate. Verification: `bun test src/prerelease-workflow.test.ts`; `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
