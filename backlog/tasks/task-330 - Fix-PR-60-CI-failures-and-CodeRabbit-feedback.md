---
id: TASK-330
title: Fix PR 60 CI failures and CodeRabbit feedback
status: Done
assignee:
  - codex
created_date: '2026-05-04 02:50'
updated_date: '2026-05-04 02:59'
labels:
  - ci
  - pr-review
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/60'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Resolve failing GitHub Actions checks and actionable unresolved CodeRabbit review feedback on PR #60 (Persist stats exclusions in DB and fix word metrics filtering). Keep fixes scoped to the PR behavior and preserve existing project patterns.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Failing GitHub Actions checks for PR #60 have an identified root cause and local fix.
- [x] #2 All actionable unresolved CodeRabbit review comments on PR #60 are addressed locally or explicitly documented as non-actionable.
- [x] #3 Relevant local verification passes for the changed code paths.
- [x] #4 Task notes summarize CI failure context, review-comment handling, and any residual verification gaps.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Resolve PR #60 context and inspect GitHub Actions failures with the gh-fix-ci workflow.
2. Fetch unresolved review threads with the gh-address-comments workflow, focusing on CodeRabbit actionable comments.
3. Read the touched files/tests around the failing paths and comments; identify root cause before edits.
4. Apply minimal fixes with regression coverage where appropriate.
5. Run targeted verification first, then broader repo gates as time permits.
6. Update Backlog notes/acceptance criteria with CI/comment outcomes and residual risks.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Resolved PR #60 CI failure by restoring raw `tokensSeen` for session summaries while keeping filtered persisted word counts in aggregate/known-word paths. Addressed CodeRabbit feedback: fixed missing `headword` test fixture binding; paged vocabulary stats past filtered rows; preserved lifetime/rollup totals when retained-session recomputation is partial; emitted flat known-word timeline points for zero-visible-word line gaps; restored localStorage mocks; added rollback/retry behavior for excluded-word store persistence/initialization.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the PR #60 CI failure and addressed actionable CodeRabbit feedback.

Key changes:
- Restored exact Yomitan token counts for session summary metrics while leaving filtered word counts for aggregate and known-word calculations.
- Fixed malformed query test fixtures by binding `headword` into `imm_words` inserts.
- Updated vocabulary stats to page until enough visible rows are collected after post-query filtering.
- Made library/detail/rollup read models preserve lifetime or stored rollup totals when retained-session recomputation is partial, including dashboard rollup-derived word metrics.
- Kept known-word timeline line positions stable by emitting flat points for missing line indexes.
- Made excluded-word persistence rollback on failed writes, allow initialization retries after transient load failures, and restored mocked `localStorage` in tests.

Verification passed:
- `bun run typecheck`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run test:smoke:dist`
- `bun run format:check:src`
- `git diff --check`
<!-- SECTION:FINAL_SUMMARY:END -->
