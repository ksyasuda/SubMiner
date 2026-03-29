---
id: TASK-243
title: 'Assess and address PR #36 latest CodeRabbit review round'
status: Done
assignee: []
created_date: '2026-03-29 07:39'
updated_date: '2026-03-29 07:41'
labels:
  - code-review
  - pr-36
dependencies: []
references:
  - 'https://github.com/ksyasuda/SubMiner/pull/36'
priority: high
ordinal: 3600
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Inspect the latest CodeRabbit review round on PR #36, verify each actionable comment against the current branch, implement the confirmed fixes, and verify the touched paths.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Confirmed review comments are implemented or explicitly deferred with rationale.
- [ ] #2 Touched paths are verified with the smallest sufficient test/build lane.
- [ ] #3 Current PR feedback is reduced to resolved or intentionally deferred suggestions.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Addressed the confirmed latest CodeRabbit review items on PR #36. `scripts/run-coverage-lane.ts` now uses the Bun-style `import.meta.main` entrypoint check with a local ts-ignore to preserve the repo's CommonJS typecheck settings. `src/core/services/immersion-tracker/maintenance.ts` no longer shadows the imported `nowMs` helper in retention functions. `src/main.ts` now centralizes the startup-mode predicates behind a shared helper and releases `resolvedSource.cleanup` on the cached-subtitle fast path so materialized sources do not leak.
<!-- SECTION:FINAL_SUMMARY:END -->
