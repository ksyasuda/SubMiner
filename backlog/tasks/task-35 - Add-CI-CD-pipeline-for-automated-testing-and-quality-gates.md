---
id: TASK-35
title: Add CI/CD pipeline for automated testing and quality gates
status: Done
assignee: []
created_date: '2026-02-14 00:57'
updated_date: '2026-02-17 07:36'
labels:
  - infrastructure
  - ci
  - quality
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
CI should focus on build, test, and type-check validation and should not enforce fixed-size implementation ceilings.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 CI is still triggered on `push` and `pull_request` to `main`.
- [x] #2 A canonical test entrypoint is added (`pnpm test`) and executed in CI, or CI explicitly runs equivalent test commands.
- [x] #3 CI focuses on functional validation (build, tests, type checks) without hardcoded size gates.
- [x] #4 Type-checking is explicitly validated in CI and failure behavior is documented (either `tsc --noEmit` or equivalent).
- [x] #5 CI build verification target is defined clearly (current `pnpm run build` or `make build`) and documented.
- [x] #6 PR visibility requirement remains satisfied (workflow check appears on PRs).
- [x] #7 CI scope (Linux-only vs multi-OS matrix) is documented and intentional.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a root `pnpm test` script that runs both `test:config` and `test:core`, or keep CI explicit on these two commands.
2. Add explicit type-check step (`pnpm exec tsc --noEmit`) unless `pnpm run build` is accepted as the intended check.
3. Confirm no hardcoded size gates are treated as mandatory CI quality gates.
4. Clarify CI build verification scope in docs and workflow (current `pnpm run build` vs optional `make build`).
5. Confirm whether security audit remains advisory or hard-fails. Optional: make advisory check non-blocking with explicit comment.
<!-- SECTION:PLAN:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated `.github/workflows/ci.yml` to complete the CI contract without hardcoded size gates: added explicit `pnpm exec tsc --noEmit`, switched test execution to a canonical `pnpm test`, and kept build verification on `pnpm run build` on `ubuntu-latest` for `push`/`pull_request` to `main`. Also removed CI line-count gate enforcement by deleting `check:main-lines*` scripts from `package.json` and removing `scripts/check-main-lines.sh` from the repo. The workflow remains Linux-only by design and continues to show PR checks.
<!-- SECTION:FINAL_SUMMARY:END -->
