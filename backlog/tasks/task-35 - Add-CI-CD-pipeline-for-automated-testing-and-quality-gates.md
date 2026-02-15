---
id: TASK-35
title: Add CI/CD pipeline for automated testing and quality gates
status: To Do
assignee: []
created_date: '2026-02-14 00:57'
labels:
  - infrastructure
  - ci
  - quality
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a GitHub Actions CI pipeline that runs on PRs and pushes to main. The project already has 23 test files (67+ tests) and a `check-main-lines.sh` quality gate script with progressive line-count targets, but none of this runs automatically.

## Motivation
Without CI, regressions in tests or quality gate violations are only caught manually. As the refactoring effort (TASK-27.x) accelerates and new features land, automated checks become essential.

## Scope
1. **Test runner**: Run `pnpm test` on every PR and push to main
2. **Quality gates**: Run `check-main-lines.sh` to enforce main.ts line-count targets
3. **Type checking**: Run `tsc --noEmit` to catch type errors
4. **Build verification**: Run `make build` to confirm the app compiles
5. **Platform matrix**: Linux at minimum (primary target), macOS if feasible

## Implementation notes
- The project uses pnpm for package management
- Tests use Node's built-in test runner
- Build uses Make + tsc + electron-builder
- Consider caching node_modules and pnpm store for speed
- MeCab is a native dependency needed for some tests — document or skip if unavailable in CI
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 GitHub Actions workflow runs pnpm test on every PR and push to main.
- [ ] #2 Quality gate script (check-main-lines.sh) runs and fails the build if line count exceeds threshold.
- [ ] #3 tsc --noEmit type check passes as a CI step.
- [ ] #4 Build step (make build) completes without errors.
- [ ] #5 CI results are visible on PR checks.
- [ ] #6 Pipeline completes in under 5 minutes for typical changes.
<!-- AC:END -->
