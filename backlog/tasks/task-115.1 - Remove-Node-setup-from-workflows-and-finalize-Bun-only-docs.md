---
id: TASK-115.1
title: Remove Node setup from workflows and finalize Bun-only docs
status: Done
assignee: []
created_date: '2026-02-23 04:26'
updated_date: '2026-02-23 04:36'
labels:
  - ci
  - docs
  - tooling
dependencies:
  - TASK-115.2
  - TASK-115.3
  - TASK-115.4
references:
  - .github/workflows/ci.yml
  - .github/workflows/release.yml
  - docs/development.md
  - docs/installation.md
  - README.md
documentation:
  - docs/plans/2026-02-23-bun-only-toolchain-migration.md
parent_task_id: TASK-115
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After script/test migration parity is verified, remove Node setup from CI/release workflows and finalize documentation so contributors and automation use a consistent Bun-only prerequisite.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 CI and release workflows no longer include Node setup steps for routine build/test/package jobs.
- [x] #2 Primary local verification gates pass with the Bun-only command set.
- [x] #3 README and setup/install docs consistently describe Bun as the JS runtime prerequisite without conflicting Node requirement text.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Removed `actions/setup-node` from `.github/workflows/ci.yml` and `.github/workflows/release.yml` (quality-gate, build-linux, build-macos jobs).

Updated docs to align Bun-only prerequisite: removed Node from `docs/development.md` prerequisites and switched stale pnpm source-build snippet in `docs/installation.md` to Bun.

Validation gates: `bun run test:launcher:smoke:src && bun run test:fast && bun run build && bun run test:config:dist && bun run test:core:dist && bun run test:smoke:dist && bun run docs:build` all pass.
<!-- SECTION:NOTES:END -->
