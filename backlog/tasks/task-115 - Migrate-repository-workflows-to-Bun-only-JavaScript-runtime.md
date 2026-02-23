---
id: TASK-115
title: Migrate repository workflows to Bun-only JavaScript runtime
status: Done
assignee: []
created_date: '2026-02-23 04:26'
updated_date: '2026-02-23 04:42'
labels:
  - tooling
  - build
  - ci
dependencies: []
references:
  - package.json
  - .github/workflows/ci.yml
  - .github/workflows/release.yml
  - docs/development.md
  - docs/installation.md
  - README.md
documentation:
  - docs/plans/2026-02-23-bun-only-toolchain-migration.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Transition the project from mixed Bun/Node tooling to a Bun-only contributor and CI workflow so setup is simpler and runtime behavior is consistent across local and automation lanes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A contributor can install dependencies, build, run source tests, run dist smoke tests, and build docs using Bun without requiring a system Node.js install.
- [x] #2 CI and release workflows no longer require Node setup steps for standard project verification and packaging jobs.
- [x] #3 Dist and utility verification lanes remain covered and pass with Bun-based commands.
- [x] #4 Repository documentation clearly states the Bun-only prerequisite and removes conflicting Node requirement guidance.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Execution started in-session after user selected Subagent-Driven option. Tracking sequence: TASK-115.2 -> TASK-115.3/TASK-115.4 -> TASK-115.1.

Completed child sequence TASK-115.2 -> TASK-115.3/TASK-115.4 -> TASK-115.1 with Bun-only command/workflow migration.

Direct Node command usage removed from package dist and utility script lanes; CI/release Node setup steps removed; contributor docs aligned to Bun-only prerequisites.

Regression hardening: refactored AniList token-store tests and storage seam to remain stable under Bun dist test execution while preserving runtime behavior.

Post-migration docs polish: updated `README.md` requirements table to list Bun as a required dependency and promoted Bun from optional-tools section into system dependencies in `docs/installation.md` to remove ambiguity.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Migrated repository verification and workflow surfaces to Bun-first execution by replacing `node --test` dist lanes and remaining direct Node utility invocation (`generate:config-example`) with Bun commands.

Removed `actions/setup-node` from CI and release jobs and aligned docs (`docs/development.md`, `docs/installation.md`) to Bun-only setup guidance; launcher smoke fixtures now use Bun shebangs so smoke tests no longer depend on Node.

Validated full gate set under Bun-only command path: launcher smoke, fast source suite, build, dist config/core/smoke suites, docs build, plus config-example generation command.
<!-- SECTION:FINAL_SUMMARY:END -->
