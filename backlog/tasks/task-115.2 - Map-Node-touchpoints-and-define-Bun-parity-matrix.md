---
id: TASK-115.2
title: Map Node touchpoints and define Bun parity matrix
status: Done
assignee: []
created_date: '2026-02-23 04:27'
updated_date: '2026-02-23 04:36'
labels:
  - tooling
  - planning
dependencies: []
references:
  - package.json
  - .github/workflows/ci.yml
  - .github/workflows/release.yml
documentation:
  - docs/plans/2026-02-23-bun-only-toolchain-migration.md
parent_task_id: TASK-115
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Document all Node-specific command and workflow touchpoints, define Bun replacements, and establish migration guardrails so follow-up implementation tasks have clear scope and risk controls.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A complete inventory of Node-dependent scripts and workflow steps is documented with proposed Bun equivalents.
- [x] #2 Migration risks and compatibility checkpoints are explicitly documented for dist and Electron-related lanes.
- [x] #3 Follow-up implementation tasks have clear boundaries and ordering based on the parity matrix.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Created `docs/plans/bun-only-parity-matrix.md` with direct Node touchpoint inventory, Bun replacements, risk checkpoints, and migration sequencing.

Parity matrix now links implementation ordering for TASK-115.3/115.4/115.1 and flags follow-up risk handling for third-party implicit Node assumptions.
<!-- SECTION:NOTES:END -->
