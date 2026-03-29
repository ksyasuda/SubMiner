---
id: TASK-242
title: Fix stats server Bun fallback in coverage lane
status: In Progress
assignee: []
created_date: '2026-03-29 07:31'
labels:
  - ci
  - bug
milestone: cleanup
dependencies: []
references:
  - 'PR #36'
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Coverage CI fails when `startStatsServer` reaches the Bun server seam under the maintained source lane. Add a runtime fallback that works when `Bun.serve` is unavailable and keep the stats-server startup path testable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `bun run test:coverage:src` passes in GitHub CI
- [ ] #2 `startStatsServer` uses `Bun.serve` when present and a Node server fallback otherwise
- [ ] #3 Regression coverage exists for the fallback startup path
<!-- AC:END -->
