---
id: TASK-242
title: Fix stats server Bun fallback in coverage lane
status: Done
assignee: []
created_date: '2026-03-29 07:31'
updated_date: '2026-03-31 19:37'
labels:
  - ci
  - bug
milestone: cleanup
dependencies: []
references:
  - 'PR #36'
priority: high
ordinal: 170500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Coverage CI fails when `startStatsServer` reaches the Bun server seam under the maintained source lane. Add a runtime fallback that works when `Bun.serve` is unavailable and keep the stats-server startup path testable.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `bun run test:coverage:src` passes in GitHub CI
- [x] #2 `startStatsServer` uses `Bun.serve` when present and a Node server fallback otherwise
- [x] #3 Regression coverage exists for the fallback startup path
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the CI failure in the coverage lane by replacing the Bun-only stats server path with a Bun-or-node/http startup fallback and by normalizing setup window options so undefined BrowserWindow fields are omitted. Verified the exact coverage lane under Bun 1.3.5 and confirmed the GitHub Actions run for PR #36 completed successfully.
<!-- SECTION:FINAL_SUMMARY:END -->
