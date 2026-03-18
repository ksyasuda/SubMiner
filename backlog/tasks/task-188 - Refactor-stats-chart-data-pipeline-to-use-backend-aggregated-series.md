---
id: TASK-188
title: Refactor stats chart data pipeline to use backend-aggregated series
status: In Progress
assignee:
  - codex
created_date: '2026-03-18 00:29'
updated_date: '2026-03-18 00:55'
labels:
  - stats
  - performance
  - refactor
dependencies: []
references:
  - src/core/services/immersion-tracker/query.ts
  - src/core/services/immersion-tracker-service.ts
  - src/core/services/stats-server.ts
  - stats/src/hooks/useTrends.ts
  - stats/src/components/trends/TrendsTab.tsx
  - stats/src/lib/api-client.ts
  - stats/src/types/stats.ts
  - stats/src/lib/dashboard-data.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce long-term dashboard performance debt by moving chart aggregation out of the stats UI and into the tracker/stats API layer. The trends dashboard should consume chart-ready series from backend rollups instead of reconstructing multiple datasets from raw session lists in the browser.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Stats API exposes chart-oriented aggregated trend data needed by the trends dashboard without requiring raw session lists for those charts.
- [ ] #2 The trends dashboard consumes the new aggregated API responses and no longer rebuilds its main chart datasets from raw sessions in the render path.
- [ ] #3 Time-range and grouping behavior remain correct for recent and all-time views, with explicit handling that keeps older history performant.
- [ ] #4 Existing overview and anime detail charts continue to behave correctly, or are migrated to the shared aggregation path where it reduces debt.
- [ ] #5 Tests cover backend aggregation/query behavior and frontend consumption of the new response shapes.
- [ ] #6 Internal docs are updated to describe the new stats chart data flow and scaling rationale.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a chart-oriented trends dashboard API response on the stats server that returns pre-aggregated series by range/grouping instead of requiring raw session lists in the UI.
2. Implement tracker/query-layer helpers that aggregate trend series on the backend, preferring rollups for scalable time-series data and centralizing chart shaping there.
3. Update stats client types and `useTrends` to consume the new response shape and stop fetching raw sessions for main chart construction.
4. Simplify `TrendsTab` and related chart components so they render backend-provided series with only lightweight UI-level filtering/state.
5. Keep overview/anime detail chart behavior intact, and reuse shared aggregation paths where it meaningfully reduces debt without widening scope.
6. Add/adjust backend and frontend tests plus internal docs to describe the new chart-data flow and performance rationale.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented a new `/api/stats/trends/dashboard` server route backed by tracker/query-layer aggregation, updated the stats client and `useTrends` to consume the new chart-ready payload, simplified `TrendsTab` to render backend-provided series, added route/query/api-client tests, and documented the new trends data flow in `docs/architecture/stats-trends-data-flow.md`.

Did not run validation commands in this pass; acceptance criteria remain unchecked pending requested verification.
<!-- SECTION:NOTES:END -->
