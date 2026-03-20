---
id: TASK-193
title: Fix session chart event popup position drift
status: Done
assignee:
  - Codex
created_date: '2026-03-17 23:55'
updated_date: '2026-03-17 23:59'
labels:
  - stats
  - ui
  - bug
milestone: m-1
dependencies: []
references:
  - stats/src/components/sessions/SessionDetail.tsx
  - stats/src/components/sessions/SessionEventOverlay.tsx
  - stats/src/lib/session-events.ts
priority: medium
ordinal: 105600
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Fix the session timeline event popup trigger positions so hover markers stay aligned with the underlying chart event lines across the full visible time range.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Event popup triggers stay horizontally aligned with chart event lines from session start through session end.
- [x] #2 Alignment logic uses the rendered chart plot area rather than guessed container percentages.
- [x] #3 Regression coverage locks the marker-position projection math.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a failing regression test for marker-position projection with chart offsets.
2. Capture the rendered plot box from Recharts and pass it into the overlay.
3. Position overlay markers in plot-area pixels, rerun targeted stats verification, then record the result.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->

Completed. Session event hover markers now read the actual Recharts plot-area offset and width, then project marker X positions into plot-area pixels instead of full-container percentages. That keeps popup triggers aligned with the underlying reference lines across long session timelines.

Verification:

- `bun test stats/src/lib/session-events.test.ts stats/src/lib/session-detail.test.tsx stats/src/components/sessions/SessionEventPopover.test.tsx`
- `cd stats && bun run build`
- `bun x prettier --check 'stats/src/components/sessions/SessionDetail.tsx' 'stats/src/components/sessions/SessionEventOverlay.tsx' 'stats/src/lib/session-events.ts' 'stats/src/lib/session-events.test.ts' 'backlog/tasks/task-193 - Fix-session-chart-event-popup-position-drift.md'`
- `bun run typecheck:stats` still fails on pre-existing unrelated errors in `src/components/anime/AnilistSelector.tsx`, `src/components/library/LibraryTab.tsx`, `src/lib/reading-utils.test.ts`, `src/lib/reading-utils.ts`, `src/lib/vocabulary-tab.test.ts`, and `src/lib/yomitan-lookup.test.tsx`

<!-- SECTION:OUTCOME:END -->
