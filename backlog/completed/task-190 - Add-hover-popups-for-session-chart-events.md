---
id: TASK-190
title: Add hover popups for session chart events
status: Done
assignee:
  - Codex
created_date: '2026-03-17 22:20'
updated_date: '2026-03-18 05:28'
labels:
  - stats
  - ui
  - bug
milestone: m-1
dependencies: []
references:
  - stats/src/components/sessions/SessionDetail.tsx
  - stats/src/lib/session-events.ts
  - stats/src/hooks/useSessions.ts
  - stats/src/lib/api-client.ts
  - docs/plans/2026-03-17-session-event-hover-popups-design.md
priority: medium
ordinal: 105500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add hover/focus popups to session chart event markers so pauses, seeks, lookups, and card-mine events explain themselves inline. Card-mine events should lazy-load available Anki note info and present it in a richer popup with browse affordances.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Hovering or focusing a session-chart marker opens an event-specific popup.
- [x] #2 Pause, seek, and lookup popups show concise event copy derived from marker metadata.
- [x] #3 Card-mine popups lazily fetch and cache Anki note info by note id.
- [x] #4 Card-mine popups show a formatted fallback when note info is missing or still loading.
- [x] #5 Regression tests cover event payload shaping and popup rendering behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for event metadata shaping and popup content selection.
2. Extend session-event shaping to parse payload JSON into typed marker metadata.
3. Add lazy note-info fetch/cache state for card-mine markers.
4. Render interactive marker overlay + custom popup in the session detail chart.
5. Run targeted stats/core verification and update this task with the result.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->
Completed. Session-chart event markers now open event-specific hover/focus popups, including lazy-loaded Anki note info for card-mine events with browse affordances. Verification passed via targeted stats tests, `bun run typecheck`, and the core verification lane in `.tmp/skill-verification/subminer-verify-20260317-222545-CQzyqK`.
<!-- SECTION:OUTCOME:END -->
