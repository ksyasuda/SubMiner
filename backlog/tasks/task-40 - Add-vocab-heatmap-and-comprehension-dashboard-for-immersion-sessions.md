---
id: TASK-40
title: Add vocab heatmap and comprehension dashboard for immersion sessions
status: To Do
assignee: []
created_date: '2026-02-14 02:06'
labels:
  - feature
  - analytics
  - immersion
  - visualization
dependencies:
  - TASK-28
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Build a visual comprehension dashboard that renders per-video "comprehension heatmaps" showing which segments had high lookup density versus segments where the user understood everything without lookups.

## Motivation

Immersion tracking data (TASK-28) captures lookup counts, words seen, and cards mined per time segment. Visualizing this data gives learners a powerful way to see their progress: early episodes of a show will be "hot" (many lookups), while later episodes should cool down as vocabulary grows.

## Features

1. **Per-video heatmap**: A timeline bar colored by lookup density (green = understood, yellow = some lookups, red = many lookups)
2. **Cross-video trends**: Show comprehension improvement across episodes of the same series
3. **Session summary cards**: After each session, show a summary with key metrics (time watched, words seen, cards mined, comprehension estimate)
4. **Historical comparison**: Compare the same video rewatched at different dates
5. **Export**: Generate a shareable image or markdown summary

## Data sources

- `imm_session_telemetry` (lookup_count, words_seen per sample interval)
- `imm_session_events` (individual lookup events with timestamps)
- `imm_daily_rollups` / `imm_monthly_rollups` for trend data

## Implementation notes

- Could be rendered in a dedicated Electron window or an overlay modal
- Use canvas or SVG for the heatmap visualization
- Consider a simple HTML report that can be opened in a browser for sharing
- Heatmap granularity should match telemetry sample interval
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [ ] #1 Per-video heatmap renders a timeline colored by lookup density.
- [ ] #2 Cross-video trend view shows comprehension change across episodes.
- [ ] #3 Session summary displays time watched, words seen, cards mined, and comprehension estimate.
- [ ] #4 Dashboard data is sourced from TASK-28 immersion tracking tables.
- [ ] #5 Visualization is exportable as an image or shareable format.
- [ ] #6 Dashboard does not block or slow down active mining sessions.
<!-- AC:END -->
