---
id: TASK-85
title: 'Remove docs Plausible analytics integration'
status: Done
assignee: []
created_date: '2026-03-03 00:00'
updated_date: '2026-03-03 00:00'
labels: []
dependencies: []
priority: medium
ordinal: 12001
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Remove Plausible analytics integration from docs theme and dependency graph. Keep docs build/runtime analytics-free.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Docs theme no longer imports or initializes Plausible tracker.
- [x] #2 `@plausible-analytics/tracker` removed from dependencies and lockfile.
- [x] #3 Docs analytics test reflects absence of Plausible wiring.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Deleted Plausible runtime wiring from VitePress theme, removed tracker package via `bun remove`, and updated docs test to assert no Plausible integration remains.

<!-- SECTION:FINAL_SUMMARY:END -->
