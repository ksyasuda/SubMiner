---
id: TASK-285
title: Rename anime visibility filter heading to title visibility
status: Done
assignee:
  - codex
created_date: '2026-04-10 00:00'
updated_date: '2026-04-10 00:00'
labels:
  - stats
  - ui
  - bug
milestone: m-1
dependencies: []
references:
  - stats/src/components/trends/TrendsTab.tsx
  - stats/src/components/trends/TrendsTab.test.tsx
priority: low
ordinal: 200000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Align the library cumulative trends filter UI with the new terminology by renaming the hardcoded anime visibility heading to title visibility.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 The trends filter heading uses `Title Visibility`
- [x] #2 The component behavior and props stay unchanged
- [x] #3 A regression test covers the rendered heading text
<!-- AC:END -->
