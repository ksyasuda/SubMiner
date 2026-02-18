---
id: TASK-23.3
title: Render JLPT token underlines with level-based colors in subtitle lines
status: Done
assignee: []
created_date: '2026-02-13 16:42'
updated_date: '2026-02-18 04:11'
labels: []
dependencies: []
parent_task_id: TASK-23
priority: high
ordinal: 44000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Render JLPT-aware token annotations as token-length colored underlines in the subtitle UI based on returned JLPT levels, without changing existing subtitle layout or primary interaction behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 For each token with JLPT level, renderer draws an underline matching token width/length.
- [x] #2 Underlines use distinct colors by JLPT level (e.g., N5/N4/N3/N2/N1) and mapping is consistent/documented.
- [x] #3 Non-tagged tokens remain visually unchanged.
- [x] #4 Rendering does not alter line height/selection behavior or break wrapping behavior.
- [x] #5 Feature degrades gracefully when level data is missing or lookup is unavailable.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Visual output validated for all mapped JLPT levels with no legibility/layout regressions.
<!-- DOD:END -->
