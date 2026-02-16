---
id: TASK-22
title: Make secondary subtitles hover-revealed but non-lookupable in Yomitan sessions
status: To Do
assignee: []
created_date: '2026-02-13 16:40'
labels: []
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate and implement a UX where secondary subtitles (e.g., English text in our current sessions) become visible when hovered, while explicitly preventing user interactions that allow text lookup (including Yomitan integration) on those subtitles. This should allow readability-on-hover without exposing the secondary overlay text to selection/lookup workflows.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Secondary subtitles become visible only while hovered (or via equivalent hover-triggered mechanism), and return to default hidden/low-visibility state when not hovered.
- [ ] #2 When hovered, secondary subtitles do not trigger Yomitan lookup behavior in sessions where Yomitan is enabled.
- [ ] #3 Secondary subtitles remain non-interactive for lookup paths (for example, text selection or lookup event propagation) while hover-visibility still works as intended.
- [ ] #4 Primary subtitles remain functional and are not regressed by the secondary-subtitle interaction changes.
- [ ] #5 If complete prevention of Yomitan lookup on secondary subtitles is not technically possible, the task includes the known limitations and a documented fallback behavior.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Acceptance criteria are reviewed and covered by explicit manual/automated test coverage for hover reveal and lookup suppression behavior.
<!-- DOD:END -->
