---
id: TASK-23.4
title: Add opt-in control and end-to-end flow + tests for JLPT tagging
status: To Do
assignee: []
created_date: '2026-02-13 16:42'
labels: []
dependencies: []
parent_task_id: TASK-23
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add user/config setting to enable JLPT tagging, wire the feature toggle through subtitle processing/rendering, and add tests/verification for positive match, non-match, and disabled-mode behavior.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 JLPT tagging is opt-in and defaults to disabled.
- [ ] #2 When disabled, lookup/rendering pipeline does not execute JLPT processing.
- [ ] #3 When enabled, end-to-end flow tags subtitle words via token-level lookup and rendering.
- [ ] #4 Add tests covering at least one positive match, one non-match, and disabled state.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 End-to-end option behavior and opt-in state persistence are implemented and verified.
<!-- DOD:END -->
