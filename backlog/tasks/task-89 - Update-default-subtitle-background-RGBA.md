---
id: TASK-89
title: Update default subtitle background RGBA
status: Done
assignee: []
created_date: '2026-02-20 05:43'
updated_date: '2026-02-20 05:44'
labels: []
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Set default subtitle background color to rgb(30, 32, 48, 0.88) as requested.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Default subtitle background color value matches requested string.
- [ ] #2 Config tests pass.
- [ ] #3 Default subtitle background color value matches requested string.
- [ ] #4 Config tests pass.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Updated default subtitle background to rgb(30, 32, 48, 0.88) in config defaults, renderer fallback CSS, and docs/examples. Added regression assertion in src/config/config.test.ts. Verified with bun test src/config/config.test.ts (33/33 pass).
<!-- SECTION:NOTES:END -->
