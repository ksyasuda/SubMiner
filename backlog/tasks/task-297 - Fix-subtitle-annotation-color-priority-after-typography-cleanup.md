---
id: TASK-297
title: Fix subtitle annotation color priority after typography cleanup
status: Done
assignee: []
created_date: '2026-04-25 23:44'
updated_date: '2026-04-25 23:46'
labels:
  - subtitles
  - renderer
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Known-word and frequency subtitle token colors should keep their configured priority after annotation CSS stopped using JLPT underlines.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Known-word token color takes priority over JLPT and frequency color classes.
- [x] #2 Frequency single-mode token color takes priority over JLPT color classes when frequency annotation is active.
- [x] #3 Regression coverage verifies CSS selectors do not allow JLPT color rules to override higher-priority annotation colors.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Scoped JLPT token color selectors so they only apply when no higher-priority known-word, N+1, name-match, or frequency class is present. This keeps known words green and frequency single-mode tokens using the single frequency color instead of being visually overridden by JLPT colors. Added CSS regression assertions for this priority behavior. Verified with `bun test src/renderer/subtitle-render.test.ts`, `bun run typecheck`, and `bun run test:fast`.
<!-- SECTION:FINAL_SUMMARY:END -->
