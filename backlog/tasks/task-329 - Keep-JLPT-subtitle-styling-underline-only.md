---
id: TASK-329
title: Keep JLPT subtitle styling underline-only
status: Done
assignee: []
created_date: '2026-05-04 02:13'
labels:
  - bug
  - renderer
  - jlpt
dependencies: []
references:
  - src/renderer/style.css
  - src/renderer/subtitle-render.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix subtitle token styling so JLPT metadata never changes token text color. JLPT should only render the level marker/underline affordance while known, n+1, name-match, and frequency colors retain priority.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 JLPT-only subtitle tokens do not set token text color.
- [x] #2 JLPT level marker/underline still uses configured JLPT color.
- [x] #3 Existing known, n+1, name-match, and frequency text colors remain unchanged.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Changed subtitle JLPT styling from text color to underline decoration and updated renderer CSS regression coverage.

Verification:
- `bun test src/renderer/subtitle-render.test.ts`
- `bunx prettier --check src/renderer/subtitle-render.test.ts src/renderer/style.css`
- `bun run typecheck`

Blocked:
- `bun run test:fast` fails in existing dirty stats/session work: `recordSubtitleLine counts exact Yomitan tokens for session metrics` expects `tokensSeen` 4 but gets 3.
<!-- SECTION:FINAL_SUMMARY:END -->
