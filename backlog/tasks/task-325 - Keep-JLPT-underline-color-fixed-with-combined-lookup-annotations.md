---
id: TASK-325
title: Keep JLPT underline color fixed with combined lookup annotations
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-04 00:25'
updated_date: '2026-05-04 00:28'
labels:
  - overlay
  - jlpt
  - renderer
dependencies: []
references:
  - TASK-318
  - TASK-308
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Yomitan lookup on a subtitle token that has a JLPT level plus another annotation such as frequency or known-word highlighting can make the JLPT underline take the other annotation color. The underline must always remain the token's JLPT level color; other annotation classes may still control text color.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 A JLPT token combined with frequency styling keeps its underline set to the configured JLPT level color during lookup/selection styling.
- [x] #2 A JLPT token combined with known-word styling keeps its underline set to the configured JLPT level color during lookup/selection styling.
- [x] #3 Regression coverage exercises combined JLPT plus non-JLPT annotation selectors, including character span selection/hover styling used by lookup.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused renderer CSS regression coverage for combined `word-jlpt-n*` plus known/frequency classes, including `.c::selection`/`.c:hover` lookup paths.
2. Run `bun test src/renderer/subtitle-render.test.ts` and confirm the new assertion fails on the current CSS.
3. Update `src/renderer/style.css` so JLPT decoration color is locked on the token and child character spans without changing text color priority for known/frequency/name/N+1 annotations.
4. Re-run the focused renderer test, then run typecheck/changelog checks as scope requires.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added red/green renderer CSS regression for combined JLPT plus known/N+1/frequency annotation classes and character hover lookup paths. Current CSS failed before the lock selectors were added; focused test passes after the CSS change.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed JLPT underline color drift for tokens that also carry known-word, N+1, or frequency annotation classes. The renderer CSS now explicitly locks the underline decoration color for combined JLPT annotation selectors, hover, character hover, and selection states while preserving the existing text color priority for other annotations.

Added renderer regression coverage for combined JLPT plus non-JLPT annotation selectors and lookup character hover paths. Added a user-visible changelog fragment.

Checks: `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`; `bun run typecheck`; `bun run format:check:src`.
<!-- SECTION:FINAL_SUMMARY:END -->
