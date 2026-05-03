---
id: TASK-318
title: Keep JLPT underline color fixed after lookup selection
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-03 03:17'
updated_date: '2026-05-03 03:19'
labels:
  - overlay
  - jlpt
  - renderer
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Looking up a subtitle token can leave browser/Yomitan selection styling active. If that token has a JLPT class and another annotation class, the underline must remain the JLPT level color because underline color represents static JLPT classification, not the currently active annotation or lookup state.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 JLPT subtitle underlines retain their configured N1-N5 color after lookup/selection styling is applied.
- [x] #2 JLPT tokens that also have known, N+1, name, or frequency annotation classes keep their annotation text color behavior without changing the JLPT underline color.
- [x] #3 Renderer regression coverage verifies the CSS contract for the combined JLPT plus annotation case.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused CSS regression in `src/renderer/subtitle-render.test.ts` for JLPT tokens combined with higher-priority annotation classes and lookup/selection styling.
2. Run the focused renderer test and confirm it fails because selection rules do not lock `text-decoration-color`.
3. Update `src/renderer/style.css` to explicitly preserve JLPT underline decoration color in lookup/selection state selectors without changing text color priority.
4. Re-run the focused renderer test, then run the smallest relevant verification gate.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verified TDD red/green for renderer CSS contract: `bun test src/renderer/subtitle-render.test.ts` first failed because `word-jlpt-n1::selection` lock was missing, then passed after adding explicit JLPT `text-decoration-color` selection rules. Also ran `bun run changelog:lint` and `bun run typecheck` successfully.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed JLPT subtitle underline color drift after dictionary lookup/selection by adding explicit `::selection` decoration-color locks for N1-N5 token classes in `src/renderer/style.css`. This preserves the JLPT underline as static classification while leaving known/N+1/name/frequency text color priority intact.

Added renderer CSS regression coverage for the JLPT selection lock and a user-visible changelog fragment.

Checks: `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`; `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
