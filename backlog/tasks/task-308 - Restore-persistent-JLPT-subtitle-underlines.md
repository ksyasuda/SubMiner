---
id: TASK-308
title: Restore persistent JLPT subtitle underlines
status: Done
assignee:
  - Codex
created_date: '2026-04-27 02:03'
updated_date: '2026-04-27 02:07'
labels:
  - overlay
  - jlpt
  - renderer
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
JLPT tagging currently exposes the JLPT level on hover, but the persistent subtitle underline is missing. When JLPT annotation is enabled and a rendered subtitle token has a JLPT level, users should see the configured JLPT color underline without needing to hover.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 JLPT-tagged subtitle tokens render a persistent underline for N1-N5 levels when JLPT tagging is enabled.
- [x] #2 Hover and keyboard-selected JLPT labels continue to appear for tagged tokens.
- [x] #3 Higher-priority annotation colors such as known words, N+1, names, and frequency styling are not overridden by JLPT text color.
- [x] #4 Regression coverage verifies the CSS contract for persistent JLPT underlines.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused renderer CSS regression asserting each `word-jlpt-n*` class provides persistent underline decoration while preserving existing typography constraints.
2. Run the focused renderer test to confirm the regression fails before production changes.
3. Restore underline CSS for JLPT classes without broadening JLPT text-color precedence over known/N+1/name/frequency tokens.
4. Re-run the focused renderer test and update acceptance criteria/task notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Verified red/green regression: tightened `src/renderer/subtitle-render.test.ts` first failed because base `word-jlpt-n*` selectors had no underline decoration, then passed after moving JLPT underline decoration to unconditional base selectors while leaving JLPT text color priority-scoped.

Checks: `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`; `bun run typecheck`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored persistent JLPT subtitle underlines by adding underline decoration to each base `word-jlpt-n*` renderer CSS class. JLPT text color remains in the existing priority-scoped selectors, so known/N+1/name/frequency coloring is not overridden while the underline still appears on any JLPT-tagged token.

Updated renderer CSS regression coverage to assert underline decoration for N1-N5 and added a fixed changelog fragment. Verified with `bun test src/renderer/subtitle-render.test.ts`, `bun run changelog:lint`, and `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
