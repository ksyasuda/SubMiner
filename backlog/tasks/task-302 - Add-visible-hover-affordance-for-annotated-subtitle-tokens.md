---
id: TASK-302
title: Add visible hover affordance for annotated subtitle tokens
status: Done
assignee: []
created_date: '2026-04-26 03:39'
updated_date: '2026-04-26 03:40'
labels:
  - overlay
  - subtitle
  - ux
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Annotated subtitle tokens keep annotation colors on hover, but with transparent hover backgrounds there is no visible hover indication. Add a subtle affordance that preserves annotation color semantics.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Annotated token hover keeps annotation color instead of switching to hover text color.
- [x] #2 Annotated token hover has a visible indication when hover background is transparent.
- [x] #3 Regression tests cover the hover CSS contract.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing CSS contract test for annotated token hover affordance.
2. Add brightness/saturation filter to annotated token hover CSS without changing annotation color.
3. Add changelog fragment and run focused verification.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented option 1 from approved design: annotated subtitle word hover keeps annotation color and adds `filter: brightness(1.18) saturate(1.08)` for visible affordance when the hover background is transparent.

Verification passed: `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Added a visible hover affordance for annotated subtitle tokens without changing annotation color semantics. Annotated word hover now applies a small brightness/saturation lift while retaining the existing background behavior, so transparent hover backgrounds still show feedback.

Regression coverage updated in `src/renderer/subtitle-render.test.ts` to assert the hover filter is present and that hover color overrides are still absent for annotated tokens. Added changelog fragment `changes/302-annotated-hover-affordance.md`.

Verification: `bun test src/renderer/subtitle-render.test.ts`; `bun run changelog:lint`.
<!-- SECTION:FINAL_SUMMARY:END -->
