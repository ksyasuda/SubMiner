---
id: TASK-333
title: Suppress aru subtitle annotations
status: Done
assignee: []
created_date: '2026-05-04 04:39'
updated_date: '2026-05-04 05:02'
labels:
  - tokenizer
  - annotations
  - bug
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add `ある` / `有る` to the subtitle annotation suppression path so `aru` tokens remain hoverable and never receive N+1, JLPT, frequency, or name-match annotation metadata. Known-word highlighting is special: if a filtered `aru` token is known and known highlighting is enabled, it should still render as known.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `ある` and kanji headword/surface variants such as `有る` are excluded by the subtitle annotation filter.
- [x] #2 Annotation stripping clears N+1, JLPT, frequency, and name metadata for `aru` tokens while preserving token hover data.
- [x] #3 Known-word highlighting still applies to filtered tokens, including `aru`, when known-word lookup marks them known.
- [x] #4 Regression coverage fails before the fix and passes after.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add `ある`/`有る`/`在る` to the shared subtitle annotation hard-exclusion terms.
2. Preserve/recompute known-word status for filtered tokens while stripping N+1, JLPT, frequency, and name metadata.
3. Add RED/GREEN unit and tokenizer regression coverage, plus a changelog fragment.
4. Run targeted tests and full handoff gate.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
TDD path: added failing annotation-stage coverage first. Initial implementation made targeted tests pass, then broader tokenizer coverage revealed an older fixture expecting `ある` to remain lexical; updated that integration expectation to the new requested behavior. Follow-up correction: known-word highlighting is the lone annotation exception for filtered tokens, so the strip path now preserves known state and `annotateTokens` recomputes known status for filtered tokens while still clearing N+1/JLPT/frequency/name metadata.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Suppressed non-known subtitle annotations for `aru` existence verbs by adding `ある`, `有る`, and `在る` to the shared hard-exclusion list. Corrected the filtered-token path so known-word highlighting still applies whenever known highlighting is enabled; filtered tokens now keep/gain `isKnown` but still lose N+1, JLPT, frequency, and name metadata.

Added and updated annotation-stage and tokenizer regression coverage for `aru`, particles, helper fragments, interjections, and other filtered known tokens. Added `changes/333-aru-annotation-filter.md`.

Validation passed: RED failures observed before implementation/correction; `bun test src/core/services/tokenizer/annotation-stage.test.ts`; `bun test src/core/services/tokenizer.test.ts`; `bun run typecheck`; `bun run format:check:src`; `bun run changelog:lint`; `bun run test:fast`; `bun run test:env`; `bun run build`; `bun run test:smoke:dist`.
<!-- SECTION:FINAL_SUMMARY:END -->
