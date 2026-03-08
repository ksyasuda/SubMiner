---
id: TASK-114
title: Fix failing CI checks on PR 15
status: Done
assignee:
  - codex
created_date: '2026-03-08 00:34'
updated_date: '2026-03-08 00:37'
labels:
  - ci
  - test
dependencies: []
references:
  - src/renderer/subtitle-render.test.ts
  - src/renderer/style.css
  - .github/workflows/ci.yml
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate the failing GitHub Actions CI run for PR #15 on branch `yomitan-fork`, fix the underlying test or code regression, and verify the affected local test/CI lane passes.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Identified the concrete failing CI job and captured the relevant failure context
- [x] #2 Implemented the minimal code or test change needed to resolve the CI failure
- [x] #3 Verified the affected local test target and the broader fast CI test lane pass
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect the failing GitHub Actions run and confirm the exact failing test/assertion.
2. Reproduce the failing renderer stylesheet test locally and compare the assertion against current CSS.
3. Apply the minimal test or stylesheet fix needed to restore the intended hover/selection behavior.
4. Re-run the targeted renderer test, then re-run `bun run test` to verify the fast CI lane is green.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
GitHub Actions run 22810400921 failed in job build-test-audit, step `Test suite (source)`, with a single failing test: `JLPT CSS rules use underline-only styling in renderer stylesheet` in src/renderer/subtitle-render.test.ts.

Reproduced the failing test locally with `bun test src/renderer/subtitle-render.test.ts`. The failure was a brittle stylesheet assertion, not a renderer behavior regression.

Updated the renderer stylesheet test helper to split selectors safely across `:is(...)` commas and normalize multiline selector whitespace, then switched the failing hover/JLPT assertions to inspect extracted rule blocks instead of matching the entire CSS file text.

Verification passed with `bun test src/renderer/subtitle-render.test.ts` and `bun run test`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Investigated GitHub Actions CI run `22810400921` for PR #15 and confirmed the only failing job was `build-test-audit`, step `Test suite (source)`, with a single failure in `src/renderer/subtitle-render.test.ts` (`JLPT CSS rules use underline-only styling in renderer stylesheet`).

The renderer CSS itself was still correct; the regression was in the test helper. `extractClassBlock` was splitting selector lists on every comma, which breaks selectors containing `:is(...)`, and the affected assertions fell back to brittle whole-file regex matching against a multiline selector. Fixed the test by teaching the helper to split selectors only at top-level commas, normalizing selector whitespace around multiline `:not(...)` / `:is(...)` clauses, and asserting on extracted rule blocks for the plain-word hover and JLPT-only hover/selection rules.

Verification: `bun test src/renderer/subtitle-render.test.ts` passed, and `bun run test` passed end to end (the same fast lane that failed in CI).
<!-- SECTION:FINAL_SUMMARY:END -->
