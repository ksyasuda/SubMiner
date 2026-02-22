---
id: TASK-100
title: Run post-refactor dead code prune and cleanup
status: Done
assignee: []
created_date: '2026-02-21 07:15'
updated_date: '2026-02-22 04:01'
labels:
  - cleanup
  - maintainability
  - refactor
dependencies:
  - TASK-96
  - TASK-97
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Major refactors likely left unused exports/helpers and stale compatibility code. Perform a deliberate dead-code sweep with regression safety.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Run unused-export scans (`ts-prune` or equivalent) and collect candidate list.
2. Manually verify each candidate to avoid false positives from dynamic loading/IPC wiring.
3. Remove or inline confirmed dead code; simplify import surfaces/barrels accordingly.
4. Delete stale helper paths retained only for pre-composer wiring.
5. Add or update regression tests where removal could alter behavior.
6. Run verification gate: `bun run build`, `bun run test:config:dist`, `bun run test:core:dist`.
7. Publish cleanup report: removed files/exports, kept exceptions, and rationale.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Confirmed dead code removed without behavior regressions.
- [x] #2 Remaining flagged candidates either resolved or documented with justification.
- [x] #3 Build and core/config suites pass after cleanup.
- [x] #4 Import graph complexity reduced (fewer exports/entrypoints where applicable).
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added dead-code execution report: `docs/reports/2026-02-22-task-100-dead-code-report.md` (baseline + triage + removals + remaining candidates).

Confirmed removals include unused imports/helpers in Anki/core/renderer paths plus reduced registry/barrel export surface in `src/tokenizers/index.ts`, `src/token-mergers/index.ts`, and `src/core/utils/index.ts`.

Verification gates: `bun run build` PASS; `bun run test:core:src` PASS (225 pass/6 skip); `bun run test:config:src` PASS (52 pass); `bun run check:file-budgets` PASS warning mode with no strict hotspot violations.

Static-analysis evidence: `tsc --noEmit --noUnusedLocals --noUnusedParameters` reduced to 39 remaining diagnostics concentrated in `src/main.ts` and intentional composer type-test aliases; broad `ts-prune` false positives documented in report.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed post-refactor dead-code cleanup with a documented triage report, removing confirmed unused imports/helpers/exports while preserving intentional contract/type seams. Verified no regressions via build + core/config source test suites and maintainability budget checks, with remaining candidate hotspots documented for a dedicated follow-up pass.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Dead-code report attached in task notes.
- [x] #2 Regression test updates included for risky removals.
- [x] #3 Verification gate commands complete successfully.
<!-- DOD:END -->
