---
id: TASK-96
title: Split config resolve into domain modules
status: Done
assignee:
  - '@codex-task96-config-resolve'
created_date: '2026-02-21 07:15'
updated_date: '2026-02-22 07:49'
labels:
  - architecture
  - refactor
  - maintainability
dependencies:
  - TASK-85
priority: high
ordinal: 86000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/config/resolve.ts` remains oversized and mixes unrelated concerns. Split into domain-focused modules while preserving behavior and public API.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Capture baseline: LOC, exported symbols, and current test coverage touching config resolution.
2. Create domain modules under `src/config/resolve/`:
   - `env-paths.ts`
   - `subtitle-style.ts`
   - `integrations.ts`
   - `validation-errors.ts`
3. Keep `src/config/resolve.ts` as thin orchestrator/barrel that composes domain resolvers.
4. Add seam tests per new module; keep existing config suite green.
5. Verify no call-site changes required outside config service layer.
6. Run verification gate: `bun run build`, `bun run test:config:dist`, `bun run check:file-budgets`.
7. Record before/after LOC and ownership map in task notes.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `src/config/resolve.ts` reduced to orchestration/barrel role.
- [x] #2 New domain modules exist with clear single-responsibility boundaries.
- [x] #3 Existing config behavior preserved with passing tests.
- [x] #4 New seam tests cover domain-specific resolution logic.
- [x] #5 File-budget report shows measurable reduction in hotspot file size.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Execution plan (2026-02-21 codex refresh):
1. Capture baseline metrics (`wc -l src/config/resolve.ts`, `bun run check:file-budgets`) and keep evidence in notes.
2. Ensure seam test coverage exists for extracted domain modules (`src/config/resolve/anki-connect.test.ts`, `src/config/resolve/subtitle-style.test.ts`, `src/config/resolve/jellyfin.test.ts`) and wire config test scripts so these tests run in src + dist lanes.
3. Reduce `src/config/resolve.ts` to orchestration facade by composing extracted modules in stable order:
   - `createResolveContext`
   - `applyTopLevelConfig`
   - `applyCoreDomainConfig`
   - `applySubtitleDomainConfig`
   - `applyIntegrationConfig`
   - `applyImmersionTrackingConfig`
   - `applyAnkiConnectResolution`
4. Verify no behavior drift with required gates: `bun run build`, `bun run test:config:dist`, `bun run check:file-budgets`.
5. Record before/after LOC + budget results and finalize acceptance criteria / DoD in task metadata (no commit).
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
[baseline] `wc -l src/config/resolve.ts` => 1414 LOC.

[baseline] `bun run check:file-budgets` reports `src/config/resolve.ts: 1415 LOC` and 17 files over 500 LOC.

[implementation] Reduced `src/config/resolve.ts` to orchestration facade that now composes extracted domain modules in stable order: top-level -> core -> subtitle -> integrations -> immersion-tracking -> anki-connect.

[implementation] Wired config seam tests into official config lanes by updating `package.json` scripts `test:config:src` and `test:config:dist` to include `src/config/resolve/{anki-connect,subtitle-style,jellyfin}.test.ts` (and compiled dist equivalents).

[verification] `bun run build && bun run test:config:dist && bun run check:file-budgets` passed. Config dist lane now runs 48 tests including seam tests; all passing.

[metrics] LOC before/after: `wc -l src/config/resolve.ts` baseline 1414 -> final 33.

[metrics] Budget before/after: baseline report showed 18 files over 500 LOC including `src/config/resolve.ts: 1415 LOC`; final report shows 17 files over 500 LOC and `src/config/resolve.ts` no longer over budget.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Refactored config resolution into domain-focused modules and reduced `src/config/resolve.ts` to a thin orchestration facade while preserving behavior. Added/wired resolver seam tests into src+dist config test lanes and verified with required gates (`build`, `test:config:dist`, `check:file-budgets`) plus before/after LOC and budget evidence.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Baseline and final LOC metrics recorded in Implementation Notes.
- [x] #2 `bun run build` and `bun run test:config:dist` pass.
- [x] #3 `bun run check:file-budgets` completed and attached in notes.
<!-- DOD:END -->
