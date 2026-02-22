---
id: TASK-95
title: Decompose remaining oversized core hotspots
status: Done
assignee:
  - codex-task95-hotspots
created_date: '2026-02-20 12:06'
updated_date: '2026-02-21 03:29'
labels:
  - architecture
  - refactor
  - maintainability
dependencies:
  - TASK-85
references:
  - docs/plans/2026-02-21-task-95-hotspot-decomposition-plan.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Three core files remain materially oversized and unfinished: `src/anki-integration.ts`, `src/config/service.ts`, and `src/core/services/immersion-tracker-service.ts`. Complete decomposition into focused collaborators/modules with behavior preserved and regression coverage.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Capture per-file baseline LOC and key responsibilities (constructor orchestration, phase lifecycle, persistence/sync/event flow).
2. For `src/anki-integration.ts`: extract facade collaborators by domain (config/field resolution/media/card creation/notifications) and reduce class surface to orchestration.
3. For `src/config/service.ts`: split lifecycle phases into explicit modules (`load`, `migrate`, `validate`, `warnings`) and keep API stable.
4. For `src/core/services/immersion-tracker-service.ts`: separate state/reducer, persistence adapter, and sync/reporting orchestration.
5. Add focused tests for extracted modules and preserve existing integration-level assertions.
6. Run full build + config/core tests + file-budget checks; record before/after LOC metrics.
7. Update `TASK-85` AC/DoD linkage to mark completed sub-scope with evidence.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `src/anki-integration.ts`, `src/config/service.ts`, and `src/core/services/immersion-tracker-service.ts` are each reduced with clear collaborator boundaries.
- [x] #2 Public behavior remains unchanged (existing config/core tests pass).
- [x] #3 New seam tests cover extracted collaborators/modules.
- [x] #4 File-budget report reflects measurable LOC reduction on all three hotspots.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Execution plan of record: `docs/plans/2026-02-21-task-95-hotspot-decomposition-plan.md`

1. Capture baseline LOC and file-budget status for the 3 hotspot files.
2. Add characterization seam tests before extraction in each domain (Anki, Config, Immersion).
3. Decompose each hotspot behind stable facades by extracting focused collaborators/modules.
4. Use parallel subagents for independent hotspot workstreams; integrate safely on shared boundaries.
5. Run required gates (`bun run build`, `bun run test:config:dist`, `bun run test:core:dist`, `bun run check:file-budgets`).
6. Record before/after LOC plus ownership map; update TASK-95 AC/DoD and TASK-85 progress evidence.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Baseline LOC (before): `src/anki-integration.ts` 1722, `src/config/service.ts` 1591, `src/core/services/immersion-tracker-service.ts` 1470.

After LOC: `src/anki-integration.ts` 1315 (-407), `src/config/service.ts` 99 (-1492), `src/core/services/immersion-tracker-service.ts` 1109 (-361).

Module ownership map: Anki -> `src/anki-integration/field-grouping-merge.ts`; Config -> `src/config/{load.ts,parse.ts,warnings.ts,resolve.ts}`; Immersion -> `src/core/services/immersion-tracker/{types.ts,reducer.ts,queue.ts,maintenance.ts,query.ts}`.

Seam tests added: `src/anki-integration.test.ts` (field-grouping merge seams), `src/config/config.test.ts` (loader/strict reload/warning determinism seams), `src/core/services/immersion-tracker-service.test.ts` (queue/reducer/month-key seams).

Verification gates passed: `bun run build`, `bun run test:config:dist` (43 pass), `bun run test:core:dist` (204 pass, 10 skipped, 0 fail), `bun run check:file-budgets` (warning-mode report with all three hotspot LOC reduced).
<!-- SECTION:NOTES:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Refactor notes include before/after LOC and module ownership map for all three files.
- [x] #2 `bun run build`, `bun run test:config:dist`, `bun run test:core:dist` pass.
- [x] #3 `TASK-85` progress/AC reflects this completion slice with evidence links.
<!-- DOD:END -->
