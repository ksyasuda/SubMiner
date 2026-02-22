---
id: TASK-85
title: Refactor large files for maintainability and readability
status: Done
assignee: []
created_date: '2026-02-19 09:46'
updated_date: '2026-02-22 07:49'
labels:
  - architecture
  - refactor
  - maintainability
dependencies: []
priority: medium
ordinal: 88000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
This task tracked decomposition of historically oversized/high-coupling core files (`src/main.ts`, `src/anki-integration.ts`, `src/config/service.ts`, `src/core/services/immersion-tracker-service.ts`) plus launcher generated-artifact workflow guardrails and verification.
<!-- SECTION:DESCRIPTION:END -->

## Suggestions

<!-- SECTION:SUGGESTIONS:BEGIN -->
- Use seam tests before each extraction.
- Keep `src/main.ts` + `src/anki-integration.ts` as thin composition/coordinator layers.
- Formalize `subminer` as generated artifact only.
<!-- SECTION:SUGGESTIONS:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Add file-budget guardrails and baseline report.
2. Split `src/main.ts` into runtime domain modules (ongoing; major wiring extracted).
3. Split `src/anki-integration.ts` into focused collaborators (started; constructor composition decomposed).
4. Split `src/config/service.ts` by load/migrate/validate/warn phases (started; load/apply flow decomposition in place).
5. Split immersion tracker service by state, persistence, sync responsibilities.
6. Clarify generated launcher artifact workflow and docs.
7. Run full build/test gate and publish maintainability report.
<!-- SECTION:PLAN:END -->

## Progress Notes

<!-- SECTION:PROGRESS:BEGIN -->
- 2026-02-20: Large `src/main.ts` composition slices extracted into runtime handler modules (startup, CLI, tray/window/bootstrap, shortcuts, OSD/secondary-sub, numeric/overlay shortcut lifecycle, and related seams) with focused parity tests.
- 2026-02-20: `src/anki-integration.ts` constructor decomposed into targeted private factory/config methods (`normalizeConfig`, `createKnownWordCache`, `createPollingRunner`, `createCardCreationService`, `createFieldGroupingService`) to reduce orchestration complexity.
- 2026-02-20: `src/config/service.ts` load/apply duplication reduced by introducing `applyResolvedConfig`, `resolveExistingConfigPath`, and `parseConfigContent`.
- 2026-02-20: added `src/main/runtime/domains/*` barrels plus `src/main/runtime/registry.ts`; migrated `src/main.ts` runtime import paths to domain barrels and added `check:main-fanin` guardrail script.
- 2026-02-21: extracted additional `src/main.ts` deps-builder orchestration into composer modules (`jellyfin-remote`, `anilist-setup`) with focused tests; tightened main fan-in guard (`<=110` import lines, `<=11` unique paths).
- Remaining gap for TASK-85/TASK-94: continue composer extraction for startup/overlay/ipc/shortcuts to reach fully thin composition root.
- Validation checkpoint: `bun run build` and focused suites (`dist/config/config.test.js`, `dist/anki-integration.test.js`) passing.
- Current strategy: prioritize high-ROI extractions (behavioral seams and churn-heavy hotspots) over further low-impact micro-shuffles.
<!-- SECTION:PROGRESS:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `src/main.ts` reduced to orchestration-focused module with extracted runtime domains
- [x] #2 `src/anki-integration.ts` reduced to facade with helper collaborators
- [x] #3 Config and immersion tracker services decomposed without behavior regressions
- [x] #4 `subminer` generated artifact ownership/workflow documented and enforced
- [x] #5 Full build + config/core tests pass after refactor
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
TASK-95 completion slice evidence: decomposed three non-main hotspots and reduced LOC in each targeted file.

TASK-95 LOC deltas: `src/anki-integration.ts` -407 (1722 -> 1315), `src/config/service.ts` -1492 (1591 -> 99), `src/core/services/immersion-tracker-service.ts` -361 (1470 -> 1109).

TASK-95 extracted collaborators: `src/anki-integration/field-grouping-merge.ts`, `src/config/{load.ts,parse.ts,warnings.ts,resolve.ts}`, `src/core/services/immersion-tracker/{types.ts,reducer.ts,queue.ts,maintenance.ts,query.ts}`.

TASK-95 verification evidence: `bun run build`, `bun run test:config:dist`, `bun run test:core:dist`, `bun run check:file-budgets` completed with no failing tests.

2026-02-21 finalization evidence: `src/main.ts` currently 2878 LOC; runtime fan-in guard passes at 85 import lines / 10 unique runtime paths (`bun run check:main-fanin`).

Main-entrypoint scope validation for this task is fan-in/orchestration behavior, not absolute LOC target.

Launcher generated artifact workflow now strictly enforced end-to-end: `scripts/verify-generated-launcher.sh` fails on stale repo-root `./subminer`; CI and release workflows verify `dist/launcher/subminer` and run the verifier script; release checksum/upload paths updated to `dist/launcher/subminer`.

Docs policy updated in `docs/development.md` and `docs/installation.md` to state repo-root `./subminer` is stale/unsupported and generated path must be `dist/launcher/subminer`.

Final verification gate passed: `bun run build && bun run test:config:dist && bun run test:core:dist && bun run check:file-budgets && bun run check:main-fanin`.

TASK-85 AC/DoD ownership map:
- AC#1 (`src/main.ts` orchestration-focused runtime domains): `TASK-94` (primary), `TASK-71` (supporting predecessor)
- AC#2 (`src/anki-integration.ts` facade + collaborators): `TASK-95`
- AC#3 (config + immersion decomposition with no regressions): `TASK-95`
- AC#4 (generated launcher ownership/workflow documented + enforced): `TASK-85` parent finalization evidence
- AC#5 (full build + config/core tests pass): shared verification evidence in `TASK-95`, `TASK-94`, and final TASK-85 gate
- DoD#1 (plan executed/decomposed): `TASK-94` + `TASK-95`
- DoD#2 (regression coverage for extracted seams): `TASK-94` + `TASK-95`
- DoD#3 (docs updated): `TASK-71` architecture docs + `TASK-85` launcher workflow docs

Remaining Scope (ordered):
1. Execution scope: none (all TASK-85 AC/DoD complete).
2. Process hygiene: TASK-93 closure synchronization and provenance note alignment.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Closed remaining TASK-85 scope by enforcing generated launcher ownership (`dist/launcher/subminer`) across local verification, CI, release packaging/checksums, and contributor docs. Revalidated refactor integrity with full build + config/core test gates plus maintainability guardrails (`check:file-budgets`, `check:main-fanin`) and marked parent AC/DoD complete.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Plan at `docs/plans/2026-02-19-repo-maintainability-refactor-plan.md` executed or decomposed into child tasks
- [x] #2 Regression coverage added for extracted seams
- [x] #3 Docs updated for architecture and contributor workflow changes
<!-- DOD:END -->
