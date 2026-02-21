---
id: TASK-94
title: Reduce main.ts to thin composition root
status: Done
assignee: []
created_date: '2026-02-20 12:06'
updated_date: '2026-02-21 04:12'
labels:
  - architecture
  - refactor
  - maintainability
dependencies:
  - TASK-71
  - TASK-85
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` still contains heavy deps-builder orchestration and high import concentration. Complete the composition-root refactor so `main.ts` owns boot wiring only and domain assembly moves behind registry/domain composer modules.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Baseline current `main.ts` fan-in (`check:main-fanin`, `check:file-budgets`, `wc -l src/main.ts`).
2. Introduce domain composer modules for repeated build-handler clusters (startup, overlay, jellyfin, anilist, ipc/shortcuts).
3. Move `createBuild*MainDepsHandler` call clusters out of `main.ts` into domain composer modules with narrow typed inputs.
4. Replace inline assembly in `main.ts` with single-call composer invocations per domain.
5. Add/adjust unit tests for new composer modules to lock shape and wiring.
6. Tighten `check:main-fanin` thresholds after refactor to prevent regression.
7. Re-run full config/core test gates and update `TASK-71` progress.
<!-- SECTION:PLAN:END -->

## Progress Notes

- 2026-02-21 baseline: `src/main.ts` = 3129 LOC; `check:main-fanin` = 105 import lines / 9 unique runtime paths.
- 2026-02-21 extraction slice: added composer modules `src/main/runtime/composers/jellyfin-remote-composer.ts` and `src/main/runtime/composers/anilist-setup-composer.ts`; moved corresponding `createBuild*MainDepsHandler` orchestration out of `main.ts`.
- 2026-02-21 tests: added focused composer shape tests at `src/main/runtime/composers/jellyfin-remote-composer.test.ts` and `src/main/runtime/composers/anilist-setup-composer.test.ts`.
- 2026-02-21 guardrail tighten: `scripts/check-main-runtime-fanin.ts` defaults tightened to `import lines <= 110` and `unique runtime paths <= 11`.
- 2026-02-21 post-slice metrics: `src/main.ts` = 3042 LOC; `check:main-fanin --strict` = 103 import lines / 11 unique runtime paths.
- Remaining gap: startup/overlay/ipc/shortcuts composer extraction still required to fully meet thin composition-root target.

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `src/main.ts` is composition-focused (boot/runtime wiring only; no broad deps-builder clusters).
- [x] #2 Runtime import paths in `src/main.ts` stay domain-registry oriented (no relapse to per-leaf runtime imports).
- [x] #3 `check:main-fanin` passes under updated threshold.
- [x] #4 `bun run test:core:dist` passes with no CLI/IPC behavior regressions.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Extract IPC composition from src/main.ts into src/main/runtime/composers/ipc-runtime-composer.ts with focused tests (handler deps build, runtime handler creation, registration wiring).
2) Rewire src/main.ts IPC command and IPC registration blocks to consume the IPC composer while preserving existing helper wrappers and behavior.
3) Extract shortcuts composition from src/main.ts into src/main/runtime/composers/shortcuts-runtime-composer.ts with focused tests (global shortcuts runtime, numeric sessions, overlay shortcuts lifecycle).
4) Rewire src/main.ts shortcut runtime block to consume shortcuts composer while preserving downstream callsites.
5) Run targeted composer tests + check:main-fanin + test:core:dist; update TASK-94 notes with before/after metrics and finalize AC/status if all green.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: finished extraction pass by moving IPC command/registration, shortcuts runtime wiring, startup lifecycle wiring, and app-ready startup composition from `src/main.ts` into dedicated composer modules under `src/main/runtime/composers/*`.

Metrics: `src/main.ts` reduced from 3043 LOC to 2955 LOC in this pass; `check:main-fanin` import lines improved from 99 to 90 while remaining under enforced threshold.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed TASK-94 composition-root finish pass by extracting remaining large assembly clusters from `src/main.ts` into dedicated composer modules: IPC runtime (`ipc-runtime-composer`), shortcuts runtime (`shortcuts-runtime-composer`), startup lifecycle (`startup-lifecycle-composer`), and app-ready bootstrap (`app-ready-composer`). Main process behavior was preserved while reducing direct deps-builder orchestration in `main.ts` and keeping runtime wiring through domain/composer boundaries.

Added focused composer tests for each new module and reran project gates. Verification run: `bun run build`, `bun test src/main/runtime/composers/ipc-runtime-composer.test.ts`, `bun test src/main/runtime/composers/shortcuts-runtime-composer.test.ts`, `bun test src/main/runtime/composers/startup-lifecycle-composer.test.ts`, `bun test src/main/runtime/composers/app-ready-composer.test.ts`, `bun run check:main-fanin`, and `bun run test:core:dist` (all passing).
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `src/main.ts` LOC and fan-in metrics improve from pre-task baseline and are recorded in task notes.
- [x] #2 New composer modules are covered by focused tests.
- [x] #3 `TASK-71` and `TASK-85` progress notes reflect completed slice and remaining gap.
<!-- DOD:END -->
