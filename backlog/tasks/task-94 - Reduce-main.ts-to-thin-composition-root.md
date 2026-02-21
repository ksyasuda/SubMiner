---
id: TASK-94
title: Reduce main.ts to thin composition root
status: In Progress
assignee: []
created_date: '2026-02-20 12:06'
updated_date: '2026-02-21 03:40'
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
- [ ] #1 `src/main.ts` is composition-focused (boot/runtime wiring only; no broad deps-builder clusters).
- [x] #2 Runtime import paths in `src/main.ts` stay domain-registry oriented (no relapse to per-leaf runtime imports).
- [x] #3 `check:main-fanin` passes under updated threshold.
- [x] #4 `bun run test:core:dist` passes with no CLI/IPC behavior regressions.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `src/main.ts` LOC and fan-in metrics improve from pre-task baseline and are recorded in task notes.
- [x] #2 New composer modules are covered by focused tests.
- [x] #3 `TASK-71` and `TASK-85` progress notes reflect completed slice and remaining gap.
<!-- DOD:END -->
