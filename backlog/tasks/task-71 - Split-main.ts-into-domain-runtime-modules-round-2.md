---
id: TASK-71
title: Split main.ts into domain runtime modules round 2
status: In Progress
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-21 03:40'
labels:
  - architecture
  - refactor
  - maintainability
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` remains >3k LOC and still mixes composition, runtime state mutation, domain workflows (Jellyfin, AniList), tray/UI setup, and IPC wiring. This creates high cognitive load and broad change blast radius.
<!-- SECTION:DESCRIPTION:END -->

## Progress Notes

- Added runtime domain barrels under `src/main/runtime/domains/*` and composed registry entrypoint at `src/main/runtime/registry.ts`.
- Migrated `src/main.ts` runtime imports from per-file runtime modules to domain barrel import paths.
- Added `scripts/check-main-runtime-fanin.ts` guardrail and package scripts (`check:main-fanin`, `check:main-fanin:strict`).
- 2026-02-21: extracted jellyfin/anilist deps-builder clusters into composer modules (`src/main/runtime/composers/jellyfin-remote-composer.ts`, `src/main/runtime/composers/anilist-setup-composer.ts`) and replaced inline `main.ts` orchestration with composer invocations.
- 2026-02-21: tightened fan-in guard defaults to `import lines <= 110` and `unique runtime paths <= 11`; strict mode passing.
- Remaining TASK-71 scope: finish startup/overlay/ipc/shortcuts composer extraction to fully thin `src/main.ts`.

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Extract cohesive runtime modules:
   - AniList runtime orchestration
   - Jellyfin runtime orchestration
   - tray runtime
   - MPV event binding/runtime hooks
2. Keep `main.ts` as composition root only (state wiring + module registration).
3. Define narrow service/deps interfaces per module to reduce closure capture.
4. Add/adjust unit tests for extracted modules.
5. Preserve existing CLI/IPC behavior exactly.
6. Update architecture docs with new ownership boundaries.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 `src/main.ts` responsibilities reduced to composition/wiring concerns
- [ ] #2 Extracted runtime modules have focused interfaces and isolated tests
- [ ] #3 No CLI/IPC regressions in existing test suite
- [ ] #4 Docs reflect new module boundaries
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 `bun run test:fast` passes
- [ ] #2 Architecture docs updated
<!-- DOD:END -->
