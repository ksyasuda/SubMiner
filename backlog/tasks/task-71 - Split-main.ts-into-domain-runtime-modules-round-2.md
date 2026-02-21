---
id: TASK-71
title: Split main.ts into domain runtime modules round 2
status: Done
assignee:
  - codex
created_date: '2026-02-18 11:35'
updated_date: '2026-02-21 04:57'
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
- [x] #1 `src/main.ts` responsibilities reduced to composition/wiring concerns
- [x] #2 Extracted runtime modules have focused interfaces and isolated tests
- [x] #3 No CLI/IPC regressions in existing test suite
- [x] #4 Docs reflect new module boundaries
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Extract AniList tracking composition from src/main.ts (refresh/guess/probe/retry/update handler assembly) into src/main/runtime/composers/anilist-tracking-composer.ts with focused seam tests.
2) Extract MPV runtime composition from src/main.ts (bind event handlers, MPV client factory, subtitle render metrics, tokenizer warmups) into src/main/runtime/composers/mpv-runtime-composer.ts with focused seam tests.
3) Rewire src/main.ts to consume both composers while preserving existing local helper contracts and downstream behavior.
4) Update architecture docs to reflect composition boundary under src/main/runtime/composers/*.
5) Run verification gate (build + check:main-fanin + test:core:dist), then record TASK-71 notes/AC/DoD and finalize if green.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: extracted two remaining high-noise main-process composition clusters from `src/main.ts` into dedicated runtime composers: `src/main/runtime/composers/anilist-tracking-composer.ts` (AniList media tracking/probe/retry/update assembly) and `src/main/runtime/composers/mpv-runtime-composer.ts` (MPV event binding/factory/metrics/tokenizer/warmups assembly).

Rewired `src/main.ts` to consume composer outputs while preserving existing helper contracts/call sites (`refreshAnilistClientSecretState`, `processNextAnilistRetryUpdate`, `maybeRunAnilistPostWatchUpdate`, `createMpvClientRuntimeService`, `updateMpvSubtitleRenderMetrics`, `tokenizeSubtitle`, `startBackgroundWarmups`).

Added focused seam tests: `src/main/runtime/composers/anilist-tracking-composer.test.ts` and `src/main/runtime/composers/mpv-runtime-composer.test.ts`.

Added `src/main/runtime/composers/index.ts` barrel and updated main imports to keep runtime fan-in strict gate green.

Docs updated for new boundaries in `docs/architecture.md` and `docs/development.md`.

Verification: `bun run build`, `bun test src/main/runtime/composers/anilist-tracking-composer.test.ts src/main/runtime/composers/mpv-runtime-composer.test.ts`, `bun run check:main-fanin:strict` (85 import lines / 10 unique runtime paths, pass), and `bun run test:fast` (pass).

`src/main.ts` LOC reduced from 2956 to 2875 in this pass.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Completed TASK-71 round 2 by extracting AniList tracking and MPV runtime assembly clusters out of `src/main.ts` into dedicated composer modules with focused seam tests, then rewiring the composition root to consume those modules without changing runtime behavior. Added a composers barrel to reduce runtime import fan-in and updated architecture/development docs to reflect composer-first boundaries; verification gates (`build`, `check:main-fanin:strict`, composer tests, `test:fast`) all pass.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 `bun run test:fast` passes
- [x] #2 Architecture docs updated
<!-- DOD:END -->
