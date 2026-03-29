---
id: TASK-238.6
title: Extract remaining inline runtime logic and composer gaps from src/main.ts
status: Done
assignee: []
created_date: '2026-03-27 00:00'
updated_date: '2026-03-27 22:13'
labels:
  - tech-debt
  - runtime
  - maintainability
  - composers
milestone: m-0
dependencies:
  - TASK-238.1
  - TASK-238.2
references:
  - src/main.ts
  - src/main/runtime/youtube-flow.ts
  - src/main/runtime/autoplay-ready-gate.ts
  - src/main/runtime/subtitle-prefetch-init.ts
  - src/main/runtime/discord-presence-runtime.ts
  - src/main/overlay-modal-state.ts
  - src/main/runtime/composers
parent_task_id: TASK-238
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`src/main.ts` still mixes two concerns: pure dependency wiring and inline runtime logic. The earlier composer extractions reduce the wiring burden, but the file still owns several substantial behavior blocks and a few large inline dependency groupings. This task tracks the next maintainability pass: move the remaining runtime logic into the appropriate domain modules, add missing composer wrappers for the biggest grouped handler blocks, and reassess whether a boot-phase split is still necessary after the entrypoint becomes mostly wiring.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [x] #1 `runYoutubePlaybackFlow`, `maybeSignalPluginAutoplayReady`, `refreshSubtitlePrefetchFromActiveTrack`, `publishDiscordPresence`, and `handleModalInputStateChange` no longer live as substantial inline logic in `src/main.ts`.
- [x] #2 The large subtitle/prefetch, stats startup, and overlay visibility dependency groupings are wrapped behind named composer helpers instead of remaining inline in `src/main.ts`.
- [x] #3 `src/main.ts` reads primarily as a boot and lifecycle coordinator, with domain behavior concentrated in named runtime modules.
- [x] #4 Focused tests cover the extracted behavior or the new composer surfaces.
- [x] #5 The task records whether the remaining size still justifies a boot-phase split or whether that follow-up can wait.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
Recommended sequence:

1. Let the current composer and `setup-window-factory` work land first so this slice starts from a stable wiring baseline.
2. Extract the five inline runtime functions into their natural domain modules or direct equivalents.
3. Add or extend composer helpers for subtitle/prefetch, stats startup, and overlay visibility handler grouping.
4. Re-scan `src/main.ts` after the extraction and decide whether a boot-phase split is still the right next task.
5. Verify the extracted behavior with focused tests first, then run the relevant broader runtime gate if the slice crosses startup boundaries.

Guardrails:

- Keep the work behavior-preserving.
- Prefer moving logic to existing runtime surfaces over creating new giant helper files.
- Do not expand into unrelated `src/main.ts` cleanup that is already tracked by other TASK-238 slices.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Extracted the remaining inline runtime seams from `src/main.ts` into focused runtime modules:
`src/main/runtime/youtube-playback-runtime.ts`,
`src/main/runtime/autoplay-ready-gate.ts`,
`src/main/runtime/subtitle-prefetch-runtime.ts`,
`src/main/runtime/discord-presence-runtime.ts`,
and `src/main/runtime/overlay-modal-input-state.ts`.

Added named composer wrappers for the grouped subtitle/prefetch, stats startup, and overlay visibility wiring in `src/main/runtime/composers/`.

Re-scan result for the boot-phase split follow-up: the entrypoint is materially closer to a boot/lifecycle coordinator now, so TASK-238.7 remains a valid future cleanup but no longer feels urgent or blocking for maintainability.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
TASK-238.6 is complete. Verification passed with `bun run typecheck`, focused runtime/composer tests, `bun run test:fast`, `bun run test:env`, and `bun run build`. The remaining `src/main.ts` work is now better isolated behind runtime modules and composer helpers, and the boot-phase split can wait for a later cleanup pass instead of being treated as immediate follow-on work.

Backlog completion now includes changelog artifact `changes/2026-03-27-task-238.6-main-runtime-refactor.md` under runtime internals.
<!-- SECTION:FINAL_SUMMARY:END -->
