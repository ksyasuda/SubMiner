---
id: TASK-131
title: Avoid duplicate tokenization warmup after background startup
status: Done
assignee:
  - codex
created_date: '2026-03-08 10:12'
updated_date: '2026-03-08 12:00'
labels:
  - bug
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/composers/mpv-runtime-composer.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/startup-warmups.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/composers/mpv-runtime-composer.test.ts
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

When SubMiner is already running in the background and mpv is launched from the launcher or mpv plugin, the live app should reuse startup tokenization warmup state instead of re-entering the Yomitan/tokenization/annotation warmup path on first overlay use.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Background startup tokenization warmup is recorded in the runtime state used by later mpv/tokenization flows.
- [x] #2 Launching mpv from the launcher or plugin against an already-running background app does not re-run duplicate Yomitan/tokenization annotation warmup work in the live process.
- [x] #3 Regression tests cover the warmed-background path and protect against re-entering duplicate warmup work.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a regression test covering the case where background startup warmups already completed and a later tokenize call must not re-enter Yomitan/MeCab/dictionary warmups.
2. Update mpv tokenization warmup composition so startup background warmups and on-demand tokenization share the same completion state.
3. Run the focused composer/runtime tests and update acceptance criteria/notes with results.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Root-cause hypothesis: startup background warmups and on-demand tokenization warmups use separate state, so later mpv launch can re-enter warmup bookkeeping even though background startup already warmed dependencies.

Implemented shared warmup state between startup background warmups and on-demand tokenization warmups by forwarding scheduled Yomitan/tokenization promises into the mpv runtime composer. Added regression coverage for the warmed-background path. Verified with `bun run test:fast` plus focused composer/startup warmup tests.

Follow-up root cause from live retest: second mpv open could still pause on the startup gate because the runtime only treated full background tokenization warmup completion as reusable readiness. In practice, first-file tokenization could already be ready while slower dictionary prewarm work was still finishing, so reopening a video waited on duplicate warmup completion even though annotations were already usable.

Adjusted `src/main/runtime/composers/mpv-runtime-composer.ts` so autoplay reuse keys off a separate playback-ready latch. The latch flips true either when background warmups fully cover tokenization or when `onTokenizationReady` fires for a real subtitle line. `src/main.ts` already uses `isTokenizationWarmupReady()` to fast-signal `subminer-autoplay-ready` on a fresh media-path change, so reopened videos can now resume immediately once tokenization has succeeded once in the persistent app.

Validation update: `bun test src/core/services/cli-command.test.ts src/main/runtime/mpv-main-event-actions.test.ts src/main/runtime/composers/mpv-runtime-composer.test.ts launcher/mpv.test.ts launcher/smoke.e2e.test.ts` passed, `lua scripts/test-plugin-start-gate.lua` passed, and `bun run typecheck` passed.

<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Background startup tokenization warmups now feed the same in-memory warmup state used by later mpv tokenization. When the app is already running and warmed in the background, launcher/plugin-driven mpv startup reuses that state instead of re-entering Yomitan/tokenization annotation warmups. Added a regression test for the warmed-background path and verified with `bun run test:fast`.

A later follow-up fixed the remaining second-open delay: autoplay reuse no longer waits for the entire background dictionary warmup pipeline to finish. After the persistent app has produced one tokenization-ready event, later mpv reconnects reuse that readiness immediately, so reopening the same or another video does not pause again on duplicate warmup bookkeeping.

<!-- SECTION:FINAL_SUMMARY:END -->
