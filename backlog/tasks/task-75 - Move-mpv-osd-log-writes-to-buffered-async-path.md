---
id: TASK-75
title: Move mpv OSD log writes to buffered async path
status: Done
assignee: []
created_date: '2026-02-18 11:35'
updated_date: '2026-02-21 23:48'
labels:
  - logging
  - performance
  - main-process
dependencies: []
priority: low
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
`appendToMpvLog` in `src/main.ts` uses synchronous fs operations on OSD writes. Frequent OSD activity can block the event loop and adds avoidable latency in main process runtime paths.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Replace sync file writes with buffered async logger flow.
2. Ensure log directory creation and write failures remain best-effort (non-fatal).
3. Add flush behavior for shutdown to avoid dropping trailing log lines.
4. Keep log format/path unchanged unless explicitly intended.
5. Add tests for non-blocking behavior and graceful error handling.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 OSD log path no longer uses synchronous fs calls
- [x] #2 Logging remains best-effort and does not break runtime behavior on fs failure
- [x] #3 Shutdown flush behavior verified by tests
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1) Add failing tests for buffered async MPV OSD log behavior + explicit flush semantics in mpv-osd runtime tests.
2) Replace sync mpv-osd log deps/handler with buffered async queue/drain implementation and expose flushMpvLog runtime surface.
3) Wire flushMpvLog into onWillQuit cleanup deps/action path and main.ts runtime composition.
4) Add shutdown-flush assertions in lifecycle cleanup tests.
5) Run `bun run build && bun run test:core:dist`, then record AC/DoD evidence and finalize task metadata.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
2026-02-21: Started execution in codex session `codex-task75-mpv-osd-buffered-20260221T231816Z-yj32`. Running writing-plans then executing-plans workflow; no commit requested.

Implemented buffered async MPV OSD log path in `src/main/runtime/mpv-osd-log.ts` with queued drains and explicit `flushMpvLog()` surface; removed sync fs dependency usage from runtime wiring in `src/main.ts`.

Wired shutdown flush into startup lifecycle cleanup via `flushMpvLog` dependency in `app-lifecycle-main-cleanup` + `app-lifecycle-actions` so quit cleanup invokes pending log drain before socket destroy.

Added/updated tests for async queueing, best-effort fs failures, runtime composition flush surface, and shutdown cleanup flush invocation across mpv-osd and lifecycle runtime tests.

Verification: `bun run test:core:src` PASS (219 pass, 0 fail); focused runtime tests PASS via bundled node lane: `bun build ... --outdir /tmp/task75-tests && node --test /tmp/task75-tests/*.js`.

`bun run build` currently fails on pre-existing unrelated tokenizer typing issue at `src/core/services/tokenizer.ts` (`Logger` not assignable to `LoggerLike`).
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Moved MPV OSD log writes off sync fs APIs by introducing buffered async queue/drain behavior with explicit flush support and preserving best-effort failure semantics. Integrated shutdown flush into on-will-quit cleanup and added targeted runtime/lifecycle tests to verify queue flush + quit-time drain behavior without user-visible OSD regression.
<!-- SECTION:FINAL_SUMMARY:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [x] #1 Core tests pass with new logging path
- [x] #2 No user-visible regression in MPV OSD/log output behavior
<!-- DOD:END -->
