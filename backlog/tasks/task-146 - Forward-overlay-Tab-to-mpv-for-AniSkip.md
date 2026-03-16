---
id: TASK-146
title: Forward overlay Tab to mpv for AniSkip
status: Done
assignee:
  - codex
created_date: '2026-03-09 00:00'
updated_date: '2026-03-16 05:13'
labels:
  - bug
  - overlay
  - aniskip
  - linux
dependencies: []
ordinal: 46500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix visible-overlay keyboard handling so bare `Tab` is forwarded to mpv instead of being consumed by Electron focus navigation. This restores the default AniSkip `TAB` binding while the overlay has focus, especially on Linux.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Visible overlay forwards bare `Tab` to mpv as `keypress TAB`.
- [x] #2 Modal overlays keep their existing local `Tab` behavior.
- [x] #3 Automated regression coverage exists for the input handler and overlay factory wiring.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a failing regression around visible-overlay `before-input-event` handling for bare `Tab`.
2. Add/extend overlay factory tests so the new mpv-forward callback is wired through runtime construction.
3. Patch overlay input handling to intercept visible-overlay `Tab` and send mpv `keypress TAB`.
4. Run focused overlay tests, typecheck, and changelog validation.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Extracted visible-overlay input handling into `src/core/services/overlay-window-input.ts` so the `Tab` forwarding decision can be unit tested without loading Electron window primitives.

Visible overlay `before-input-event` now intercepts bare `Tab`, prevents the browser default, and forwards mpv `keypress TAB` through the existing mpv runtime command path. Modal overlays remain unchanged.

Verification:

- `bun test src/core/services/overlay-window.test.ts src/main/runtime/overlay-window-factory.test.ts src/main/runtime/overlay-window-factory-main-deps.test.ts src/main/runtime/overlay-window-runtime-handlers.test.ts`
- `bun x tsc --noEmit`
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Visible overlay focus no longer blocks the default AniSkip `Tab` binding. Bare `Tab` is now forwarded straight to mpv while the visible overlay is active, and modal overlays still retain their own normal focus behavior.

Added regression coverage for both the input-routing decision and the runtime plumbing that carries the new mpv forwarder into overlay window creation.
<!-- SECTION:FINAL_SUMMARY:END -->
