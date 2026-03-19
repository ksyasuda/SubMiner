---
id: TASK-197
title: Eliminate per-line plain subtitle flash on prefetch cache hit
status: Done
assignee: []
created_date: '2026-03-18 16:28'
labels: []
dependencies:
  - TASK-196
references:
  - /home/sudacode/projects/japanese/SubMiner/src/core/services/subtitle-processing-controller.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/mpv-main-event-actions.ts
  - /home/sudacode/projects/japanese/SubMiner/src/main/runtime/mpv-main-event-main-deps.ts
documentation: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Remove the remaining small per-line subtitle annotation delay after prefetch warmup by avoiding the unconditional plain-subtitle broadcast on mpv subtitle-change events when a cached annotated payload already exists.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 On a subtitle cache hit, the mpv subtitle-change path can emit annotated subtitle payload synchronously instead of first broadcasting `tokens: null`.
- [x] #2 Cache-miss behavior still preserves immediate plain-text subtitle display while async tokenization runs.
- [x] #3 Regression tests cover the controller cache-consume path and the mpv subtitle-change handler cache-hit branch.
- [x] #4 Verification covers the touched core/runtime lane.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add failing tests for controller cache consumption and mpv subtitle-change immediate annotated emission.
2. Add a controller method that consumes cached subtitle payload synchronously while updating internal latest/emitted state.
3. Wire the mpv subtitle-change handler to use the immediate cached payload when present, falling back to the existing plain-text path on misses.
4. Run focused tests and the cheapest sufficient verification lane.
<!-- SECTION:PLAN:END -->

## Outcome

<!-- SECTION:OUTCOME:BEGIN -->
Added `consumeCachedSubtitle` to the subtitle processing controller so cache hits can be claimed synchronously without reprocessing, then wired the mpv subtitle-change handler to emit cached annotated payloads immediately while preserving the existing plain-text fallback for misses. Verified with focused unit tests plus the `runtime-compat` lane.
<!-- SECTION:OUTCOME:END -->
