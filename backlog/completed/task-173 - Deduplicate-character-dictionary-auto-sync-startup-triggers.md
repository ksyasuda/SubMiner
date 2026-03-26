---
id: TASK-173
title: Deduplicate character dictionary auto-sync startup triggers
status: Done
assignee: []
created_date: '2026-03-16 11:05'
updated_date: '2026-03-16 11:20'
labels:
  - bug
  - character-dictionary
  - startup
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/mpv-client-event-bindings.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/mpv-main-event-actions.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/character-dictionary-auto-sync.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/character-dictionary-runtime.ts
priority: medium
ordinal: 36500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Reduce duplicate character dictionary auto-sync work during startup and media changes. The current runtime schedules auto-sync from mpv connection, media-path, and media-title events, and the auto-sync runtime only debounces bursty calls for 800ms before queueing another full run. On slower macOS startup paths this can surface repeated checking/generating/building/importing progress for the same title and unnecessarily retrigger tokenization/annotation refresh work after sync completion.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Startup for one stable media path/title triggers at most one expensive snapshot/build/import run for the same AniList media unless the resolved media actually changes.
- [x] #2 Repeated mpv connection/title/path events within the same startup sequence are coalesced without losing legitimate media-change updates.
- [x] #3 Focused regression coverage exists for the deduped trigger path and same-media cache-miss races.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Reduced the auto-sync trigger surface to mpv `media-path-change` only. `connection-change` still refreshes Discord presence and overlay subtitle suppression, and `media-title-change` still updates title/guess/immersion state, but neither path schedules character-dictionary auto-sync anymore.

That keeps the auto-sync runtime itself unchanged and fixes the duplicate-startup behavior at the source: one stable startup sequence now produces one path-triggered sync instead of stacking extra runs from connection and title events that often arrive slightly later on macOS.

Updated focused regression coverage in `src/main/runtime/mpv-client-event-bindings.test.ts` and `src/main/runtime/mpv-main-event-actions.test.ts`, then re-ran the related mpv binding/deps tests plus `src/main/runtime/character-dictionary-auto-sync.test.ts`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed repeated character-dictionary startup work by stopping auto-sync scheduling from mpv `connection-change` and `media-title-change`; only `media-path-change` now triggers the sync. This preserves the existing media-state updates while removing the two extra startup triggers that were queueing redundant auto-sync runs for the same title.

Verification:
- `bun test src/main/runtime/mpv-client-event-bindings.test.ts src/main/runtime/mpv-main-event-actions.test.ts`
- `bun test src/main/runtime/mpv-client-event-bindings.test.ts src/main/runtime/mpv-main-event-actions.test.ts src/main/runtime/mpv-main-event-bindings.test.ts src/main/runtime/mpv-main-event-main-deps.test.ts src/main/runtime/character-dictionary-auto-sync.test.ts`
- `bun run typecheck`
<!-- SECTION:FINAL_SUMMARY:END -->
