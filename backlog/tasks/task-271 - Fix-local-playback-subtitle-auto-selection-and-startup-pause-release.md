---
id: TASK-271
title: Fix local playback subtitle auto-selection and startup pause release
status: Done
assignee:
  - codex
created_date: '2026-04-03 07:47'
updated_date: '2026-04-03 08:03'
labels:
  - bug
  - playback
  - subtitles
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Investigate local-video startup on desktop playback where managed subtitle defaults can bind the wrong primary subtitle track and startup readiness retries can force playback to resume after the user manually pauses. Scope includes fixing the pause/unpause loop and making local subtitle auto-selection prefer the intended primary/secondary tracks for sentence mining sessions.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Desktop local playback no longer forces `pause=no` after the user manually pauses during or after startup readiness handling.
- [x] #2 Managed local subtitle startup selects the expected primary track before secondary track selection for mixed-language subtitle sets like Japanese primary plus English secondary.
- [x] #3 Regression tests cover the pause-release bug and the local subtitle auto-selection behavior.
- [x] #4 Internal docs are updated if runtime behavior or operator expectations change.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add regression coverage for startup autoplay release so duplicate ready handling cannot unpause playback after the user manually pauses the same media.
2. Add regression coverage for managed local subtitle startup selection using configured primary/secondary subtitle language preferences, with JA/EN remaining the default fallback behavior.
3. Extract or add reusable subtitle track ranking logic that prefers configured primary and secondary subtitle languages, with stable scoring for external tracks and non-SDH labels.
4. Update local playback startup/runtime wiring so managed subtitle defaults use explicit ranked selection instead of raw sid=auto / secondary-sid=auto while preserving config-driven language preference ordering.
5. Run focused subtitle/playback tests, then update task notes/final summary with any behavior or docs impact.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented config-aware managed local subtitle selection runtime for local media path changes and playlist-browser local playback rearm, using `youtube.primarySubLanguages` for primary preference and `secondarySub.secondarySubLanguages` for secondary preference with JA/EN fallback defaults.

Updated autoplay ready gate to ignore duplicate readiness signals for the same media so later manual pauses are not overridden by repeated tokenization-ready events.

Updated config/template wording to document that the existing subtitle language preferences now drive managed subtitle auto-selection beyond YouTube-only flows.

Verification: `bun test src/config/config.test.ts src/main/runtime/autoplay-ready-gate.test.ts src/main/runtime/local-subtitle-selection.test.ts src/main/runtime/playlist-browser-runtime.test.ts`; `bun run typecheck`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stopped startup readiness retries from re-unpausing the same media after a later manual pause, and added config-aware managed local subtitle selection so local playback prefers the configured primary/secondary subtitle languages instead of relying on raw mpv `sid=auto` behavior. Added regression coverage for autoplay-ready gating, local subtitle selection, playlist-browser local playback rearm, and config template wording updates.
<!-- SECTION:FINAL_SUMMARY:END -->
