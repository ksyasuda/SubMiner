---
id: TASK-223
title: Fix YouTube overlay Anki initialization regression
status: In Progress
assignee:
  - codex
created_date: '2026-03-23 08:41'
updated_date: '2026-03-23 18:24'
labels:
  - bug
  - youtube
  - anki
dependencies: []
references:
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-runtime-init.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/main/runtime/cli-command-runtime-handler.ts
documentation:
  - /Users/sudacode/projects/japanese/SubMiner/docs/workflow/verification.md
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Restore Anki-backed lookup and known-word behavior during YouTube playback. Recent startup changes appear to let the YouTube flow initialize the overlay before runtime prerequisites exist, leaving the Anki integration unavailable for popup Mine actions and known-word highlighting.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 YouTube playback initializes Anki integration once overlay startup prerequisites are available so lookup can offer card-add actions again
- [x] #2 Known-word / N+1 state is available during YouTube playback when the user has Anki-backed known-word highlighting enabled
- [x] #3 Regression coverage fails before the fix and passes after it for the YouTube startup path
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a regression test covering the YouTube playback command path and assert overlay startup prerequisites are established before the flow runs.
2. Reuse the overlay startup prerequisite bootstrap for the YouTube playback path so Anki integration sees subtitle tracker, mpv client, and runtime options manager before initialization.
3. Verify with focused runtime/CLI tests, then run the cheapest sufficient verification lane for the touched files.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Identified regression path in CLI command runtime: YouTube playback commands could reach overlay initialization without first materializing overlay startup prerequisites, leaving Anki integration unavailable during the initial startup attempt.

Added a regression test at src/main/runtime/cli-command-runtime-handler.test.ts covering youtubePlay command dispatch outside texthooker-only mode.

Verified with bun test src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-prechecks.test.ts src/main/runtime/cli-command-prechecks-main-deps.test.ts and bun test src/core/services/cli-command.test.ts src/main/runtime/cli-command-context-main-deps.test.ts src/main/runtime/cli-command-prechecks-main-deps.test.ts src/main/runtime/cli-command-runtime-handler.test.ts.

Removed the YouTube-only ensureOverlayRuntimeReady path from src/main.ts after confirming regular app startup already loads Yomitan and the shared CLI overlay pre-dispatch bootstrap now covers overlay prerequisites.

Moved overlay bootstrap into the generic initial-args startup path for any initial command that needs overlay runtime, so overlay prerequisites and overlay initialization happen before CLI dispatch instead of inside a YouTube-only or last-moment command path.

Additional verification passed: bun test src/main/runtime/initial-args-runtime-handler.test.ts src/main/runtime/initial-args-handler.test.ts src/main/runtime/initial-args-main-deps.test.ts, bun run typecheck, bun run test:runtime:compat

Follow-up regression: subtitle picker was still auto-submitting the default selection during YouTube startup. Investigating renderer-side immediate Enter key bleed-through on picker open.

Root cause for remaining picker regression: the YouTube track picker accepted Enter immediately on open, so the launch keypress could auto-submit the default track selection before the modal was visible to the user.

Added renderer regression coverage in src/renderer/modals/youtube-track-picker.test.ts proving immediate Enter after open is ignored and a later Enter still submits normally.

Implemented a 200ms open-key guard in src/renderer/modals/youtube-track-picker.ts for Enter-based submission only; Escape/click behavior unchanged.

New follow-up regression report: YouTube subtitle picker can open before the mpv playback window is ready, leaving the picker behind the overlay after geometry snaps into place. Investigating picker-open gating and modal-targeting timing.

Identified likely cause of picker-behind-overlay regression: YouTube picker open logic mixed overlay targets. First attempt preferred the visible main overlay, timeout retry switched to the dedicated modal window, allowing a late first open to cover the modal.

Extracted picker-open policy into src/main/runtime/youtube-picker-open.ts and changed YouTube picker startup to always target the dedicated modal window, including retries. This keeps the picker on a single window path and lets overlay-runtime hide/click-through the main overlay while the modal is active.

Added regression tests in src/main/runtime/youtube-picker-open.test.ts covering dedicated modal first-open, dedicated-modal retry, and failure when no modal target is available.

User reports overlay flow still feels wrong: YouTube path appears to preload subtitles before mandatory selection and may open the picker before mpv window readiness. Re-evaluating flow design against regular video startup before further implementation.

New follow-up regression report: duplicate overlay windows appear during YouTube playback and only one window shows subtitles. Investigating main-overlay versus dedicated modal-window handoff/cleanup.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the YouTube Anki initialization regression by making CLI commands that require overlay runtime bootstrap overlay startup prerequisites before command dispatch when not in texthooker-only mode. This ensures the YouTube playback flow has the mpv client, runtime options manager, and subtitle timing tracker ready before overlay/Anki initialization runs, restoring Mine actions and known-word-backed behavior.

Added a regression test covering youtubePlay command dispatch in src/main/runtime/cli-command-runtime-handler.test.ts.

Verification:
- bun test src/main/runtime/cli-command-runtime-handler.test.ts src/main/runtime/cli-command-prechecks.test.ts src/main/runtime/cli-command-prechecks-main-deps.test.ts
- bun test src/core/services/cli-command.test.ts src/main/runtime/cli-command-context-main-deps.test.ts src/main/runtime/cli-command-prechecks-main-deps.test.ts src/main/runtime/cli-command-runtime-handler.test.ts

Updated the fix to avoid a YouTube-specific startup path: removed the dedicated ensureOverlayRuntimeReady helper from src/main.ts and relied on the shared CLI overlay prerequisite bootstrap instead.

Additional verification: bun run typecheck

Follow-up adjustment: initial overlay-runtime commands now bootstrap overlay prerequisites and initialize overlay during the shared initial-args startup path, rather than waiting for command dispatch. This keeps YouTube on the regular startup path while preserving earlier overlay availability.

Additional verification: bun test src/main/runtime/initial-args-runtime-handler.test.ts src/main/runtime/initial-args-handler.test.ts src/main/runtime/initial-args-main-deps.test.ts; bun run test:runtime:compat

Follow-up fix: the YouTube subtitle picker now ignores immediate Enter key bleed-through right after opening, preventing the startup keypress from auto-submitting the default track selection before the modal is visible.

Added renderer regression coverage for immediate Enter suppression and verified with bun test src/renderer/modals/youtube-track-picker.test.ts plus the runtime-compat verification lane for the touched files.

Follow-up fix: YouTube subtitle picker startup now uses a dedicated modal-window path consistently instead of mixing main-overlay first-open with modal-window retry. That prevents late overlay opens from covering the interactive picker while mpv/window tracking settles.

Verified with bun test src/main/runtime/youtube-picker-open.test.ts, bun test src/renderer/modals/youtube-track-picker.test.ts, and the runtime-compat verification lane for src/main.ts plus the touched picker files.
<!-- SECTION:FINAL_SUMMARY:END -->
