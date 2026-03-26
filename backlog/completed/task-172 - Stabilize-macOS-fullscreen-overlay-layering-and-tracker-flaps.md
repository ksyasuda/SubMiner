---
id: TASK-172
title: Stabilize macOS fullscreen overlay layering and tracker flaps
status: Done
assignee:
  - '@codex'
created_date: '2026-03-16 10:45'
updated_date: '2026-03-18 05:28'
labels:
  - bug
  - macos
  - overlay
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-window.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-visibility.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-runtime-init.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/window-trackers/macos-tracker.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main.ts
priority: high
ordinal: 54500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Fix the macOS fullscreen overlay bug where the visible overlay can slip behind mpv or become briefly hidden/non-interactable after tracker/helper churn. Keep the passive visible overlay from stealing focus, reassert topmost ordering more aggressively on macOS, and tolerate transient tracker misses so fullscreen playback does not flash the overlay away.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 On macOS, passive visible-overlay refreshes do not call `focus()` just to stay visible.
- [x] #2 macOS overlay window level reassertion actively raises the visible overlay above fullscreen video.
- [x] #3 A single transient macOS tracker/helper miss does not immediately drop tracking and hide the overlay.
- [x] #4 Focused regression coverage exists for the macOS overlay/runtime/tracker paths touched by the fix.
- [x] #5 Subtitle tokenization warmup only gates the first ready cycle per app launch, even if fullscreen/macOS runtime churn re-emits media updates later.
- [x] #6 macOS fullscreen enter/leave churn does not immediately hide the overlay just because the helper reports a short burst of transient misses.
- [x] #7 Initial startup does not invalidate subtitle/tokenization state again when the character dictionary auto-sync completes with `changed=false`.
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Changed `src/core/services/overlay-visibility.ts` so the passive visible overlay no longer calls `focus()` on macOS just to stay visible, which avoids the fullscreen activation tug-of-war with mpv while preserving the existing Windows click-through path and the existing non-macOS focus behavior.

Changed `src/core/services/overlay-window.ts` to call `moveTop()` as part of macOS level reassertion, and changed `src/core/services/overlay-runtime-init.ts` so tracker focus flips now refresh visible-overlay visibility before shortcut re-sync. That gives the visible overlay another z-order recovery path during fullscreen focus churn instead of waiting for a later blur/show cycle.

Changed `src/window-trackers/macos-tracker.ts` to add a small helper runner seam plus consecutive-miss tolerance. The tracker now keeps the last-known tracked geometry through one transient helper miss and only drops tracking after repeated misses, which prevents immediate hide/flash-back behavior when the macOS helper briefly times out or returns `not-found`.

Follow-up after live macOS fullscreen feedback: changed `src/main/runtime/current-media-tokenization-gate.ts` so the tokenization-ready gate becomes one-shot for the lifetime of the app after the first successful ready signal, and changed `src/main/runtime/startup-osd-sequencer.ts` so media-change resets no longer clear that ready bit after first warmup. That keeps later fullscreen/runtime churn from pausing on a fresh tokenization warmup or replaying the startup sequencing path after the app has already warmed once.

Second follow-up after reproducer refinement around fullscreen toggles: changed `src/window-trackers/macos-tracker.ts` again so helper misses use a bounded loss-grace window instead of dropping tracking as soon as a short burst crosses the raw miss threshold. The tracker now keeps the last-known mpv geometry through fullscreen enter/leave transitions long enough for the macOS helper to restabilize, which avoids the overlay hide/reload loop driven by `Overlay loading...` during transient fullscreen churn.

Third follow-up after initial-startup testing: extracted the character-dictionary auto-sync completion side effects into `src/main/runtime/character-dictionary-auto-sync-completion.ts` and stopped running the expensive parser-cache/tokenization/subtitle refresh path when sync completes with `changed=false`. That leaves the completion log/ready notification intact, but avoids replaying subtitle refresh work for media whose character dictionary was already current at startup.

Added focused regressions in `src/core/services/overlay-visibility.test.ts`, `src/core/services/overlay-runtime-init.test.ts`, `src/window-trackers/macos-tracker.test.ts`, `src/main/runtime/current-media-tokenization-gate.test.ts`, and `src/main/runtime/startup-osd-sequencer.test.ts`. Verified with targeted Bun tests, `bun run typecheck`, and the repo runtime-compat verifier lane except for an unrelated pre-existing `bun run build` failure in `src/main/runtime/stats-cli-command.test.ts`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Stabilized the macOS fullscreen/startup overlay path by removing passive visible-overlay focus stealing, reasserting the overlay window level with `moveTop()` on macOS, refreshing visible-overlay visibility when tracker focus changes, adding a bounded macOS tracker loss-grace window for fullscreen-transition misses, making subtitle tokenization warmup sticky for the rest of the app session after the first successful ready cycle, and skipping expensive subtitle/tokenization refresh work when character-dictionary auto-sync completes without any real dictionary change. This reduces the main failure modes from the investigation: the visible overlay slipping behind fullscreen mpv, tracker flaps hiding the overlay during fullscreen transitions, fullscreen/runtime churn replaying startup warmup after playback was already running, and initial startup flashing/reloading after an already-current character dictionary reports ready.

Verification:
- `bun test src/core/services/overlay-window.test.ts src/core/services/overlay-visibility.test.ts src/core/services/overlay-runtime-init.test.ts src/window-trackers/x11-tracker.test.ts src/window-trackers/macos-tracker.test.ts`
- `bun run typecheck`
- `bun test src/main/runtime/current-media-tokenization-gate.test.ts src/main/runtime/startup-osd-sequencer.test.ts src/main/runtime/character-dictionary-auto-sync.test.ts src/main/runtime/composers/mpv-runtime-composer.test.ts`
- `bun test src/window-trackers/macos-tracker.test.ts src/core/services/overlay-visibility.test.ts src/core/services/overlay-runtime-init.test.ts`
- `bun test src/main/runtime/character-dictionary-auto-sync-completion.test.ts src/main/runtime/character-dictionary-auto-sync.test.ts src/main/runtime/current-media-tokenization-gate.test.ts src/main/runtime/startup-osd-sequencer.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane runtime-compat src/core/services/overlay-visibility.ts src/core/services/overlay-window.ts src/core/services/overlay-runtime-init.ts src/window-trackers/macos-tracker.ts src/core/services/overlay-visibility.test.ts src/core/services/overlay-runtime-init.test.ts src/window-trackers/macos-tracker.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane runtime-compat src/main/runtime/current-media-tokenization-gate.ts src/main/runtime/startup-osd-sequencer.ts src/main/runtime/current-media-tokenization-gate.test.ts src/main/runtime/startup-osd-sequencer.test.ts` [build blocked by unrelated `src/main/runtime/stats-cli-command.test.ts` typing errors already present in workspace]
<!-- SECTION:FINAL_SUMMARY:END -->
