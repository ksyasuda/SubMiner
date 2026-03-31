---
id: TASK-260
title: >-
  Fix macOS overlay subtitle sidebar passthrough without requiring a subtitle
  hover cycle
status: Done
assignee:
  - '@codex'
created_date: '2026-03-31 00:58'
updated_date: '2026-03-31 19:37'
labels:
  - bug
  - macos
  - overlay
  - subtitle-sidebar
  - passthrough
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/overlay-mouse-ignore.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/handlers/mouse.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/main/overlay-runtime.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/core/services/overlay-visibility.ts
documentation:
  - docs/workflow/verification.md
priority: high
ordinal: 156500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
On macOS, opening the overlay-layout subtitle sidebar should allow click-through outside the sidebar immediately. Users should not need to first hover subtitle content before passthrough/click-through starts working, including when no subtitle line is currently visible.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 With the overlay-layout subtitle sidebar open on macOS, areas outside the sidebar pass clicks through immediately after open without requiring a prior subtitle hover.
- [x] #2 When no subtitle line is currently visible, opening the subtitle sidebar still leaves non-sidebar overlay regions click-through on macOS.
- [x] #3 Regression coverage exercises the first-open/idle passthrough path so overlay interactivity does not depend on a later hover cycle.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add/adjust focused overlay visibility regressions for the tracked macOS visible overlay so the default idle state stays click-through instead of forcing mouse interaction.
2. Update main-process visible overlay visibility sync to keep the tracked macOS overlay passive by default and let renderer hover/sidebar state opt into interaction.
3. Run focused verification for overlay visibility and any dependent runtime tests, then update task notes/criteria/final summary with the confirmed outcome.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Investigation points to a main-process override on macOS: renderer sidebar open path already requests mouse passthrough outside the panel, but visible-overlay visibility sync still hard-sets the tracked overlay window interactive on macOS (`mouse-ignore:false`). Window-tracker focus/visibility resync can therefore undo renderer passthrough until a later hover cycle re-applies it.

Added a failing regression in `src/core/services/overlay-visibility.test.ts` showing the tracked macOS visible overlay was still forced interactive by main-process visibility sync (`mouse-ignore:false`) instead of staying forwarded click-through.

Updated `src/core/services/overlay-visibility.ts` so tracked macOS visible overlays now default to `setIgnoreMouseEvents(true, { forward: true })`, matching the renderer-side passthrough model and preventing window-tracker/focus resync from undoing idle sidebar clickthrough.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed the macOS subtitle-sidebar passthrough regression by changing tracked visible-overlay startup/visibility sync to stay click-through by default in the main process. Previously `updateVisibleOverlayVisibility` forced the macOS overlay window interactive, which could override renderer sidebar passthrough until a later hover cycle repaired it. Added a regression in `src/core/services/overlay-visibility.test.ts` and verified with `bun test src/core/services/overlay-visibility.test.ts`, `bun test src/renderer/modals/subtitle-sidebar.test.ts src/renderer/handlers/mouse.test.ts`, and `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
