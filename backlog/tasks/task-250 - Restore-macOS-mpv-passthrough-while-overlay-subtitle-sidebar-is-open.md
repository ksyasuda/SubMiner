---
id: TASK-250
title: Restore macOS mpv passthrough while overlay subtitle sidebar is open
status: Done
assignee:
  - '@codex'
created_date: '2026-03-29 10:10'
updated_date: '2026-03-31 19:37'
labels:
  - bug
  - macos
  - subtitle-sidebar
  - overlay
  - mpv
dependencies: []
references:
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/overlay-mouse-ignore.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.ts
  - /Users/sudacode/projects/japanese/SubMiner/src/renderer/handlers/keyboard.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/modals/subtitle-sidebar.test.ts
  - >-
    /Users/sudacode/projects/japanese/SubMiner/src/renderer/overlay-mouse-ignore.test.ts
priority: high
ordinal: 167500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When the overlay-layout subtitle sidebar is open on macOS, users should still be able to click through outside the sidebar and return keyboard focus to mpv so native mpv keybindings continue to work. The sidebar should stay interactive when hovered or focused, but it must not make the whole visible overlay behave like a blocking modal.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Opening the overlay-layout subtitle sidebar does not keep the entire visible overlay mouse-interactive outside sidebar hover or focus.
- [x] #2 With the subtitle sidebar open, clicking outside the sidebar can refocus mpv so native mpv keybindings continue to work.
- [x] #3 Focused regression coverage exists for overlay-layout sidebar passthrough behavior on mouse-ignore state changes.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add renderer regression coverage for overlay-layout subtitle sidebar passthrough so open-but-unhovered sidebar no longer holds global mouse interaction.
2. Update overlay mouse-ignore gating to keep the subtitle sidebar interactive only while hovered or otherwise actively interacting, instead of treating overlay layout as a blocking modal.
3. Run focused renderer tests for subtitle sidebar and mouse-ignore behavior, then update task notes/criteria with the verified outcome.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Confirmed the regression only affects the default overlay-layout subtitle sidebar: open sidebar state was treated as a blocking overlay modal, which prevented click-through outside the sidebar and stranded native mpv keybindings until focus was manually recovered.

Added a failing regression in src/renderer/modals/subtitle-sidebar.test.ts for overlay-layout passthrough before changing the gate.

Verification: bun test src/renderer/modals/subtitle-sidebar.test.ts src/renderer/overlay-mouse-ignore.test.ts; bun run typecheck

User reported the first renderer-only fix did not resolve the macOS issue in practice. Reopening investigation to trace visible-overlay window focus and hit-testing outside the renderer mouse-ignore gate.

Follow-up root cause: sidebar hover handlers were attached to the full-screen `.subtitle-sidebar-modal` shell instead of the actual sidebar panel. On the transparent visible overlay that shell spans the viewport, so sidebar-active state could persist outside the panel and keep the overlay interactive longer than intended.

Updated the sidebar modal to track hover/focus on `subtitleSidebarContent` and derive sidebar interaction state from panel hover or focus-within before recomputing mouse passthrough.

Verification refresh: bun test src/renderer/modals/subtitle-sidebar.test.ts src/renderer/overlay-mouse-ignore.test.ts; bun run typecheck
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Restored overlay subtitle sidebar passthrough in two layers. First, the visible overlay mouse-ignore gate no longer treats the subtitle sidebar as a global blocking modal. Second, the sidebar panel now tracks interaction on the real sidebar content instead of the full-screen modal shell, and keeps itself active only while the panel is hovered or focused. Added regressions for overlay-layout passthrough and focus-within behavior. Verification: `bun test src/renderer/modals/subtitle-sidebar.test.ts src/renderer/overlay-mouse-ignore.test.ts` and `bun run typecheck`.
<!-- SECTION:FINAL_SUMMARY:END -->
