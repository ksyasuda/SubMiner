---
id: TASK-217
title: Fix embedded overlay passthrough sync between subtitle and sidebar
status: Done
assignee:
  - codex
created_date: '2026-03-21 23:16'
updated_date: '2026-03-23 03:22'
labels:
  - bug
  - overlay
  - macos
dependencies: []
references:
  - src/renderer/handlers/mouse.ts
  - src/renderer/modals/subtitle-sidebar.ts
  - src/renderer/renderer.ts
documentation:
  - docs/workflow/verification.md
priority: high
ordinal: 118500
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
On macOS, when both the subtitle overlay and embedded subtitle sidebar are visible, mouse passthrough to mpv can remain stale until the user hovers the sidebar. After closing the sidebar, passthrough can likewise remain stale until the user hovers the subtitle again. Fix the overlay input-state synchronization so passthrough reflects the current hover/open state immediately instead of relying on the last hover target.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 When the embedded subtitle sidebar is open and the pointer is not over subtitle or sidebar content, the overlay returns to mouse passthrough immediately without requiring a sidebar hover cycle.
- [x] #2 When transitioning between subtitle hover and sidebar hover states on macOS embedded sidebar mode, mouse ignore state stays in sync with the currently interactive region.
- [x] #3 Closing the embedded subtitle sidebar restores the correct passthrough state based on remaining subtitle hover/modal state without requiring an additional hover.
- [x] #4 Regression tests cover the passthrough synchronization behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a shared renderer-side passthrough sync helper that derives whether the overlay should ignore mouse events from subtitle hover, embedded sidebar visibility/hover, popup visibility, and modal state.
2. Replace direct embedded-sidebar passthrough toggles in subtitle hover/sidebar handlers with calls to the shared sync helper so state is recomputed on every transition.
3. Add regression tests for macOS embedded sidebar mode covering sidebar-open idle passthrough, subtitle-to-sidebar transitions, and sidebar-close restore behavior.
4. Run targeted renderer tests for mouse/sidebar passthrough coverage, then summarize any residual risk.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Added shared renderer overlay mouse-ignore recompute so subtitle hover, embedded sidebar hover/open/close, and popup idle transitions all derive passthrough from current state instead of last hover target.

Added regression coverage for embedded sidebar idle passthrough on subtitle leave and for sidebar-close recompute behavior.

Verification: `bun run typecheck` passed; `bun test src/renderer/handlers/mouse.test.ts` passed; `bun test src/renderer/modals/subtitle-sidebar.test.ts` passed; core verification wrapper artifact at `.tmp/skill-verification/subminer-verify-20260321-162743-XhSBxw` hit an unrelated `bun run test:fast` failure in `scripts/update-aur-package.test.ts` because macOS system bash lacks `mapfile`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed stale embedded-sidebar passthrough sync on macOS by introducing a shared renderer mouse-ignore recompute path and tracking sidebar-hover state separately from subtitle hover. Subtitle hover leave, sidebar hover enter/leave, sidebar open, and sidebar close now all recompute passthrough from the current overlay state instead of waiting for a later hover event to repair it. Added regression tests covering subtitle-leave passthrough while the embedded sidebar is open but idle, plus sidebar-close restore behavior based on remaining subtitle hover state.

Tests run:
- `bun run typecheck`
- `bun test src/renderer/handlers/mouse.test.ts`
- `bun test src/renderer/modals/subtitle-sidebar.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/classify_subminer_diff.sh src/renderer/state.ts src/renderer/overlay-mouse-ignore.ts src/renderer/handlers/mouse.ts src/renderer/handlers/mouse.test.ts src/renderer/modals/subtitle-sidebar.ts src/renderer/modals/subtitle-sidebar.test.ts`
- `bash .agents/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane core src/renderer/state.ts src/renderer/overlay-mouse-ignore.ts src/renderer/handlers/mouse.ts src/renderer/handlers/mouse.test.ts src/renderer/modals/subtitle-sidebar.ts src/renderer/modals/subtitle-sidebar.test.ts` (typecheck passed; `test:fast` blocked by unrelated `scripts/update-aur-package.test.ts` failure on macOS Bash 3.2 lacking `mapfile`)

Risk: the classifier flagged this as a real-runtime candidate, so actual Electron/mpv macOS pointer behavior was not exercised in a live runtime during this turn.
<!-- SECTION:FINAL_SUMMARY:END -->
