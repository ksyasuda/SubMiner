---
id: TASK-216
title: 'Address PR #28 CodeRabbit follow-ups on subtitle sidebar'
status: In Progress
assignee:
  - '@codex'
created_date: '2026-03-21 00:00'
updated_date: '2026-03-22 04:05'
labels:
  - pr-review
  - subtitle-sidebar
  - renderer
dependencies: []
references:
  - src/main/runtime/subtitle-prefetch-init.ts
  - src/main/runtime/subtitle-prefetch-init.test.ts
  - src/renderer/handlers/mouse.ts
  - src/renderer/handlers/mouse.test.ts
  - src/renderer/modals/subtitle-sidebar.ts
  - src/renderer/modals/subtitle-sidebar.test.ts
  - src/renderer/style.css
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Validate the CodeRabbit follow-ups on PR #28 for the subtitle sidebar workstream, implement the confirmed fixes, and verify the touched runtime and renderer paths.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Review comments that described real regressions are fixed in code
- [x] #2 Focused regression coverage exists for the fixed behaviors
- [x] #3 Targeted typecheck and runtime-compat verification pass
<!-- AC:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Completed follow-up fixes for PR #28:
- Cleared parsed subtitle cues on subtitle prefetch init failure so stale snapshot cache entries do not survive a failed refresh.
- Treated primary and secondary subtitle containers as one hover region so moving between them does not resume playback mid-transition.
- Kept the subtitle sidebar closed when disabled, serialized snapshot polling with timeouts, made cue rows keyboard-activatable, resolved stale cue selection fallback, and resumed hover-paused playback when the modal closes.

Regression coverage added:
- `src/main/runtime/subtitle-prefetch-init.test.ts`
- `src/renderer/handlers/mouse.test.ts`
- `src/renderer/modals/subtitle-sidebar.test.ts`

Verification:
- `bun test src/main/runtime/subtitle-prefetch-init.test.ts`
- `bun test src/renderer/handlers/mouse.test.ts`
- `bun test src/renderer/modals/subtitle-sidebar.test.ts`
- `bun run typecheck`
- `bun run test:runtime:compat`

2026-03-21: Reopened to assess a newer CodeRabbit review pass on PR #28 and address any remaining valid action items before push/reply.

2026-03-21: Addressed the latest CodeRabbit follow-up pass in commit d70c6448 after rebasing onto the updated remote branch tip.

2026-03-21: Reopened for the latest CodeRabbit round on commit d70c6448; current actionable item is the invalid ctx.state.isOverSubtitleSidebar assignment in subtitle-sidebar.ts.

2026-03-22: Addressed the live hover-state and startup mouse-ignore follow-ups from the latest CodeRabbit pass. `handleMouseLeave()` now clears `isOverSubtitle` and drops `secondary-sub-hover-active` when leaving the secondary subtitle container toward the primary container, and renderer startup now calls `syncOverlayMouseIgnoreState(ctx)` instead of forcing `setIgnoreMouseEvents(true, { forward: true })`. The sidebar IPC hover catch and CSS spacing comments were already satisfied in the current tree.

2026-03-22: Regenerated `bun.lock` from a clean install so the `electron-builder-squirrel-windows` override now resolves at `26.8.2` in the lockfile alongside `app-builder-lib@26.8.2`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Implemented the confirmed PR #28 CodeRabbit follow-ups for subtitle sidebar behavior and added regression coverage plus verification for the touched renderer and runtime paths.

Handled the latest CodeRabbit review pass for PR #28: accepted zero sidebar opacity, closed/inerted the sidebar when refresh sees config disabled, moved poll rescheduling out of finally, caught hover pause IPC failures, and fixed the stylelint spacing issue.

Verification: bun test src/config/resolve/subtitle-sidebar.test.ts; bun test src/renderer/modals/subtitle-sidebar.test.ts; bun test src/renderer/handlers/mouse.test.ts; bun run typecheck; bun run test:fast; bun run test:env; bun run build; SubMiner verifier lanes config + runtime-compat (including test:runtime:compat and test:smoke:dist).
<!-- SECTION:FINAL_SUMMARY:END -->
