---
id: TASK-107
title: Fix post-rebase overlay toggle regression
status: In Progress
assignee:
  - codex
created_date: '2026-02-21 23:34'
updated_date: '2026-02-21 23:45'
labels:
  - bug
  - overlay
  - regression
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
After recent rebase, toggling visible or invisible overlay opens a transparent non-interactable window. Symptoms:
- invisible subtitle mode no longer renders/works,
- visible subtitles do not show,
- overlay keybinds do not work on the created window.

Need root-cause fix so both overlay modes render and interactive/keybind behavior returns to pre-rebase behavior.
<!-- SECTION:DESCRIPTION:END -->

## Action Steps

<!-- SECTION:PLAN:BEGIN -->
1. Reproduce regression in source tests around overlay toggle/window behavior.
2. Add failing regression test for transparent/non-interactable overlay window path.
3. Identify rebase-introduced break in overlay creation/runtime toggle wiring.
4. Implement minimal fix for renderer loading + window interaction state.
5. Verify with focused tests and smoke checks.
<!-- SECTION:PLAN:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [ ] #1 Toggling visible overlay shows subtitle content again.
- [ ] #2 Toggling invisible overlay restores interactive subtitle behavior.
- [ ] #3 Overlay keybinds work after overlay toggle in both modes.
- [ ] #4 Added regression coverage for broken toggle path.
<!-- AC:END -->

## Definition of Done
<!-- DOD:BEGIN -->
- [ ] #1 Focused overlay/runtime tests pass.
- [ ] #2 No new type/build regressions introduced by fix.
- [ ] #3 Task notes include root cause and validation commands.
<!-- DOD:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
- Root-cause candidate fixed: renderer layer detection could trust preload `process.argv` over per-window query params, which can drift under shared renderer process and force wrong layer behavior.
- Changed `src/renderer/utils/platform.ts` to prioritize `window.location.search` (`?layer=visible|invisible`) and only fallback to preload layer when query is missing/invalid.
- Added regression test in `src/renderer/error-recovery.test.ts` asserting query-layer precedence.
- Added explicit `webPreferences.sandbox = false` in `src/core/services/overlay-window.ts` to keep preload Node-style API availability stable on newer Electron defaults.
- Added regression test `src/core/services/overlay-window-config.test.ts` to guard sandbox setting.
- Validation: `bun test src/core/services/overlay-window-config.test.ts src/renderer/error-recovery.test.ts`; `bun run build`.
<!-- SECTION:NOTES:END -->
