---
id: TASK-303
title: Update tray menu help action
status: Done
assignee:
  - Codex
created_date: '2026-04-26 03:54'
updated_date: '2026-04-26 04:12'
labels:
  - tray
  - overlay
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the tray menu's direct visible-overlay open action with an action that opens the existing in-session help modal. The tray should no longer expose an "Open Overlay" menu item; users should be able to open help from the tray instead.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Tray menu no longer includes an "Open Overlay" option.
- [x] #2 Tray menu includes an option to open the session help modal.
- [x] #3 Selecting the new tray help option initializes overlay runtime if needed and invokes the existing session help modal path.
- [x] #4 Focused regression tests cover the menu label and action wiring.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add focused regression coverage for tray menu template labels and main-process action wiring: assert Open Overlay is absent, Open Help is present, and clicking help initializes overlay runtime if needed before opening the existing session help modal path.
2. Update tray runtime action types/template to replace openOverlay with openSessionHelp.
3. Update tray main action builder dependencies to call the existing openSessionHelpModal function after overlay runtime initialization.
4. Run targeted tray tests, then broader relevant fast tests if needed.
5. Check acceptance criteria and finalize backlog notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented tray menu replacement via existing session help overlay path. Verification passed: targeted tray tests (`bun test src/main/runtime/tray-runtime.test.ts src/main/runtime/tray-main-actions.test.ts src/main/runtime/tray-main-deps.test.ts src/main/runtime/tray-runtime-handlers.test.ts`), SubMiner verifier lanes `runtime-compat` and `docs`, and `bun run changelog:lint`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Replaced the tray menu's `Open Overlay` item with `Open Help`, wired it to initialize overlay runtime when needed, and then open the existing session help modal path. Updated tray runtime/main-deps/action tests to assert the old label is absent, the new label is present, and the new action calls the help modal. Added changelog fragment `changes/303-tray-help-menu.md`.

Verification:
- `bun test src/main/runtime/tray-runtime.test.ts src/main/runtime/tray-main-actions.test.ts src/main/runtime/tray-main-deps.test.ts src/main/runtime/tray-runtime-handlers.test.ts`
- `bash plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh --lane runtime-compat --lane docs src/main/runtime/tray-runtime.ts src/main/runtime/tray-main-actions.ts src/main/runtime/tray-main-deps.ts src/main.ts changes/303-tray-help-menu.md`
- `bun run changelog:lint`

Verifier artifacts: `.tmp/skill-verification/subminer-verify-20260425-211156-9fkdDf/`.
<!-- SECTION:FINAL_SUMMARY:END -->
