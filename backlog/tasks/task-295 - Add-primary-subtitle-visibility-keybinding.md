---
id: TASK-295
title: Add primary subtitle visibility keybinding
status: Done
assignee:
  - Codex
created_date: '2026-04-25 23:09'
updated_date: '2026-04-25 23:45'
labels:
  - renderer
  - keybindings
  - subtitles
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Add a `v` keybinding that overrides mpv's default `v` subtitle visibility toggle and instead toggles SubMiner's primary subtitle bar visibility on and off. Secondary subtitle hover behavior is out of scope.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 Pressing `v` toggles the primary subtitle bar from visible to hidden.
- [x] #2 Pressing `v` again restores the primary subtitle bar visibility.
- [x] #3 The keybinding does not add or change secondary subtitle hover behavior.
- [x] #4 Relevant automated coverage verifies the toggle behavior.
- [x] #5 Pressing `v` in the mpv/plugin keybinding path also toggles the primary subtitle bar visibility instead of mpv native subtitle visibility.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect existing renderer keybinding and subtitle bar visibility code, including current local edits in touched files.
2. Add a focused failing test for `v` toggling primary subtitle bar visibility without changing secondary hover behavior.
3. Implement the minimal renderer/keybinding change.
4. Run targeted tests and update acceptance criteria/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented renderer-local `KeyV` handling before session/mpv binding dispatch so mpv `sub-visibility` is not touched. Visibility state is stored in renderer state and applied via `primary-sub-hidden` class on the primary subtitle container.

Scope updated after user clarified the toggle must work when focus is in mpv as well as in the overlay renderer. Added a forced mpv plugin binding for `v` that runs `--toggle-primary-subtitle-bar`, then broadcasts a renderer IPC toggle event and reuses the same primary subtitle bar toggle path.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Summary:
- Added a renderer-local `v` key handler that toggles primary subtitle bar visibility by adding/removing `primary-sub-hidden` on the primary subtitle container.
- Added renderer state for the toggle so repeated presses restore the bar without issuing mpv `sub-visibility` commands.
- Added a forced mpv plugin `v` binding that invokes `--toggle-primary-subtitle-bar` and broadcasts the same renderer toggle event.
- Added CSS for the hidden primary subtitle bar state and regression coverage for both overlay and mpv/plugin entry points.

Tests:
- `bun test src/renderer/handlers/keyboard.test.ts --test-name-pattern "primary subtitle visibility key"`
- `bun test src/cli/args.test.ts --test-name-pattern "session action"`
- `bun test src/core/services/cli-command.test.ts --test-name-pattern "visibility and utility"`
- `bun test src/cli/args.test.ts src/cli/help.test.ts src/core/services/cli-command.test.ts src/main/runtime/cli-command-context.test.ts src/main/runtime/cli-command-context-deps.test.ts src/main/runtime/cli-command-context-main-deps.test.ts src/main/runtime/cli-command-context-factory.test.ts src/main/runtime/composers/cli-startup-composer.test.ts src/main/runtime/first-run-setup-service.test.ts`
- `lua scripts/test-plugin-start-gate.lua && lua scripts/test-plugin-lua-compat.lua`
- `bun run typecheck`
- `bun run test:fast`
<!-- SECTION:FINAL_SUMMARY:END -->
