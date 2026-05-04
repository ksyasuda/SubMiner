---
id: TASK-325
title: Fix keyboard-only Yomitan popup shortcut precedence
status: Done
assignee:
  - codex
created_date: '2026-05-04 01:19'
updated_date: '2026-05-04 01:22'
labels:
  - bug
  - keyboard
  - yomitan
dependencies: []
priority: high
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
When keyboard-only mode is active and a Yomitan popup is visible, popup keyboard controls must win over overlay/mpv/session keybindings. Currently default overlay bindings such as bare `j` can fire instead of scrolling/navigating the Yomitan popup.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 With keyboard-only mode active and a Yomitan popup visible, pressing `j`/`k` forwards to the Yomitan popup instead of dispatching default session bindings such as primary subtitle track cycling.
- [x] #2 With keyboard-only mode inactive, existing popup-visible session binding behavior remains unchanged for bound keys.
- [x] #3 Regression coverage captures the keyboard-only popup precedence behavior.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add a focused regression in `src/renderer/handlers/keyboard.test.ts`: keyboard-only mode + visible Yomitan popup + bare `KeyJ` session binding should forward `KeyJ` to the popup and not dispatch the mpv/session binding.
2. Verify the new test fails before production changes.
3. Patch `src/renderer/handlers/keyboard.ts` so popup key handling ignores session-binding precedence only while keyboard-driven mode is enabled.
4. Run targeted renderer keyboard tests, then update acceptance criteria and final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implementation: `handleYomitanPopupKeybind` now only lets configured session bindings take precedence when keyboard-driven mode is not enabled. In keyboard-only mode with a visible Yomitan popup, bare popup keys such as `KeyJ` forward to the popup instead of dispatching overlay/mpv keybindings. Added a regression covering `KeyJ` bound to `cycle sid`.

Verification: targeted test failed before the production change, then passed after the fix. Full local gates run: `bun test src/renderer/handlers/keyboard.test.ts --test-name-pattern "keyboard mode: popup keybinds take precedence"`, `bun test src/renderer/handlers/keyboard.test.ts`, `bun run changelog:lint`, `bun run typecheck`, `bun run test:fast`, `bun run test:env`, `bun run build`, `bun run test:smoke:dist`. Build initially required `bun install --frozen-lockfile`, submodule init, and `stats/` locked deps install because this worktree had no dependencies/submodules checked out.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Fixed keyboard-only Yomitan popup shortcut precedence by allowing popup key forwarding to bypass configured session bindings only while keyboard-driven mode is active. This makes popup controls such as `j`/`k` win over default overlay/mpv bindings like primary subtitle track cycling, while preserving existing non-keyboard-only popup behavior where configured bindings still fire.

Added renderer keyboard regression coverage for the reported `KeyJ`/`cycle sid` conflict and added a changelog fragment for the user-visible overlay fix.

Verification passed: targeted red/green regression, full renderer keyboard test file, changelog lint, typecheck, `test:fast`, `test:env`, build, and dist smoke tests.
<!-- SECTION:FINAL_SUMMARY:END -->
