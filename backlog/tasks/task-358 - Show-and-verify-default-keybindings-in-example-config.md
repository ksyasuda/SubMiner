---
id: TASK-358
title: Show and verify default keybindings in example config
status: Done
assignee:
  - '@Codex'
created_date: '2026-05-13 03:33'
updated_date: '2026-05-13 03:45'
labels:
  - config
  - keybindings
  - overlay
  - mpv
dependencies: []
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Keep the shipped example configuration, overlay runtime, and mpv plugin aligned with the built-in default keybindings. The example `keybindings` array should show the same defaults that are active by default, and focused tests should catch drift between documented defaults and actual overlay/mpv wiring.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria
<!-- AC:BEGIN -->
- [x] #1 `config.example.jsonc` and the docs-site copy show all default keybindings in the `keybindings` array.
- [x] #2 Default keybindings are registered without conflicts in the overlay session-binding path.
- [x] #3 Default keybindings are registered and dispatched correctly inside the mpv plugin.
- [x] #4 Focused regression tests cover default keybinding/config-example parity and mpv/plugin dispatch.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Inspect default keybinding definitions, config-example generation, overlay shortcut/session-binding tests, and mpv plugin binding tests.
2. Add failing tests for config-example keybinding parity and any missing default overlay/mpv wiring.
3. Update generated/example config and source wiring only where tests show drift.
4. Run focused Bun/Lua tests, regenerate examples if needed, update task AC/final notes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Implemented the config-template path by injecting `DEFAULT_KEYBINDINGS` into generated examples when the resolved config has an empty `keybindings` array, preserving runtime merge semantics. Added coverage for template parity, default binding compile/action mapping, overlay keyboard dispatch, and mpv plugin registration/dispatch. Regenerated both `config.example.jsonc` artifacts and added changelog fragment `changes/358-default-keybindings-config-example.md`.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Updated generated example configuration so `config.example.jsonc` and `docs-site/public/config.example.jsonc` now show every built-in default keybinding in the `keybindings` array instead of `[]`. The template copy now describes the array as default plus custom keybindings, while runtime default merge behavior remains unchanged.

Added regression coverage that the generated template parses back to `DEFAULT_KEYBINDINGS`, that every default binding compiles to the expected mpv command or session action, that the overlay keyboard handler dispatches all compiled defaults, and that the mpv plugin registers and invokes default mpv/session-action bindings. Also updated docs tables to include the default fullscreen binding and clarified that keybindings can target mpv commands or SubMiner session actions.

Verification passed: `bun run format:check:src`, `bun run changelog:lint`, `bun run docs:test`, `bun run docs:build`, `bun run verify:config-example`, focused config/session/renderer/plugin tests, `bun run typecheck`, `bun run test:env`, `bun run test:fast`, `bun run build`, and `bun run test:smoke:dist`.
<!-- SECTION:FINAL_SUMMARY:END -->
