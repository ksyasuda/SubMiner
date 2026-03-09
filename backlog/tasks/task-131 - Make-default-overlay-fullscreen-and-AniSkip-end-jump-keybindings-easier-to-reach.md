---
id: TASK-131
title: Make default overlay fullscreen and AniSkip end-jump keybindings easier to reach
status: Done
assignee:
  - codex
created_date: '2026-03-09 00:00'
updated_date: '2026-03-09 00:30'
labels:
  - enhancement
  - overlay
  - mpv
  - aniskip
dependencies: []
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Make two default keyboard actions easier to hit during playback: add `f` as the built-in overlay fullscreen toggle, and make AniSkip's default intro-end jump use `Tab`.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Default overlay keybindings include `KeyF` mapped to mpv fullscreen toggle.
- [x] #2 Default AniSkip hint/button key defaults to `Tab` and the plugin registers that binding.
- [x] #3 Automated regression coverage exists for both default bindings.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->

1. Add a failing TypeScript regression proving default overlay keybindings include fullscreen on `KeyF`.
2. Add a failing Lua/plugin regression proving AniSkip defaults to `Tab`, updates the OSD hint text, and registers the expected keybinding.
3. Patch the default keybinding/config values with minimal behavior changes and keep fallback binding behavior intentional.
4. Run focused tests plus touched verification commands, then record results and a short changelog fragment.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->

Added `KeyF -> ['cycle', 'fullscreen']` to the built-in overlay keybindings in `src/config/definitions/shared.ts`.

Changed the mpv plugin AniSkip default button key from `y-k` to `TAB` in both the runtime default options and the shipped `plugin/subminer.conf`. The AniSkip OSD hint now also falls back to `TAB` when no explicit key is configured.

Adjusted `plugin/subminer/ui.lua` fallback registration so the legacy `y-k` binding is only added for custom non-default AniSkip bindings, instead of always shadowing the new default.

Extended regression coverage:

- `src/config/definitions/domain-registry.test.ts` now asserts the default fullscreen binding on `KeyF`.
- `scripts/test-plugin-start-gate.lua` now isolates plugin runs correctly, records keybinding/observer registration, and asserts the default AniSkip keybinding/prompt behavior for `TAB`.

Verification:

- `bun test src/config/definitions/domain-registry.test.ts`
- `bun run test:config:src`
- `lua scripts/test-plugin-start-gate.lua`
- `bun run changelog:lint`
- `bun run typecheck`

Known unrelated verification gap:

- `bun run test:plugin:src` still fails in `scripts/test-plugin-binary-windows.lua` on this Linux host (`windows env override should resolve .exe suffix`), outside the keybinding changes in this task.
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Default overlay playback now has an easier fullscreen toggle on `f`, and AniSkip's default intro-end jump now uses `Tab`. The mpv plugin hint text and registration logic were updated to match the new default, while keeping legacy `y-k` fallback behavior limited to custom non-default bindings.

Regression coverage was added for both defaults, and the plugin test harness now resets plugin bootstrap state between scenarios so keybinding assertions can run reliably.

<!-- SECTION:FINAL_SUMMARY:END -->
