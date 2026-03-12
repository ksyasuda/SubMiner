---
id: TASK-159
title: Add overlay controller support for keyboard-only mode
status: Done
assignee:
  - codex
created_date: '2026-03-11 00:30'
updated_date: '2026-03-11 04:05'
labels:
  - enhancement
  - renderer
  - overlay
  - input
dependencies:
  - TASK-86
references:
  - src/renderer/handlers/keyboard.ts
  - src/renderer/renderer.ts
  - src/renderer/state.ts
  - src/renderer/index.html
  - src/renderer/style.css
  - src/preload.ts
  - src/types.ts
  - src/config/definitions/defaults-core.ts
  - src/config/definitions/options-core.ts
  - src/config/definitions/template-sections.ts
  - config.example.jsonc
priority: medium
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add Chrome Gamepad API support to the visible overlay as a supplement to keyboard-only mode. By default SubMiner should bind to the first available controller, allow the user to pick and persist a preferred controller, expose a raw-input debug modal, and map controller actions onto the existing keyboard-only/Yomitan flow without breaking keyboard input. Also fix the current keyboard-only cleanup bug so the selected-token highlight clears when keyboard-only mode turns off or when the Yomitan popup closes.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Controller input is ignored unless keyboard-only mode is enabled, except the controller binding for toggling keyboard-only mode itself.
- [x] #2 Default logical mappings work: smooth popup scroll, token selection, lookup toggle/close, mining, Yomitan audio navigation/play, and mpv play/pause.
- [x] #3 Controller config supports named logical bindings plus tuning knobs (preferred controller, deadzones, smooth-scroll speed/repeat), not raw axis/button maps.
- [x] #4 `Alt+C` opens a controller selection modal listing connected controllers; saving a choice persists the preferred controller for next launch.
- [x] #5 `Alt+Shift+C` opens a debug modal showing live raw controller axes/buttons as seen by SubMiner.
- [x] #6 Keyboard-only selection highlight clears immediately when keyboard-only mode is disabled or the Yomitan popup closes.
- [x] #7 Renderer/config regression tests cover controller gating, mappings, modal behavior, persisted selection, and highlight cleanup.
- [x] #8 Docs/config example describe the controller feature and new shortcuts.

<!-- AC:END -->

## Implementation Notes

- Added renderer-side gamepad polling and logical action mapping in `src/renderer/handlers/gamepad-controller.ts`.
- Added controller select/debug modals, persisted preferred-controller IPC, and top-level `controller` config defaults/schema/template output.
- Added a transient in-overlay controller status indicator when a controller is first detected.
- Tuned controller defaults and routing after live testing: d-pad fallback navigation, slower repeat timing, DOM-backed popup-open detection, and direct pixel scroll/audio-source popup bridge commands.
- Reused existing keyboard-only lookup/mining/navigation flows so controller input stays a supplement to keyboard-only mode instead of a parallel input path.
- Verified keyboard-only highlight cleanup on mode-off and popup-close paths with renderer tests.

## Verification

- `bun test src/config/config.test.ts src/config/definitions/domain-registry.test.ts src/renderer/handlers/keyboard.test.ts src/renderer/handlers/gamepad-controller.test.ts src/renderer/modals/controller-select.test.ts src/renderer/modals/controller-debug.test.ts src/core/services/ipc.test.ts`
- `bun test src/main/runtime/composers/ipc-runtime-composer.test.ts`
- `bun run generate:config-example`
- `bun run typecheck`
- `bun run docs:test`
- `bun run test:fast`
- `bun run test:env`
- `bun run build`
- `bun run docs:build`
- `bun run test:smoke:dist`
