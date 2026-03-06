---
id: TASK-86
title: 'Renderer: keyboard-driven Yomitan lookup mode and popup key forwarding'
status: Done
assignee:
  - Codex
created_date: '2026-03-04 13:40'
updated_date: '2026-03-05 11:30'
labels:
  - enhancement
  - renderer
  - yomitan
dependencies:
  - TASK-77
references:
  - src/renderer/handlers/keyboard.ts
  - src/renderer/handlers/mouse.ts
  - src/renderer/renderer.ts
  - src/renderer/state.ts
  - src/renderer/yomitan-popup.ts
  - src/core/services/overlay-window.ts
  - src/preload.ts
  - src/shared/ipc/contracts.ts
  - src/types.ts
  - vendor/yomitan/js/app/frontend.js
  - vendor/yomitan/js/app/popup.js
  - vendor/yomitan/js/display/display.js
  - vendor/yomitan/js/display/popup-main.js
  - vendor/yomitan/js/display/display-audio.js
documentation:
  - README.md
  - docs/usage.md
  - docs/shortcuts.md
priority: medium
ordinal: 13000
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Add true keyboard-driven token lookup flow in overlay:

- Toggle keyboard token-selection mode and navigate tokens by keyboard (`Arrow` + `HJKL`).
- Toggle Yomitan lookup window for selected token via fixed accelerator (`Ctrl/Cmd+Y`) without requiring mouse click.
- Preserve keyboard-only workflow while popup is open by forwarding popup keys (`J/K`, `M`, `P`, `[`, `]`) and restoring overlay focus on popup close.
- Ensure selection styling and hover metadata tooltips (frequency/JLPT) work for keyboard-selected token.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Keyboard mode toggle exists and shows visual selection outline for active token.
- [x] #2 Navigation works via arrows and vim keys while keyboard mode is enabled.
- [x] #3 Lookup window toggles from selected token with `Ctrl/Cmd+Y`; close path restores overlay keyboard focus.
- [x] #4 Popup-local controls work via keyboard forwarding (`J/K`, `M`, `P`, `[`, `]`), including mine action.
- [x] #5 Frequency/JLPT hover tags render for keyboard-selected token.
- [x] #6 Renderer/runtime tests cover new visibility/selection behavior, and docs are updated.
<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented keyboard-driven Yomitan workflow end-to-end in renderer + bundled Yomitan runtime bridge. Added overlay-level keyboard mode state, token selection sync, lookup toggle routing, popup command forwarding, and focus recovery after popup close. Follow-up fixes kept lookup open while moving between tokens, made popup-local `J/K` and `ArrowUp/ArrowDown` scroll work from overlay-owned focus with key repeat, skipped keyboard/token annotation flow for parser groups that have no dictionary-backed headword, and preserved paused playback when token navigation jumps across subtitle lines. Updated user docs/README to document the final shortcut behavior.

<!-- SECTION:FINAL_SUMMARY:END -->
