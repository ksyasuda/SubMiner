---
id: TASK-165
title: Make controller configuration easier with inline remapping modal
status: To Do
assignee:
  - Codex
created_date: '2026-03-13 00:10'
updated_date: '2026-03-13 00:10'
labels:
  - enhancement
  - renderer
  - overlay
  - input
  - config
dependencies:
  - TASK-159
references:
  - src/renderer/modals/controller-select.ts
  - src/renderer/modals/controller-debug.ts
  - src/renderer/handlers/gamepad-controller.ts
  - src/renderer/index.html
  - src/renderer/style.css
  - src/renderer/utils/dom.ts
  - src/preload.ts
  - src/core/services/ipc.ts
  - src/main.ts
  - src/types.ts
  - src/config/definitions/defaults-core.ts
  - src/config/definitions/options-core.ts
  - config.example.jsonc
  - docs/plans/2026-03-13-overlay-controller-config-remap-design.md
  - docs/plans/2026-03-13-overlay-controller-config-remap.md
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->
Replace the current controller-selection-only modal with a denser controller configuration surface that keeps device selection and adds inline controller remapping. The new flow should feel like emulator configuration: pick an overlay action, arm capture, then press the matching controller button, trigger, d-pad direction, or stick direction to bind it. Keep the current overlay-local renderer architecture, preserve controller gating to keyboard-only mode, and retain the separate raw debug modal for troubleshooting.
<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->
- [ ] #1 `Alt+C` opens a controller modal that includes both preferred-controller selection and controller-config editing in one surface.
- [ ] #2 Controller device selection uses a compact dropdown or equivalent compact picker instead of the current full-height device list.
- [ ] #3 Each remappable controller action shows its current binding and supports learn/capture, clear, and reset-to-default flows.
- [ ] #4 Learn mode captures the next fresh controller input edge or stick/d-pad direction, not a held/stale input.
- [ ] #5 Captured bindings can represent non-standard controllers without depending only on the browser's standard semantic button names.
- [ ] #6 Updated bindings persist through the existing config pipeline and take effect in the renderer without restart unless a field explicitly requires reopen/reload.
- [ ] #7 Existing controller behavior remains gated to keyboard-only mode except for the controller action that toggles keyboard-only mode itself.
- [ ] #8 Renderer/config/IPC regression tests cover the new modal layout, capture flow, persistence, and runtime mapping behavior.
- [ ] #9 Docs/config example explain the new controller-config flow and when to use the debug modal.
<!-- AC:END -->

## Implementation Plan

<!-- SECTION:PLAN:BEGIN -->
1. Add the design doc and implementation plan for inline controller remapping, tied to a new backlog task instead of reopening the already-completed base controller-support task.
2. Expand controller config types/defaults/template output so action bindings can store captured input descriptors, not only semantic button-name enums.
3. Extend preload/main/IPC write paths from preferred-controller-only saves to full controller-config patching needed by the modal.
4. Redesign the controller modal UI into a compact device picker plus action-binding editor with learn, clear, and reset affordances.
5. Add renderer capture state and a learn-mode runtime that waits for neutral-to-active transitions before saving a binding.
6. Update the gamepad runtime to resolve the new stored descriptors into actions while preserving current gating and repeat/deadzone behavior.
7. Keep the raw debug modal as a separate advanced surface; optionally expose copyable input-descriptor text for troubleshooting.
8. Add focused regression tests first, then run the maintained gate needed for docs/config/renderer/main changes.
<!-- SECTION:PLAN:END -->

## Implementation Notes

<!-- SECTION:NOTES:BEGIN -->
Planning only in this pass.

Current-state findings:

- `src/renderer/modals/controller-select.ts` only persists `preferredGamepadId` / `preferredGamepadLabel`.
- `src/preload.ts`, `src/core/services/ipc.ts`, and `src/main.ts` only expose a narrow save path for preferred controller, not general controller config writes.
- `src/renderer/handlers/gamepad-controller.ts` currently resolves actions from semantic button bindings plus a few axis slots; this is fine for defaults but too narrow for emulator-style learn mode on non-standard controllers.
- `src/renderer/modals/controller-debug.ts` already provides the raw input surface needed for troubleshooting and for validating capture behavior.

Recommended direction:

- keep `Alt+C` as the single controller-config entrypoint
- keep `Alt+Shift+C` as raw debug
- introduce stored input descriptors for discrete bindings so learn mode can capture buttons, triggers, d-pad directions, and stick directions directly
- defer per-controller profiles; keep one global binding set plus preferred-controller selection for this pass
<!-- SECTION:NOTES:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->
Planned follow-up work to make controller configuration materially easier than the current “pick preferred device” modal. The proposed change keeps existing controller runtime/debug foundations, but upgrades the selection modal into a compact controller-config surface with inline learn-mode remapping and persistent binding storage.

Main architectural change in scope: move from semantic-button-only binding storage toward captured input descriptors so the UI can reliably learn from buttons, triggers, d-pad directions, and stick directions on non-standard controllers.
<!-- SECTION:FINAL_SUMMARY:END -->
