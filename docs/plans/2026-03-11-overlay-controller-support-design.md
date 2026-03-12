# Overlay Controller Support Design

**Date:** 2026-03-11
**Backlog:** `TASK-159`

## Goal

Add controller support to the visible overlay through the Chrome Gamepad API without replacing the existing keyboard-only workflow. Controller input should only supplement keyboard-only mode, preserve existing behavior, and expose controller selection plus raw-input debugging in overlay-local modals.

## Scope

- Poll connected gamepads from the visible overlay renderer.
- Default to the first connected controller unless config specifies a preferred controller.
- Add logical controller bindings and tuning knobs to config.
- Add `Alt+C` controller selection modal.
- Add `Alt+Shift+C` controller debug modal.
- Map controller actions onto existing keyboard-only/Yomitan behaviors.
- Fix stale selected-token highlight cleanup when keyboard-only mode turns off or popup closes.

Out of scope for this pass:

- Raw arbitrary axis/button index remapping in config.
- Controller support outside the visible overlay renderer.
- Haptics or vibration.

## Architecture

Use a renderer-local controller runtime. The overlay already owns keyboard-only token selection, Yomitan popup integration, and modal UX, and the Gamepad API is browser-native. A renderer module can poll `navigator.getGamepads()` on animation frames, normalize sticks/buttons into logical actions, and call the same helpers used by keyboard-only mode.

Avoid synthetic keyboard events as the primary implementation. Analog sticks need deadzones, continuous smooth scrolling, and per-action repeat behavior that do not fit cleanly into key event emulation. Direct logical actions keep tests clear and make the debug modal show the exact values the runtime uses.

## Behavior

Controller actions are active only while keyboard-only mode is enabled, except the controller action that toggles keyboard-only mode can always fire so the user can enter the mode from the controller.

Default logical mappings:

- left stick vertical: smooth Yomitan popup/window scroll when popup is open
- left stick horizontal: move token selection left/right
- right stick vertical: smooth Yomitan popup/window scroll
- right stick horizontal: jump horizontally inside Yomitan popup/window
- `A`: toggle lookup
- `B`: close lookup
- `Y`: toggle keyboard-only mode
- `X`: mine card
- `L1` / `R1`: previous / next Yomitan audio
- `R2`: activate current Yomitan audio button
- `L2`: toggle mpv play/pause

Selection-highlight cleanup:

- disabling keyboard-only mode clears the selected token class immediately
- closing the Yomitan popup also clears the selected token class if keyboard-only mode is no longer active
- helper ownership should live in the shared keyboard-only selection sync path so keyboard and controller exits stay consistent

## Config

Add a top-level `controller` block in resolved config with:

- `enabled`
- `preferredGamepadId`
- `preferredGamepadLabel`
- `smoothScroll`
- `scrollPixelsPerSecond`
- `horizontalJumpPixels`
- `stickDeadzone`
- `triggerDeadzone`
- `repeatDelayMs`
- `repeatIntervalMs`
- `bindings` logical fields for the named actions/sticks

Persist the preferred controller by stable browser-exposed `id` when possible, with label stored as a diagnostic/display fallback.

## UI

Controller selection modal:

- overlay-hosted modal in the visible renderer
- lists currently connected controllers
- highlights current active choice
- selecting one persists config and makes it the active controller immediately if connected

Controller debug modal:

- overlay-hosted modal
- shows selected controller and all connected controllers
- live raw axis array values
- live raw button values, pressed flags, and touched flags if available

## Testing

Test first:

- controller gating outside keyboard-only mode
- logical mapping to existing helpers
- continuous stick scroll and repeat behavior
- modal open shortcuts
- preferred-controller selection persistence
- highlight cleanup on keyboard-only disable and popup close
- config defaults/parse/template generation coverage

## Risks

- Browser gamepad identity strings can differ across OS/browser/runtime versions.
  Mitigation: match by exact preferred id first; fall back to first connected controller.
- Continuous stick input can spam actions.
  Mitigation: deadzones plus repeat throttling and frame-time-based smooth scroll.
- Popup DOM/audio controls may vary.
  Mitigation: target stable Yomitan popup/document selectors and cover with focused renderer tests.

