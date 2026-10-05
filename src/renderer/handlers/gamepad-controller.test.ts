import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import type { ControllerRuntimeSnapshot, ResolvedControllerConfig } from '../../types';
import { createGamepadController } from './gamepad-controller.js';

type ControllerOptions = Parameters<typeof createGamepadController>[0];
type Bindings = ResolvedControllerConfig['bindings'];
type TestButton = { value: number; pressed?: boolean; touched?: boolean };

type TestGamepad = {
  id: string;
  index: number;
  connected: boolean;
  mapping: string;
  axes: number[];
  buttons: TestButton[];
};

const PRESSED: TestButton = { value: 1, pressed: true, touched: true };
const RELEASED: TestButton = { value: 0, pressed: false, touched: false };

/** Builds a 16-button standard gamepad; `buttons` only lists the non-idle buttons. */
function createGamepad(
  id: string,
  options: { index?: number; axes?: number[]; buttons?: Record<number, TestButton> } = {},
): TestGamepad {
  const buttons = Array.from({ length: 16 }, () => ({ ...RELEASED }));
  for (const [index, button] of Object.entries(options.buttons ?? {})) {
    buttons[Number(index)] = button;
  }
  return {
    id,
    index: options.index ?? 0,
    connected: true,
    mapping: 'standard',
    axes: options.axes ?? [0, 0, 0, 0],
    buttons,
  };
}

/** Shipped controller defaults with support enabled, plus per-test overrides. */
function createControllerConfig(
  overrides: Omit<Partial<ResolvedControllerConfig>, 'bindings' | 'buttonIndices'> & {
    bindings?: Partial<Bindings>;
    buttonIndices?: Partial<ResolvedControllerConfig['buttonIndices']>;
  } = {},
): ResolvedControllerConfig {
  const { bindings, buttonIndices, ...rest } = overrides;
  const defaults = CORE_DEFAULT_CONFIG.controller;
  return {
    ...defaults,
    enabled: true,
    ...rest,
    buttonIndices: { ...defaults.buttonIndices, ...buttonIndices },
    bindings: { ...defaults.bindings, ...bindings },
  };
}

const VOID_ACTIONS = [
  'toggleKeyboardMode',
  'toggleLookup',
  'closeLookup',
  'mineCard',
  'quitMpv',
  'previousAudio',
  'nextAudio',
  'playCurrentAudio',
  'toggleMpvPause',
] as const;
const DELTA_ACTIONS = ['moveSelection', 'scrollPopup', 'jumpPopup'] as const;

type HarnessCall = (typeof VOID_ACTIONS)[number] | `${(typeof DELTA_ACTIONS)[number]}:${number}`;

/**
 * Wires a controller where every action is recorded into one `calls` log
 * (deltas are rounded, e.g. `scrollPopup:-67`). Keyboard mode defaults on,
 * the lookup window closed, and interaction unblocked. Action overrides in
 * `deps` run after the call is logged; getter overrides replace the default.
 */
function createHarness(
  options: {
    config?: ResolvedControllerConfig;
    gamepads?: Array<TestGamepad | null> | (() => Array<TestGamepad | null>);
    deps?: Partial<ControllerOptions>;
  } = {},
) {
  const calls: HarnessCall[] = [];
  const states: ControllerRuntimeSnapshot[] = [];
  const deps = options.deps ?? {};
  const { gamepads = [], config = createControllerConfig() } = options;

  const controllerOptions: ControllerOptions = {
    getGamepads: typeof gamepads === 'function' ? gamepads : () => gamepads,
    getConfig: () => config,
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: () => {},
    jumpPopup: () => {},
    ...deps,
    onState: (state) => {
      states.push(state);
      deps.onState?.(state);
    },
  };
  for (const action of VOID_ACTIONS) {
    controllerOptions[action] = () => {
      calls.push(action);
      deps[action]?.();
    };
  }
  controllerOptions.moveSelection = (delta) => {
    calls.push(`moveSelection:${delta}`);
    deps.moveSelection?.(delta);
  };
  controllerOptions.scrollPopup = (delta) => {
    calls.push(`scrollPopup:${Math.round(delta)}`);
    deps.scrollPopup?.(delta);
  };
  controllerOptions.jumpPopup = (delta) => {
    calls.push(`jumpPopup:${delta}`);
    deps.jumpPopup?.(delta);
  };

  const controller = createGamepadController(controllerOptions);
  return {
    calls,
    states,
    controller,
    poll: (...times: number[]) => {
      for (const now of times) controller.poll(now);
    },
  };
}

test('gamepad controller selects the first connected controller by default', () => {
  const harness = createHarness({
    gamepads: [null, createGamepad('pad-2', { index: 1 }), createGamepad('pad-3', { index: 2 })],
  });

  harness.poll(0);

  assert.equal(harness.controller.getActiveGamepadId(), 'pad-2');
  assert.equal(harness.states.at(-1)?.activeGamepadId, 'pad-2');
});

test('gamepad controller prefers saved controller id when connected', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1'), createGamepad('pad-2', { index: 1 })],
    config: createControllerConfig({ preferredGamepadId: 'pad-2' }),
  });

  harness.poll(0);

  assert.equal(harness.controller.getActiveGamepadId(), 'pad-2');
});

test('gamepad controller allows keyboard-mode toggle while other actions stay gated', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 0: PRESSED, 3: PRESSED } })],
    deps: { getKeyboardModeEnabled: () => false },
  });

  harness.poll(0);

  assert.deepEqual(harness.calls, ['toggleKeyboardMode']);
});

test('gamepad controller re-evaluates interaction gating after toggling keyboard mode', () => {
  let keyboardModeEnabled = true;
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 0: PRESSED, 3: PRESSED } })],
    deps: {
      getKeyboardModeEnabled: () => keyboardModeEnabled,
      toggleKeyboardMode: () => {
        keyboardModeEnabled = false;
      },
    },
  });

  harness.poll(0);

  assert.deepEqual(harness.calls, ['toggleKeyboardMode']);
});

test('gamepad controller resets edge state when active controller changes', () => {
  let gamepads = [createGamepad('pad-1', { buttons: { 0: PRESSED } })];
  const harness = createHarness({ gamepads: () => gamepads });

  harness.poll(0);
  gamepads = [createGamepad('pad-2', { buttons: { 0: PRESSED } })];
  harness.poll(50);

  assert.deepEqual(harness.calls, ['toggleLookup', 'toggleLookup']);
});

test('gamepad controller does not toggle keyboard mode when controller support is disabled', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 3: PRESSED } })],
    config: createControllerConfig({ enabled: false }),
    deps: { getKeyboardModeEnabled: () => false },
  });

  harness.poll(0);

  assert.deepEqual(harness.calls, []);
});

test('gamepad controller does not treat blocked held inputs as fresh edges when interaction resumes', () => {
  let held = true;
  let interactionBlocked = true;
  const harness = createHarness({
    gamepads: () => [
      createGamepad('pad-1', {
        buttons: { 0: held ? PRESSED : RELEASED },
        axes: [held ? 0.9 : 0, 0, 0, 0],
      }),
    ],
    deps: { getInteractionBlocked: () => interactionBlocked },
  });

  harness.poll(0);
  interactionBlocked = false;
  harness.poll(100);

  assert.deepEqual(harness.calls, []);

  held = false;
  harness.poll(200);
  held = true;
  harness.poll(300);

  assert.deepEqual(harness.calls, ['toggleLookup', 'moveSelection:1']);
});

test('gamepad controller maps left stick horizontal movement to token selection repeats', () => {
  let axes = [0.9, 0, 0, 0];
  const harness = createHarness({ gamepads: () => [createGamepad('pad-1', { axes })] });

  // Default repeat delay is 320ms, then 120ms between repeats.
  harness.poll(0, 100, 260);
  assert.deepEqual(harness.calls, ['moveSelection:1']);

  harness.poll(340);
  assert.deepEqual(harness.calls, ['moveSelection:1', 'moveSelection:1']);

  axes = [0, 0, 0, 0];
  harness.poll(360);
  axes = [-0.9, 0, 0, 0];
  harness.poll(380);

  assert.deepEqual(harness.calls, ['moveSelection:1', 'moveSelection:1', 'moveSelection:-1']);
});

test('gamepad controller uses active controller profile bindings before global bindings', () => {
  const globalConfig = createControllerConfig({
    bindings: { toggleLookup: { kind: 'button', buttonIndex: 0 } },
  });
  const harness = createHarness({
    gamepads: [createGamepad('pad-profile', { buttons: { 11: PRESSED } })],
    config: {
      ...globalConfig,
      profiles: {
        'pad-profile': {
          label: 'Profile Pad',
          buttonIndices: globalConfig.buttonIndices,
          bindings: { ...globalConfig.bindings, toggleLookup: { kind: 'button', buttonIndex: 11 } },
        },
      },
    },
  });

  harness.poll(0);

  assert.deepEqual(harness.calls, ['toggleLookup']);
});

test('gamepad controller maps L1 play-current, R1 next-audio, and popup navigation', () => {
  const harness = createHarness({
    gamepads: [
      createGamepad('pad-1', {
        // Left stick up scrolls, right stick down jumps.
        axes: [0, -0.75, 0.1, 0, 0.8],
        buttons: {
          4: PRESSED,
          5: PRESSED,
          6: { value: 0.8, pressed: true, touched: true },
          7: { value: 0.9, pressed: true, touched: true },
          8: PRESSED,
        },
      }),
    ],
    config: createControllerConfig({
      bindings: {
        playCurrentAudio: { kind: 'button', buttonIndex: 4 },
        nextAudio: { kind: 'button', buttonIndex: 5 },
        previousAudio: { kind: 'none' },
        // Shares raw button 6 with the default quitMpv binding.
        toggleMpvPause: { kind: 'button', buttonIndex: 6 },
      },
    }),
    deps: { getLookupWindowOpen: () => true },
  });

  harness.poll(0, 100);

  assert.deepEqual(
    [...harness.calls].sort(),
    [
      'jumpPopup:160',
      'nextAudio',
      'playCurrentAudio',
      'quitMpv',
      'scrollPopup:-67',
      'toggleMpvPause',
    ].sort(),
  );
});

for (const c of [
  { name: 'raw button 6 (default select)', buttonIndex: 6 },
  { name: 'configured raw button 11', buttonIndex: 11 },
]) {
  test(`gamepad controller maps quit mpv from ${c.name}`, () => {
    const harness = createHarness({
      gamepads: [createGamepad('pad-1', { buttons: { [c.buttonIndex]: PRESSED } })],
      config: createControllerConfig({
        bindings: { quitMpv: { kind: 'button', buttonIndex: c.buttonIndex } },
      }),
    });

    harness.poll(0);

    assert.deepEqual(harness.calls, ['quitMpv']);
  });
}

test('gamepad controller maps right stick vertical to popup jump and ignores horizontal movement', () => {
  let axes = [0, 0, 0.85, 0, 0];
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { axes })],
    deps: { getLookupWindowOpen: () => true },
  });

  harness.poll(0, 100);
  assert.deepEqual(harness.calls, []);

  axes = [0, 0, 0.85, 0, -0.85];
  harness.poll(200);

  assert.deepEqual(harness.calls, ['jumpPopup:-160']);
});

for (const c of [
  {
    name: 'd-pad right/up buttons',
    pad: createGamepad('pad-1', {
      buttons: {
        15: { value: 1, pressed: false, touched: true },
        12: { value: 1, pressed: false, touched: true },
      },
    }),
  },
  { name: 'd-pad axes 6 and 7', pad: createGamepad('pad-1', { axes: [0, 0, 0, 0, 0, 0, 1, -1] }) },
]) {
  test(`gamepad controller maps ${c.name} to selection and popup scroll`, () => {
    const harness = createHarness({
      gamepads: [c.pad],
      deps: { getLookupWindowOpen: () => true },
    });

    harness.poll(0, 100);

    // Scroll is time-based, so it only fires once 100ms have elapsed.
    assert.deepEqual(harness.calls, ['moveSelection:1', 'scrollPopup:-90']);
  });
}

for (const c of [
  { mode: 'analog', deadzone: 0.6, button: { value: 0.7, pressed: false }, fired: true },
  { mode: 'digital', deadzone: 1, button: { value: 0.9, pressed: true }, fired: true },
  { mode: 'digital', deadzone: 0.6, button: { value: 0.9, pressed: false }, fired: false },
] as const) {
  test(`gamepad controller ${c.mode} trigger mode (deadzone ${c.deadzone}, value ${c.button.value}, pressed ${c.button.pressed}) ${c.fired ? 'fires' : 'stays idle'}`, () => {
    const trigger = { ...c.button, touched: true };
    const harness = createHarness({
      gamepads: [createGamepad('pad-1', { buttons: { 6: trigger, 7: trigger } })],
      config: createControllerConfig({
        triggerInputMode: c.mode,
        triggerDeadzone: c.deadzone,
        bindings: {
          playCurrentAudio: { kind: 'button', buttonIndex: 7 },
          toggleMpvPause: { kind: 'button', buttonIndex: 6 },
          // Free the left trigger from the default quitMpv binding.
          quitMpv: { kind: 'none' },
        },
      }),
      deps: { getLookupWindowOpen: () => true },
    });

    harness.poll(0);

    assert.deepEqual(harness.calls, c.fired ? ['playCurrentAudio', 'toggleMpvPause'] : []);
  });
}

test('gamepad controller maps L3 to mpv pause and keeps unbound audio action inactive', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 9: PRESSED } })],
    config: createControllerConfig({
      bindings: {
        toggleMpvPause: { kind: 'button', buttonIndex: 9 },
        playCurrentAudio: { kind: 'none' },
      },
    }),
    deps: { getLookupWindowOpen: () => true },
  });

  harness.poll(0);

  assert.deepEqual(harness.calls, ['toggleMpvPause']);
});
