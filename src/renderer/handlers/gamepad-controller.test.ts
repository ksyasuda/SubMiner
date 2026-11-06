import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import type { ControllerRuntimeSnapshot, ResolvedControllerConfig } from '../../types';
import { createGamepadController, type ControllerScreenLayout } from './gamepad-controller.js';

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
  'previousEntry',
  'nextEntry',
  'toggleCharacterNavigation',
  'toggleMpvPause',
  'previousSubtitle',
  'nextSubtitle',
  'replaySubtitle',
  'playNextSubtitle',
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
  // Hold changes are kept apart from `calls` so tests that press R2 for other
  // actions keep their exact call lists.
  const secondarySubHeld: boolean[] = [];
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
    previousEntry: () => {},
    nextEntry: () => {},
    toggleCharacterNavigation: () => {},
    toggleMpvPause: () => {},
    previousSubtitle: () => {},
    nextSubtitle: () => {},
    replaySubtitle: () => {},
    playNextSubtitle: () => {},
    setSecondarySubHeld: () => {},
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
  controllerOptions.setSecondarySubHeld = (held) => {
    secondarySubHeld.push(held);
    deps.setSecondarySubHeld?.(held);
  };
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
    secondarySubHeld,
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
          faceButtonLabels: globalConfig.faceButtonLabels,
          buttonIndices: globalConfig.buttonIndices,
          bindings: { ...globalConfig.bindings, toggleLookup: { kind: 'button', buttonIndex: 11 } },
          reviewBindings: globalConfig.reviewBindings,
          fieldGroupingBindings: globalConfig.fieldGroupingBindings,
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

    // Pause works outside keyboard-only mode, so it is handled before the popup actions.
    assert.deepEqual(harness.calls, c.fired ? ['toggleMpvPause', 'playCurrentAudio'] : []);
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

const HOLD_ON_R2 = { holdSecondarySub: { kind: 'button', buttonIndex: 7 } } as const;

test('gamepad controller leaves the secondary subtitle hold unbound by default', () => {
  const harness = createHarness({
    gamepads: [
      createGamepad('pad-1', {
        buttons: Object.fromEntries(Array.from({ length: 16 }, (_, i) => [i, PRESSED])),
      }),
    ],
  });

  harness.poll(0);

  assert.deepEqual(CORE_DEFAULT_CONFIG.controller.bindings.holdSecondarySub, { kind: 'none' });
  assert.deepEqual(harness.secondarySubHeld, []);
});

test('gamepad controller shows the secondary subtitle only while R2 is held, even outside keyboard-only mode', () => {
  let r2: TestButton = RELEASED;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: { 7: r2 } })],
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
    deps: { getKeyboardModeEnabled: () => false },
  });

  harness.poll(0);
  r2 = { value: 1, pressed: true, touched: true };
  harness.poll(16, 32, 48);
  r2 = RELEASED;
  harness.poll(64, 80);

  assert.deepEqual(harness.secondarySubHeld, [true, false]);
  assert.deepEqual(harness.calls, []);
});

test('gamepad controller releases the secondary subtitle hold when the pad disconnects', () => {
  let connected = true;
  const harness = createHarness({
    gamepads: () => (connected ? [createGamepad('pad-1', { buttons: { 7: PRESSED } })] : []),
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
  });

  harness.poll(0);
  connected = false;
  harness.poll(16);

  assert.deepEqual(harness.secondarySubHeld, [true, false]);
});

test('gamepad controller releases the secondary subtitle hold when a blocking modal opens', () => {
  let blocked = false;
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 7: PRESSED } })],
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
    deps: { getInteractionBlocked: () => blocked },
  });

  harness.poll(0);
  blocked = true;
  harness.poll(16);
  blocked = false;
  harness.poll(32);

  assert.deepEqual(harness.secondarySubHeld, [true, false, true]);
});

test('gamepad controller ignores the secondary subtitle hold while controller support is disabled', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 7: PRESSED } })],
    config: createControllerConfig({ enabled: false, bindings: HOLD_ON_R2 }),
  });

  harness.poll(0, 16);

  assert.deepEqual(harness.secondarySubHeld, []);
});

test('gamepad controller honors a remapped or explicitly unbound secondary subtitle hold', () => {
  const remapped = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 5: PRESSED, 7: PRESSED } })],
    config: createControllerConfig({
      bindings: { holdSecondarySub: { kind: 'button', buttonIndex: 5 } },
    }),
  });
  remapped.poll(0);
  assert.deepEqual(remapped.secondarySubHeld, [true]);

  const unbound = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 7: PRESSED } })],
    config: createControllerConfig({ bindings: { holdSecondarySub: { kind: 'none' } } }),
  });
  unbound.poll(0);
  assert.deepEqual(unbound.secondarySubHeld, []);
});

test('gamepad controller requires an analog R2 pull past the trigger deadzone to hold', () => {
  let r2: TestButton = { value: 0.3, pressed: false, touched: true };
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: { 7: r2 } })],
    config: createControllerConfig({
      triggerInputMode: 'analog',
      triggerDeadzone: 0.5,
      bindings: HOLD_ON_R2,
    }),
  });

  harness.poll(0);
  assert.deepEqual(harness.secondarySubHeld, []);

  r2 = { value: 0.7, pressed: false, touched: true };
  harness.poll(16);
  assert.deepEqual(harness.secondarySubHeld, [true]);
});

const ENTRY_ON_DPAD = {
  previousEntry: { kind: 'button', buttonIndex: 12 },
  nextEntry: { kind: 'button', buttonIndex: 13 },
  // Free the D-pad from the popup-scroll fallback, as a user binding it would.
  leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'none' },
} as const;

test('gamepad controller leaves popup entry navigation unbound by default', () => {
  assert.deepEqual(CORE_DEFAULT_CONFIG.controller.bindings.previousEntry, { kind: 'none' });
  assert.deepEqual(CORE_DEFAULT_CONFIG.controller.bindings.nextEntry, { kind: 'none' });
});

test('gamepad controller moves one popup entry per press while the lookup is open', () => {
  let buttons: Record<number, TestButton> = {};
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons })],
    config: createControllerConfig({ bindings: ENTRY_ON_DPAD }),
    deps: { getLookupWindowOpen: () => true },
  });

  harness.poll(0);
  buttons = { 13: PRESSED };
  // Holding the button does not repeat.
  harness.poll(16, 32, 400, 800);
  buttons = {};
  harness.poll(816);
  buttons = { 13: PRESSED };
  harness.poll(832);
  buttons = {};
  harness.poll(848);
  buttons = { 12: PRESSED };
  harness.poll(864);

  assert.deepEqual(harness.calls, ['nextEntry', 'nextEntry', 'previousEntry']);
});

test('gamepad controller counts a popup press right as the popup reopens after a release', () => {
  let buttons: Record<number, TestButton> = {};
  let lookupOpen = true;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons })],
    config: createControllerConfig({ bindings: ENTRY_ON_DPAD }),
    deps: { getLookupWindowOpen: () => lookupOpen },
  });

  harness.poll(0);
  buttons = { 13: PRESSED };
  harness.poll(16);
  // The popup closes while the button is still down, and the button is let go meanwhile.
  lookupOpen = false;
  harness.poll(32);
  buttons = {};
  harness.poll(48);
  // The popup reopens on the same frame the button is pressed again.
  lookupOpen = true;
  buttons = { 13: PRESSED };
  harness.poll(64);

  assert.deepEqual(harness.calls, ['nextEntry', 'nextEntry']);
});

test('gamepad controller waits for a fresh press when the popup opens under a held button', () => {
  let lookupOpen = false;
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 13: PRESSED } })],
    config: createControllerConfig({ bindings: ENTRY_ON_DPAD }),
    deps: { getLookupWindowOpen: () => lookupOpen },
  });

  harness.poll(0, 16);
  lookupOpen = true;
  harness.poll(32, 48);

  assert.deepEqual(harness.calls, []);
});

test('gamepad controller switches character navigation once per press in keyboard-only mode', () => {
  assert.deepEqual(CORE_DEFAULT_CONFIG.controller.bindings.toggleCharacterNavigation, {
    kind: 'none',
  });
  let buttons: Record<number, TestButton> = {};
  let keyboardMode = true;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons })],
    config: createControllerConfig({
      bindings: { toggleCharacterNavigation: { kind: 'button', buttonIndex: 10 } },
    }),
    deps: { getKeyboardModeEnabled: () => keyboardMode },
  });

  harness.poll(0);
  buttons = { 10: PRESSED };
  harness.poll(16, 32, 400);
  buttons = {};
  harness.poll(416);
  // Like the other lookup actions, it waits for keyboard-only mode.
  keyboardMode = false;
  buttons = { 10: PRESSED };
  harness.poll(432);

  assert.deepEqual(harness.calls, ['toggleCharacterNavigation']);
});

test('gamepad controller ignores popup entry navigation while the lookup is closed', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 12: PRESSED, 13: PRESSED } })],
    config: createControllerConfig({ bindings: ENTRY_ON_DPAD }),
    deps: { getLookupWindowOpen: () => false },
  });

  harness.poll(0, 16);

  assert.deepEqual(harness.calls, []);
});

test('gamepad controller requires keyboard-only mode for popup entry navigation', () => {
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 13: PRESSED } })],
    config: createControllerConfig({ bindings: ENTRY_ON_DPAD }),
    deps: { getLookupWindowOpen: () => true, getKeyboardModeEnabled: () => false },
  });

  harness.poll(0, 16);

  assert.deepEqual(harness.calls, []);
});

function createRecordingScreenLayout() {
  const events: string[] = [];
  const layout: ControllerScreenLayout = {
    activate: (gamepad, _config, _now, { latchHeld }) => {
      events.push(`activate:${gamepad.id}:${latchHeld ? 'latch' : 'fresh'}`);
    },
    poll: (_gamepad, _config, now) => {
      events.push(`poll:${now}`);
    },
  };
  return { layout, events };
}

test('gamepad controller hands every input to an open screen layout', () => {
  let screenOpen = true;
  const { layout, events } = createRecordingScreenLayout();
  const harness = createHarness({
    // A (lookup), Y (keyboard mode) and R2 (secondary hold) are all pressed.
    gamepads: [createGamepad('pad-1', { buttons: { 0: PRESSED, 3: PRESSED, 7: PRESSED } })],
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
    deps: { getScreenLayout: () => (screenOpen ? layout : null) },
  });

  harness.poll(0, 16, 32);
  assert.deepEqual(events, ['activate:pad-1:latch', 'poll:16', 'poll:32']);
  assert.deepEqual(harness.calls, []);
  assert.deepEqual(harness.secondarySubHeld, []);

  // Buttons still held when the screen closes do not fire their normal actions...
  screenOpen = false;
  harness.poll(48);
  assert.deepEqual(harness.calls, []);
  // ...but a level-triggered hold resumes, since it reflects what is held right now.
  assert.deepEqual(harness.secondarySubHeld, [true]);
});

test('gamepad controller re-activates the screen layout for each new session', () => {
  let screenOpen = true;
  const { layout, events } = createRecordingScreenLayout();
  const harness = createHarness({
    gamepads: [createGamepad('pad-1')],
    deps: { getScreenLayout: () => (screenOpen ? layout : null) },
  });

  harness.poll(0, 16);
  screenOpen = false;
  harness.poll(32);
  screenOpen = true;
  harness.poll(48, 64);

  assert.deepEqual(events, ['activate:pad-1:latch', 'poll:16', 'activate:pad-1:latch', 'poll:64']);
});

test('gamepad controller ignores the screen layout while controller support is disabled', () => {
  const { layout, events } = createRecordingScreenLayout();
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 0: PRESSED } })],
    config: createControllerConfig({ enabled: false }),
    deps: { getScreenLayout: () => layout },
  });

  harness.poll(0, 16);

  assert.deepEqual(events, []);
});

test('gamepad controller stands down entirely while another window owns the controller', () => {
  let suspended = true;
  let r2: TestButton = PRESSED;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: { 0: PRESSED, 3: PRESSED, 7: r2 } })],
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
  });
  harness.poll(0);
  assert.deepEqual(harness.secondarySubHeld, [true]);

  const suspendedHarness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: { 0: PRESSED, 3: PRESSED, 7: r2 } })],
    config: createControllerConfig({ bindings: HOLD_ON_R2 }),
    deps: { getSuspended: () => suspended },
  });
  suspendedHarness.poll(0, 16);
  // Not even the keyboard-mode toggle (Y) fires, and the hold stays released.
  assert.deepEqual(suspendedHarness.calls, []);
  assert.deepEqual(suspendedHarness.secondarySubHeld, []);

  suspended = false;
  r2 = RELEASED;
  suspendedHarness.poll(32);
  assert.deepEqual(suspendedHarness.calls, []);
});

test('gamepad controller counts the press that reveals a pad to an open screen', () => {
  let visible = false;
  const { layout, events } = createRecordingScreenLayout();
  const harness = createHarness({
    // Until a button is pressed in a fresh window, the browser reports no pads at all.
    gamepads: () => (visible ? [createGamepad('pad-1', { buttons: { 0: PRESSED } })] : []),
    deps: { getScreenLayout: () => layout },
  });

  harness.poll(0, 16);
  visible = true;
  harness.poll(32, 48);

  assert.deepEqual(events, ['activate:pad-1:fresh', 'poll:48']);
});

test('gamepad controller still ignores a button held when the screen opened', () => {
  let screenOpen = false;
  const { layout, events } = createRecordingScreenLayout();
  const harness = createHarness({
    gamepads: [createGamepad('pad-1', { buttons: { 2: PRESSED } })],
    deps: { getScreenLayout: () => (screenOpen ? layout : null) },
  });

  harness.poll(0);
  screenOpen = true;
  harness.poll(16);

  assert.deepEqual(events, ['activate:pad-1:latch']);
});

const SUBTITLE_NAVIGATION_BINDINGS = {
  previousSubtitle: { kind: 'button', buttonIndex: 10 },
  nextSubtitle: { kind: 'button', buttonIndex: 11 },
  replaySubtitle: { kind: 'button', buttonIndex: 8 },
  playNextSubtitle: { kind: 'button', buttonIndex: 16 },
} as const;

test('gamepad controller leaves subtitle navigation unbound by default', () => {
  for (const action of [
    'previousSubtitle',
    'nextSubtitle',
    'replaySubtitle',
    'playNextSubtitle',
  ] as const) {
    assert.deepEqual(CORE_DEFAULT_CONFIG.controller.bindings[action], { kind: 'none' }, action);
  }
});

test('gamepad controller navigates subtitles once per press, even outside keyboard-only mode', () => {
  let pressed: number[] = [];
  const harness = createHarness({
    gamepads: () => [
      createGamepad('pad-1', {
        buttons: Object.fromEntries(pressed.map((index) => [index, PRESSED])),
      }),
    ],
    config: createControllerConfig({ bindings: SUBTITLE_NAVIGATION_BINDINGS }),
    deps: { getKeyboardModeEnabled: () => false },
  });

  harness.poll(0);
  for (const index of [11, 10, 8, 16]) {
    pressed = [index];
    // Holding the button does not repeat.
    harness.poll(16, 500, 1_000);
    pressed = [];
    harness.poll(1_016);
  }

  assert.deepEqual(harness.calls, [
    'nextSubtitle',
    'previousSubtitle',
    'replaySubtitle',
    'playNextSubtitle',
  ]);
});

test('gamepad controller pauses subtitle navigation while a blocking modal is open', () => {
  let blocked = true;
  let pressed = true;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: pressed ? { 11: PRESSED } : {} })],
    config: createControllerConfig({ bindings: SUBTITLE_NAVIGATION_BINDINGS }),
    deps: { getInteractionBlocked: () => blocked },
  });

  harness.poll(0, 16);
  // Still held when the modal closes: not a fresh press.
  blocked = false;
  harness.poll(32);
  assert.deepEqual(harness.calls, []);

  pressed = false;
  harness.poll(48);
  pressed = true;
  harness.poll(64);
  assert.deepEqual(harness.calls, ['nextSubtitle']);
});

test('gamepad controller pauses mpv outside keyboard-only mode, but not under a blocking modal', () => {
  let blocked = false;
  let pressed = false;
  const harness = createHarness({
    gamepads: () => [createGamepad('pad-1', { buttons: pressed ? { 9: PRESSED } : {} })],
    config: createControllerConfig({
      bindings: { toggleMpvPause: { kind: 'button', buttonIndex: 9 } },
    }),
    deps: {
      getKeyboardModeEnabled: () => false,
      getInteractionBlocked: () => blocked,
    },
  });

  harness.poll(0);
  pressed = true;
  // Held: fires once.
  harness.poll(16, 500);
  pressed = false;
  harness.poll(516);
  blocked = true;
  pressed = true;
  harness.poll(532);

  assert.deepEqual(harness.calls, ['toggleMpvPause']);
});

/** Harness driven frame by frame: `pads(4, 15)` polls once with those buttons down. */
function createComboHarness(
  bindings: Partial<Bindings>,
  deps: Partial<ControllerOptions> = {},
  extraConfig: Partial<ResolvedControllerConfig> = {},
) {
  let pressed: number[] = [];
  let now = 0;
  const harness = createHarness({
    gamepads: () => [
      createGamepad('pad-1', {
        buttons: Object.fromEntries(pressed.map((index) => [index, PRESSED])),
      }),
    ],
    config: createControllerConfig({ ...extraConfig, bindings }),
    deps: { getLookupWindowOpen: () => true, ...deps },
  });
  return {
    ...harness,
    pads: (...buttons: number[]) => {
      pressed = buttons;
      now += 16;
      harness.poll(now);
    },
  };
}

const L1 = 4;
const R1 = 5;
const DPAD_UP = 12;
const DPAD_RIGHT = 15;
const L1_DPAD_RIGHT = {
  kind: 'button',
  buttonIndex: DPAD_RIGHT,
  hold: { kind: 'button', buttonIndex: L1 },
} as const;

test('holding L1 and pressing D-pad right twice fires the combo twice and nothing else', () => {
  // L1 keeps its default Play Current Audio, and the D-pad keeps moving between words.
  const harness = createComboHarness({ nextAudio: L1_DPAD_RIGHT });

  harness.pads();
  harness.pads(L1);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads(L1);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads(L1);
  harness.pads();

  assert.deepEqual(harness.calls, ['nextAudio', 'nextAudio']);
});

test('a plain button that is also a hold button fires when tapped, on release', () => {
  const harness = createComboHarness({ nextAudio: L1_DPAD_RIGHT });

  harness.pads();
  harness.pads(L1);
  assert.deepEqual(harness.calls, []);
  harness.pads();
  assert.deepEqual(harness.calls, ['playCurrentAudio']);

  // The next tap is unaffected by a combo used on an earlier hold.
  harness.pads(L1);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads();
  harness.pads(L1);
  harness.pads();
  assert.deepEqual(harness.calls, ['playCurrentAudio', 'nextAudio', 'playCurrentAudio']);
});

test('a plain button stays on press while no combo holds it', () => {
  const harness = createComboHarness({ nextAudio: { kind: 'button', buttonIndex: R1 } });

  harness.pads();
  harness.pads(L1);

  assert.deepEqual(harness.calls, ['playCurrentAudio']);
});

test('a tap on a hold button also fires outside keyboard-only mode', () => {
  const harness = createComboHarness(
    {
      previousSubtitle: { kind: 'button', buttonIndex: L1 },
      nextSubtitle: { kind: 'button', buttonIndex: R1, hold: { kind: 'button', buttonIndex: L1 } },
    },
    { getKeyboardModeEnabled: () => false },
  );

  harness.pads();
  harness.pads(L1);
  harness.pads();

  assert.deepEqual(harness.calls, ['previousSubtitle']);
});

test('a tap on a hold button is dropped if its action stops applying before the release', () => {
  let lookupOpen = true;
  const harness = createComboHarness(
    { nextAudio: L1_DPAD_RIGHT },
    { getLookupWindowOpen: () => lookupOpen },
  );

  harness.pads();
  harness.pads(L1);
  lookupOpen = false;
  harness.pads(L1);
  lookupOpen = true;
  harness.pads(L1);
  harness.pads();

  assert.deepEqual(harness.calls, []);
});

test('pressing a plain binding during a hold does nothing and cancels the tap', () => {
  const harness = createComboHarness({ nextAudio: L1_DPAD_RIGHT });

  harness.pads();
  harness.pads(L1);
  harness.pads(L1, 2); // X is Mine Card; there is no L1 + X.
  harness.pads();

  assert.deepEqual(harness.calls, []);
});

test('letting go of the hold first does not fire the plain binding on the same input', () => {
  const harness = createComboHarness({
    playCurrentAudio: { kind: 'none' },
    nextAudio: L1_DPAD_RIGHT,
    nextEntry: { kind: 'button', buttonIndex: DPAD_RIGHT },
    leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'none' },
  });

  harness.pads();
  harness.pads(L1);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads(DPAD_RIGHT);
  harness.pads();
  assert.deepEqual(harness.calls, ['nextAudio']);

  harness.pads(DPAD_RIGHT);
  assert.deepEqual(harness.calls, ['nextAudio', 'nextEntry']);
});

test('a combo needs its hold down first, and two holds at once match nothing', () => {
  const harness = createComboHarness({
    playCurrentAudio: { kind: 'none' },
    nextAudio: L1_DPAD_RIGHT,
    previousAudio: {
      kind: 'button',
      buttonIndex: DPAD_RIGHT,
      hold: { kind: 'button', buttonIndex: R1 },
    },
  });

  harness.pads();
  harness.pads(DPAD_RIGHT);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads();
  harness.pads(L1, R1);
  harness.pads(L1, R1, DPAD_RIGHT);

  // Only the first press of the D-pad moves a word; the combos never fire.
  assert.deepEqual(harness.calls, ['moveSelection:1']);
});

test('a combo keeps the D-pad from driving its stick row until the D-pad is released', () => {
  const harness = createComboHarness({
    playCurrentAudio: { kind: 'none' },
    nextAudio: L1_DPAD_RIGHT,
    nextEntry: {
      kind: 'button',
      buttonIndex: DPAD_UP,
      hold: { kind: 'button', buttonIndex: L1 },
    },
  });

  harness.pads();
  harness.pads(L1);
  harness.pads(L1, DPAD_RIGHT);
  harness.pads(L1, DPAD_UP);
  harness.pads(DPAD_UP);
  harness.pads();
  assert.deepEqual(harness.calls, ['nextAudio', 'nextEntry']);

  // L1 with a direction no combo uses still moves between words.
  harness.pads(L1);
  harness.pads(L1, 14);
  assert.deepEqual(harness.calls, ['nextAudio', 'nextEntry', 'moveSelection:-1']);
});

test('L1 and L2 can each hold the same press for different actions', () => {
  const L2 = 6;
  const DPAD_LEFT = 14;
  const harness = createComboHarness({
    // L2 keeps its default Quit MPV, so this also pins that a combo never quits on release.
    playCurrentAudio: { kind: 'none' },
    nextAudio: {
      kind: 'button',
      buttonIndex: DPAD_LEFT,
      hold: { kind: 'button', buttonIndex: L1 },
    },
    nextEntry: {
      kind: 'button',
      buttonIndex: DPAD_LEFT,
      hold: { kind: 'button', buttonIndex: L2 },
    },
  });

  harness.pads();
  harness.pads(L1);
  harness.pads(L1, DPAD_LEFT);
  harness.pads();
  harness.pads(L2);
  harness.pads(L2, DPAD_LEFT);
  harness.pads();

  assert.deepEqual(harness.calls, ['nextAudio', 'nextEntry']);
});

test('the secondary subtitle hold can be a combo', () => {
  const harness = createComboHarness({
    playCurrentAudio: { kind: 'none' },
    holdSecondarySub: {
      kind: 'button',
      buttonIndex: R1,
      hold: { kind: 'button', buttonIndex: L1 },
    },
    nextAudio: { kind: 'none' },
  });

  harness.pads();
  harness.pads(R1);
  harness.pads();
  harness.pads(L1);
  harness.pads(L1, R1);
  harness.pads(L1);

  assert.deepEqual(harness.secondarySubHeld, [true, false]);
});

test('a dialog owning the pad keeps combos from firing and from leaving a tap behind', () => {
  let suspended = false;
  const harness = createComboHarness(
    { nextAudio: L1_DPAD_RIGHT },
    { getSuspended: () => suspended },
  );

  harness.pads();
  harness.pads(L1);
  suspended = true;
  harness.pads(L1, DPAD_RIGHT);
  suspended = false;
  harness.pads(L1, DPAD_RIGHT);
  harness.pads();

  assert.deepEqual(harness.calls, []);
});
