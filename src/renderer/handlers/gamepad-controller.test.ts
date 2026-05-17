import assert from 'node:assert/strict';
import test from 'node:test';

import type { ResolvedControllerConfig } from '../../types';
import { createGamepadController } from './gamepad-controller.js';

type TestGamepad = {
  id: string;
  index: number;
  connected: boolean;
  mapping: string;
  axes: number[];
  buttons: Array<{ value: number; pressed?: boolean; touched?: boolean }>;
};

const DEFAULT_BUTTON_INDICES = {
  select: 6,
  buttonSouth: 0,
  buttonEast: 1,
  buttonWest: 2,
  buttonNorth: 3,
  leftShoulder: 4,
  rightShoulder: 5,
  leftStickPress: 9,
  rightStickPress: 10,
  leftTrigger: 6,
  rightTrigger: 7,
} satisfies ResolvedControllerConfig['buttonIndices'];

function createGamepad(
  id: string,
  options: Partial<Pick<TestGamepad, 'index' | 'axes' | 'buttons'>> = {},
): TestGamepad {
  return {
    id,
    index: options.index ?? 0,
    connected: true,
    mapping: 'standard',
    axes: options.axes ?? [0, 0, 0, 0],
    buttons:
      options.buttons ??
      Array.from({ length: 16 }, () => ({
        value: 0,
        pressed: false,
        touched: false,
      })),
  };
}

function createControllerConfig(
  overrides: Omit<Partial<ResolvedControllerConfig>, 'bindings' | 'buttonIndices'> & {
    bindings?: Partial<Record<keyof ResolvedControllerConfig['bindings'], unknown>>;
    buttonIndices?: Partial<ResolvedControllerConfig['buttonIndices']>;
  } = {},
): ResolvedControllerConfig {
  const {
    bindings: bindingOverrides,
    buttonIndices: buttonIndexOverrides,
    ...restOverrides
  } = overrides;
  return {
    enabled: true,
    preferredGamepadId: '',
    preferredGamepadLabel: '',
    smoothScroll: true,
    scrollPixelsPerSecond: 900,
    horizontalJumpPixels: 160,
    stickDeadzone: 0.2,
    triggerInputMode: 'auto',
    triggerDeadzone: 0.5,
    repeatDelayMs: 320,
    repeatIntervalMs: 120,
    buttonIndices: {
      ...DEFAULT_BUTTON_INDICES,
      ...(buttonIndexOverrides ?? {}),
    },
    bindings: {
      toggleLookup: { kind: 'button', buttonIndex: 0 },
      closeLookup: { kind: 'button', buttonIndex: 1 },
      toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
      mineCard: { kind: 'button', buttonIndex: 2 },
      quitMpv: { kind: 'button', buttonIndex: 6 },
      previousAudio: { kind: 'none' },
      nextAudio: { kind: 'button', buttonIndex: 5 },
      playCurrentAudio: { kind: 'button', buttonIndex: 4 },
      toggleMpvPause: { kind: 'button', buttonIndex: 9 },
      leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
      leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
      rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
      rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
      ...normalizeBindingOverrides(bindingOverrides ?? {}, {
        ...DEFAULT_BUTTON_INDICES,
        ...(buttonIndexOverrides ?? {}),
      }),
    },
    profiles: {},
    ...restOverrides,
  };
}

function normalizeBindingOverrides(
  overrides: Partial<Record<keyof ResolvedControllerConfig['bindings'], unknown>>,
  buttonIndices: ResolvedControllerConfig['buttonIndices'],
): Partial<ResolvedControllerConfig['bindings']> {
  const legacyButtonIndices = {
    select: buttonIndices.select,
    buttonSouth: buttonIndices.buttonSouth,
    buttonEast: buttonIndices.buttonEast,
    buttonWest: buttonIndices.buttonWest,
    buttonNorth: buttonIndices.buttonNorth,
    leftShoulder: buttonIndices.leftShoulder,
    rightShoulder: buttonIndices.rightShoulder,
    leftStickPress: buttonIndices.leftStickPress,
    rightStickPress: buttonIndices.rightStickPress,
    leftTrigger: buttonIndices.leftTrigger,
    rightTrigger: buttonIndices.rightTrigger,
  } as const;
  const legacyAxisIndices = {
    leftStickX: 0,
    leftStickY: 1,
    rightStickX: 3,
    rightStickY: 4,
  } as const;
  const axisFallbackByKey = {
    leftStickHorizontal: 'horizontal',
    leftStickVertical: 'vertical',
    rightStickHorizontal: 'none',
    rightStickVertical: 'none',
  } as const;

  const normalized: Partial<ResolvedControllerConfig['bindings']> = {};
  for (const [key, value] of Object.entries(overrides) as Array<
    [keyof ResolvedControllerConfig['bindings'], unknown]
  >) {
    if (typeof value === 'string') {
      if (value === 'none') {
        normalized[key] = { kind: 'none' } as never;
        continue;
      }
      if (value in legacyButtonIndices) {
        normalized[key] = {
          kind: 'button',
          buttonIndex: legacyButtonIndices[value as keyof typeof legacyButtonIndices],
        } as never;
        continue;
      }
      if (value in legacyAxisIndices) {
        normalized[key] = {
          kind: 'axis',
          axisIndex: legacyAxisIndices[value as keyof typeof legacyAxisIndices],
          dpadFallback: axisFallbackByKey[key as keyof typeof axisFallbackByKey] ?? 'none',
        } as never;
        continue;
      }
    }
    normalized[key] = value as never;
  }
  return normalized;
}

test('gamepad controller selects the first connected controller by default', () => {
  const updates: string[] = [];
  const controller = createGamepadController({
    getGamepads: () => [
      null,
      createGamepad('pad-2', { index: 1 }),
      createGamepad('pad-3', { index: 2 }),
    ],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => false,
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
    onState: (state) => {
      updates.push(state.activeGamepadId ?? 'none');
    },
  });

  controller.poll(0);

  assert.equal(controller.getActiveGamepadId(), 'pad-2');
  assert.deepEqual(updates.at(-1), 'pad-2');
});

test('gamepad controller prefers saved controller id when connected', () => {
  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1'), createGamepad('pad-2', { index: 1 })],
    getConfig: () => createControllerConfig({ preferredGamepadId: 'pad-2' }),
    getKeyboardModeEnabled: () => false,
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
    onState: () => {},
  });

  controller.poll(0);

  assert.equal(controller.getActiveGamepadId(), 'pad-2');
});

test('gamepad controller allows keyboard-mode toggle while other actions stay gated', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 8 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[0] = { value: 1, pressed: true, touched: true };
  buttons[3] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => false,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => calls.push('toggle-keyboard-mode'),
    toggleLookup: () => calls.push('toggle-lookup'),
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
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['toggle-keyboard-mode']);
});

test('gamepad controller re-evaluates interaction gating after toggling keyboard mode', () => {
  const calls: string[] = [];
  let keyboardModeEnabled = true;
  const buttons = Array.from({ length: 8 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[0] = { value: 1, pressed: true, touched: true };
  buttons[3] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => keyboardModeEnabled,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {
      calls.push('toggle-keyboard-mode');
      keyboardModeEnabled = false;
    },
    toggleLookup: () => calls.push('toggle-lookup'),
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
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['toggle-keyboard-mode']);
});

test('gamepad controller resets edge state when active controller changes', () => {
  const calls: string[] = [];
  let currentGamepads = [
    createGamepad('pad-1', {
      buttons: [{ value: 1, pressed: true, touched: true }],
    }),
  ];

  const controller = createGamepadController({
    getGamepads: () => currentGamepads,
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => calls.push('toggle-lookup'),
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
    onState: () => {},
  });

  controller.poll(0);
  currentGamepads = [
    createGamepad('pad-2', {
      buttons: [{ value: 1, pressed: true, touched: true }],
    }),
  ];
  controller.poll(50);

  assert.deepEqual(calls, ['toggle-lookup', 'toggle-lookup']);
});

test('gamepad controller does not toggle keyboard mode when controller support is disabled', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 8 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[3] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () => createControllerConfig({ enabled: false }),
    getKeyboardModeEnabled: () => false,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => calls.push('toggle-keyboard-mode'),
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
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, []);
});

test('gamepad controller does not treat blocked held inputs as fresh edges when interaction resumes', () => {
  const calls: string[] = [];
  const selectionCalls: number[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[0] = { value: 1, pressed: true, touched: true };
  let axes = [0.9, 0, 0, 0];
  let keyboardModeEnabled = true;
  let interactionBlocked = true;

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons, axes })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => keyboardModeEnabled,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => interactionBlocked,
    toggleKeyboardMode: () => {},
    toggleLookup: () => calls.push('toggle-lookup'),
    closeLookup: () => {},
    moveSelection: (delta) => selectionCalls.push(delta),
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);
  interactionBlocked = false;
  controller.poll(100);

  assert.deepEqual(calls, []);
  assert.deepEqual(selectionCalls, []);

  buttons[0] = { value: 0, pressed: false, touched: false };
  axes = [0, 0, 0, 0];
  controller.poll(200);

  buttons[0] = { value: 1, pressed: true, touched: true };
  axes = [0.9, 0, 0, 0];
  controller.poll(300);

  assert.deepEqual(calls, ['toggle-lookup']);
  assert.deepEqual(selectionCalls, [1]);
});

test('gamepad controller maps left stick horizontal movement to token selection repeats', () => {
  const calls: number[] = [];
  let axes = [0.9, 0, 0, 0];
  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { axes })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: (delta) => calls.push(delta),
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);
  controller.poll(100);
  controller.poll(260);

  assert.deepEqual(calls, [1]);

  controller.poll(340);

  assert.deepEqual(calls, [1, 1]);

  axes = [0, 0, 0, 0];
  controller.poll(360);
  axes = [-0.9, 0, 0, 0];
  controller.poll(380);

  assert.deepEqual(calls, [1, 1, -1]);
});

test('gamepad controller uses active controller profile bindings before global bindings', () => {
  let lookupToggles = 0;
  const buttons = Array.from({ length: 12 }, () => ({
    value: 0,
    pressed: false,
    touched: false,
  }));
  buttons[11] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-profile', { buttons })],
    getConfig: () =>
      ({
        ...createControllerConfig({
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 0 },
          },
        }),
        profiles: {
          'pad-profile': {
            label: 'Profile Pad',
            buttonIndices: DEFAULT_BUTTON_INDICES,
            bindings: {
              ...createControllerConfig().bindings,
              toggleLookup: { kind: 'button', buttonIndex: 11 },
            },
          },
        },
      }) as ResolvedControllerConfig,
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {
      lookupToggles += 1;
    },
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
    onState: () => {},
  });

  controller.poll(0);

  assert.equal(lookupToggles, 1);
});

test('gamepad controller maps L1 play-current, R1 next-audio, and popup navigation', () => {
  const calls: string[] = [];
  const scrollCalls: number[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[8] = { value: 1, pressed: true, touched: true };
  buttons[4] = { value: 1, pressed: true, touched: true };
  buttons[5] = { value: 1, pressed: true, touched: true };
  buttons[6] = { value: 0.8, pressed: true, touched: true };
  buttons[7] = { value: 0.9, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [
      createGamepad('pad-1', {
        axes: [0, -0.75, 0.1, 0, 0.8],
        buttons,
      }),
    ],
    getConfig: () =>
      createControllerConfig({
        bindings: {
          playCurrentAudio: 'leftShoulder',
          nextAudio: 'rightShoulder',
          previousAudio: 'none',
          toggleMpvPause: 'leftTrigger',
        },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => calls.push('quit-mpv'),
    previousAudio: () => calls.push('prev-audio'),
    nextAudio: () => calls.push('next-audio'),
    playCurrentAudio: () => calls.push('play-audio'),
    toggleMpvPause: () => calls.push('toggle-mpv-pause'),
    scrollPopup: (delta) => scrollCalls.push(delta),
    jumpPopup: (delta) => calls.push(`jump:${delta}`),
    onState: () => {},
  });

  controller.poll(0);
  controller.poll(100);

  assert.equal(calls.includes('next-audio'), true);
  assert.equal(calls.includes('play-audio'), true);
  assert.equal(calls.includes('prev-audio'), false);
  assert.equal(calls.includes('toggle-mpv-pause'), true);
  assert.equal(calls.includes('quit-mpv'), true);
  assert.deepEqual(
    scrollCalls.map((value) => Math.round(value)),
    [-67],
  );
  assert.equal(calls.includes('jump:160'), true);
});

test('gamepad controller maps quit mpv select binding from raw button 6 by default', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[6] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () => createControllerConfig({ bindings: { quitMpv: 'select' } }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => calls.push('quit-mpv'),
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['quit-mpv']);
});

test('gamepad controller honors configured raw button index overrides', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[11] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () =>
      createControllerConfig({
        buttonIndices: {
          select: 11,
        },
        bindings: { quitMpv: 'select' },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => false,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => calls.push('quit-mpv'),
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['quit-mpv']);
});

test('gamepad controller maps right stick vertical to popup jump and ignores horizontal movement', () => {
  const calls: string[] = [];
  let axes = [0, 0, 0.85, 0, 0];

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { axes })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
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
    jumpPopup: (delta) => calls.push(`jump:${delta}`),
    onState: () => {},
  });

  controller.poll(0);
  controller.poll(100);

  assert.deepEqual(calls, []);

  axes = [0, 0, 0.85, 0, -0.85];
  controller.poll(200);

  assert.deepEqual(calls, ['jump:-160']);
});

test('gamepad controller maps d-pad left/right to selection and d-pad up/down to popup scroll', () => {
  const selectionCalls: number[] = [];
  const scrollCalls: number[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[15] = { value: 1, pressed: false, touched: true };
  buttons[12] = { value: 1, pressed: false, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: (delta) => selectionCalls.push(delta),
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: (delta) => scrollCalls.push(delta),
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);
  controller.poll(100);

  assert.deepEqual(selectionCalls, [1]);
  assert.deepEqual(
    scrollCalls.map((value) => Math.round(value)),
    [-90],
  );
});

test('gamepad controller maps d-pad axes 6 and 7 to selection and popup scroll', () => {
  const selectionCalls: number[] = [];
  const scrollCalls: number[] = [];

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { axes: [0, 0, 0, 0, 0, 0, 1, -1] })],
    getConfig: () => createControllerConfig(),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: (delta) => selectionCalls.push(delta),
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => {},
    toggleMpvPause: () => {},
    scrollPopup: (delta) => scrollCalls.push(delta),
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);
  controller.poll(100);

  assert.deepEqual(selectionCalls, [1]);
  assert.deepEqual(
    scrollCalls.map((value) => Math.round(value)),
    [-90],
  );
});

test('gamepad controller trigger analog mode uses trigger values above threshold', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[6] = { value: 0.7, pressed: false, touched: true };
  buttons[7] = { value: 0.8, pressed: false, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () =>
      createControllerConfig({
        triggerInputMode: 'analog',
        triggerDeadzone: 0.6,
        bindings: {
          playCurrentAudio: 'rightTrigger',
          toggleMpvPause: 'leftTrigger',
        },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => calls.push('play-audio'),
    toggleMpvPause: () => calls.push('toggle-mpv-pause'),
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['play-audio', 'toggle-mpv-pause']);
});

test('gamepad controller trigger digital mode uses pressed state only', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[6] = { value: 0.9, pressed: true, touched: true };
  buttons[7] = { value: 0.9, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () =>
      createControllerConfig({
        triggerInputMode: 'digital',
        triggerDeadzone: 1,
        bindings: {
          playCurrentAudio: 'rightTrigger',
          toggleMpvPause: 'leftTrigger',
        },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => calls.push('play-audio'),
    toggleMpvPause: () => calls.push('toggle-mpv-pause'),
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['play-audio', 'toggle-mpv-pause']);
});

test('gamepad controller digital trigger bindings ignore analog-only trigger values', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[6] = { value: 0.9, pressed: false, touched: true };
  buttons[7] = { value: 0.9, pressed: false, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () =>
      createControllerConfig({
        triggerInputMode: 'digital',
        triggerDeadzone: 0.6,
        bindings: {
          playCurrentAudio: 'rightTrigger',
          toggleMpvPause: 'leftTrigger',
        },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => calls.push('play-audio'),
    toggleMpvPause: () => calls.push('toggle-mpv-pause'),
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, []);
});

test('gamepad controller maps L3 to mpv pause and keeps unbound audio action inactive', () => {
  const calls: string[] = [];
  const buttons = Array.from({ length: 16 }, () => ({ value: 0, pressed: false, touched: false }));
  buttons[9] = { value: 1, pressed: true, touched: true };

  const controller = createGamepadController({
    getGamepads: () => [createGamepad('pad-1', { buttons })],
    getConfig: () =>
      createControllerConfig({
        bindings: {
          toggleMpvPause: 'leftStickPress',
          playCurrentAudio: 'none',
        },
      }),
    getKeyboardModeEnabled: () => true,
    getLookupWindowOpen: () => true,
    getInteractionBlocked: () => false,
    toggleKeyboardMode: () => {},
    toggleLookup: () => {},
    closeLookup: () => {},
    moveSelection: () => {},
    mineCard: () => {},
    quitMpv: () => {},
    previousAudio: () => {},
    nextAudio: () => {},
    playCurrentAudio: () => calls.push('play-audio'),
    toggleMpvPause: () => calls.push('toggle-mpv-pause'),
    scrollPopup: () => {},
    jumpPopup: () => {},
    onState: () => {},
  });

  controller.poll(0);

  assert.deepEqual(calls, ['toggle-mpv-pause']);
});
