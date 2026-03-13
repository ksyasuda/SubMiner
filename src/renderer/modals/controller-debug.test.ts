import assert from 'node:assert/strict';
import test from 'node:test';

import { createRendererState } from '../state.js';
import { createControllerDebugModal } from './controller-debug.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

test('controller debug modal renders active controller axes, buttons, and config-ready button indices', () => {
  const globals = globalThis as typeof globalThis & { window?: unknown };
  const previousWindow = globals.window;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const state = createRendererState();
    state.connectedGamepads = [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }];
    state.activeGamepadId = 'pad-1';
    state.controllerRawAxes = [0.5, -0.25];
    state.controllerRawButtons = [{ value: 1, pressed: true, touched: true }];
    state.controllerConfig = {
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
      },
    };

    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        controllerDebugModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
        },
        controllerDebugClose: { addEventListener: () => {} },
        controllerDebugToast: { textContent: '', classList: createClassList(['hidden']) },
        controllerDebugStatus: { textContent: '', classList: createClassList() },
        controllerDebugSummary: { textContent: '' },
        controllerDebugAxes: { textContent: '' },
        controllerDebugButtons: { textContent: '' },
        controllerDebugButtonIndices: { textContent: '' },
      },
      state,
    };

    const modal = createControllerDebugModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    modal.openControllerDebugModal();

    assert.match(ctx.dom.controllerDebugStatus.textContent, /pad-1/);
    assert.match(ctx.dom.controllerDebugSummary.textContent, /standard/);
    assert.match(ctx.dom.controllerDebugAxes.textContent, /axis\[0\] = 0\.500/);
    assert.match(
      ctx.dom.controllerDebugButtons.textContent,
      /button\[0\] value=1\.000 pressed=true/,
    );
    assert.match(ctx.dom.controllerDebugButtonIndices.textContent, /"buttonIndices": \{/);
    assert.match(ctx.dom.controllerDebugButtonIndices.textContent, /"select": 6/);
    assert.match(ctx.dom.controllerDebugButtonIndices.textContent, /"leftStickPress": 9/);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
  }
});

test('controller debug modal copies buttonIndices config to clipboard', async () => {
  const globals = globalThis as typeof globalThis & {
    window?: unknown;
    navigator?: unknown;
  };
  const previousWindow = globals.window;
  const previousNavigator = globals.navigator;
  const copied: string[] = [];
  const handlers: { copy: null | (() => void) } = { copy: null };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        notifyOverlayModalClosed: () => {},
      },
    },
  });
  Object.defineProperty(globalThis, 'navigator', {
    configurable: true,
    value: {
      clipboard: {
        writeText: async (text: string) => {
          copied.push(text);
        },
      },
    },
  });

  try {
    const state = createRendererState();
    state.controllerConfig = {
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
      },
    };

    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        controllerDebugModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
        },
        controllerDebugClose: { addEventListener: () => {} },
        controllerDebugCopy: {
          addEventListener: (_event: string, handler: () => void) => {
            handlers.copy = handler;
          },
        },
        controllerDebugToast: { textContent: '', classList: createClassList(['hidden']) },
        controllerDebugStatus: { textContent: '', classList: createClassList() },
        controllerDebugSummary: { textContent: '' },
        controllerDebugAxes: { textContent: '' },
        controllerDebugButtons: { textContent: '' },
        controllerDebugButtonIndices: { textContent: '' },
      },
      state,
    };

    const modal = createControllerDebugModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerDebugModal();
    if (handlers.copy) {
      handlers.copy();
    }
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(copied, [ctx.dom.controllerDebugButtonIndices.textContent]);
    assert.match(
      ctx.dom.controllerDebugStatus.textContent,
      /Copied controller buttonIndices config/,
    );
    assert.match(
      ctx.dom.controllerDebugToast.textContent,
      /Copied controller buttonIndices config/,
    );
    assert.equal(ctx.dom.controllerDebugToast.classList.contains('hidden'), false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: previousNavigator,
    });
  }
});
