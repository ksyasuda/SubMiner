import assert from 'node:assert/strict';
import test from 'node:test';

import { createRendererState } from '../state.js';
import { createControllerSelectModal } from './controller-select.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    toggle: (entry: string, force?: boolean) => {
      if (force === undefined) {
        if (tokens.has(entry)) tokens.delete(entry);
        else tokens.add(entry);
        return tokens.has(entry);
      }
      if (force) tokens.add(entry);
      else tokens.delete(entry);
      return force;
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

test('controller select modal saves the selected preferred controller', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const saved: Array<{ preferredGamepadId: string; preferredGamepadLabel: string }> = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerPreference: async (update: {
          preferredGamepadId: string;
          preferredGamepadLabel: string;
        }) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({
        className: '',
        textContent: '',
        classList: createClassList(),
        appendChild: () => {},
        addEventListener: () => {},
      }),
    },
  });

  try {
    const overlayClassList = createClassList();
    const state = createRendererState();
    state.controllerConfig = {
      enabled: true,
      preferredGamepadId: 'pad-2',
      preferredGamepadLabel: 'pad-2',
      smoothScroll: true,
      scrollPixelsPerSecond: 960,
      horizontalJumpPixels: 160,
      stickDeadzone: 0.2,
      triggerInputMode: 'auto',
      triggerDeadzone: 0.5,
      repeatDelayMs: 220,
      repeatIntervalMs: 80,
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
        toggleLookup: 'buttonSouth',
        closeLookup: 'buttonEast',
        toggleKeyboardOnlyMode: 'buttonNorth',
        mineCard: 'buttonWest',
        quitMpv: 'select',
        previousAudio: 'leftShoulder',
        nextAudio: 'rightShoulder',
        playCurrentAudio: 'rightTrigger',
        toggleMpvPause: 'leftTrigger',
        leftStickHorizontal: 'leftStickX',
        leftStickVertical: 'leftStickY',
        rightStickHorizontal: 'rightStickX',
        rightStickVertical: 'rightStickY',
      },
    };
    state.connectedGamepads = [
      { id: 'pad-1', index: 0, mapping: 'standard', connected: true },
      { id: 'pad-2', index: 1, mapping: 'standard', connected: true },
    ];
    state.activeGamepadId = 'pad-2';

    const ctx = {
      dom: {
        overlay: { classList: overlayClassList, focus: () => {} },
        controllerSelectModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
        },
        controllerSelectClose: { addEventListener: () => {} },
        controllerSelectHint: { textContent: '' },
        controllerSelectStatus: { textContent: '', classList: createClassList() },
        controllerSelectList: {
          innerHTML: '',
          appendChild: () => {},
        },
        controllerSelectSave: { addEventListener: () => {} },
      },
      state,
    };

    const modal = createControllerSelectModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    modal.openControllerSelectModal();
    assert.equal(state.controllerDeviceSelectedIndex, 1);

    await modal.handleControllerSelectKeydown({
      key: 'Enter',
      preventDefault: () => {},
    } as KeyboardEvent);

    assert.deepEqual(saved, [
      {
        preferredGamepadId: 'pad-2',
        preferredGamepadLabel: 'pad-2',
      },
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('controller select modal preserves manual selection while controller polling updates', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerPreference: async () => {},
        notifyOverlayModalClosed: () => {},
      },
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({
        className: '',
        textContent: '',
        classList: createClassList(),
        appendChild: () => {},
        addEventListener: () => {},
      }),
    },
  });

  try {
    const state = createRendererState();
    state.controllerConfig = {
      enabled: true,
      preferredGamepadId: 'pad-1',
      preferredGamepadLabel: 'pad-1',
      smoothScroll: true,
      scrollPixelsPerSecond: 960,
      horizontalJumpPixels: 160,
      stickDeadzone: 0.2,
      triggerInputMode: 'auto',
      triggerDeadzone: 0.5,
      repeatDelayMs: 220,
      repeatIntervalMs: 80,
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
        toggleLookup: 'buttonSouth',
        closeLookup: 'buttonEast',
        toggleKeyboardOnlyMode: 'buttonNorth',
        mineCard: 'buttonWest',
        quitMpv: 'select',
        previousAudio: 'none',
        nextAudio: 'rightShoulder',
        playCurrentAudio: 'leftShoulder',
        toggleMpvPause: 'leftStickPress',
        leftStickHorizontal: 'leftStickX',
        leftStickVertical: 'leftStickY',
        rightStickHorizontal: 'rightStickX',
        rightStickVertical: 'rightStickY',
      },
    };
    state.connectedGamepads = [
      { id: 'pad-1', index: 0, mapping: 'standard', connected: true },
      { id: 'pad-2', index: 1, mapping: 'standard', connected: true },
    ];
    state.activeGamepadId = 'pad-1';

    const ctx = {
      dom: {
        overlay: { classList: createClassList(), focus: () => {} },
        controllerSelectModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
        },
        controllerSelectClose: { addEventListener: () => {} },
        controllerSelectHint: { textContent: '' },
        controllerSelectStatus: { textContent: '', classList: createClassList() },
        controllerSelectList: {
          innerHTML: '',
          appendChild: () => {},
        },
        controllerSelectSave: { addEventListener: () => {} },
      },
      state,
    };

    const modal = createControllerSelectModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    modal.openControllerSelectModal();
    assert.equal(state.controllerDeviceSelectedIndex, 0);

    modal.handleControllerSelectKeydown({
      key: 'ArrowDown',
      preventDefault: () => {},
    } as KeyboardEvent);
    assert.equal(state.controllerDeviceSelectedIndex, 1);

    modal.updateDevices();

    assert.equal(state.controllerDeviceSelectedIndex, 1);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('controller select modal does not rerender unchanged device snapshots every poll', () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  let appendCount = 0;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerPreference: async () => {},
        notifyOverlayModalClosed: () => {},
      },
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({
        className: '',
        textContent: '',
        classList: createClassList(),
        appendChild: () => {},
        addEventListener: () => {},
      }),
    },
  });

  try {
    const state = createRendererState();
    state.controllerConfig = {
      enabled: true,
      preferredGamepadId: 'pad-1',
      preferredGamepadLabel: 'pad-1',
      smoothScroll: true,
      scrollPixelsPerSecond: 960,
      horizontalJumpPixels: 160,
      stickDeadzone: 0.2,
      triggerInputMode: 'auto',
      triggerDeadzone: 0.5,
      repeatDelayMs: 220,
      repeatIntervalMs: 80,
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
        toggleLookup: 'buttonSouth',
        closeLookup: 'buttonEast',
        toggleKeyboardOnlyMode: 'buttonNorth',
        mineCard: 'buttonWest',
        quitMpv: 'select',
        previousAudio: 'none',
        nextAudio: 'rightShoulder',
        playCurrentAudio: 'leftShoulder',
        toggleMpvPause: 'leftStickPress',
        leftStickHorizontal: 'leftStickX',
        leftStickVertical: 'leftStickY',
        rightStickHorizontal: 'rightStickX',
        rightStickVertical: 'rightStickY',
      },
    };
    state.connectedGamepads = [
      { id: 'pad-1', index: 0, mapping: 'standard', connected: true },
      { id: 'pad-2', index: 1, mapping: 'standard', connected: true },
    ];
    state.activeGamepadId = 'pad-1';

    const ctx = {
      dom: {
        overlay: { classList: createClassList(), focus: () => {} },
        controllerSelectModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
        },
        controllerSelectClose: { addEventListener: () => {} },
        controllerSelectHint: { textContent: '' },
        controllerSelectStatus: { textContent: '', classList: createClassList() },
        controllerSelectList: {
          innerHTML: '',
          appendChild: () => {
            appendCount += 1;
          },
        },
        controllerSelectSave: { addEventListener: () => {} },
      },
      state,
    };

    const modal = createControllerSelectModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
    });

    modal.openControllerSelectModal();
    const initialAppendCount = appendCount;

    modal.updateDevices();

    assert.equal(appendCount, initialAppendCount);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});
