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

function createFakeElement() {
  const attributes = new Map<string, string>();
  const el = {
    className: '',
    textContent: '',
    _innerHTML: '',
    value: '',
    disabled: false,
    selected: false,
    type: '',
    children: [] as any[],
    listeners: new Map<string, Array<(e?: any) => void>>(),
    classList: createClassList(),
    appendChild(child: any) {
      this.children.push(child);
      return child;
    },
    addEventListener(type: string, listener: (e?: any) => void) {
      const existing = this.listeners.get(type) ?? [];
      existing.push(listener);
      this.listeners.set(type, existing);
    },
    dispatch(type: string) {
      const fakeEvent = { stopPropagation: () => {}, preventDefault: () => {} };
      for (const listener of this.listeners.get(type) ?? []) {
        listener(fakeEvent);
      }
    },
    setAttribute(name: string, value: string) {
      attributes.set(name, value);
    },
    getAttribute(name: string) {
      return attributes.get(name) ?? null;
    },
    querySelector(selector: string) {
      const match = selector.match(/^\[data-testid="(.+)"\]$/);
      if (!match) return null;
      const testId = match[1];
      for (const child of el.children) {
        if (
          typeof child.getAttribute === 'function' &&
          child.getAttribute('data-testid') === testId
        ) {
          return child;
        }
        if (typeof child.querySelector === 'function') {
          const nested = child.querySelector(selector);
          if (nested) return nested;
        }
      }
      return null;
    },
    focus: () => {},
  };
  Object.defineProperty(el, 'innerHTML', {
    get() {
      return el._innerHTML;
    },
    set(v: string) {
      el._innerHTML = v;
      if (v === '') el.children.length = 0;
    },
  });
  return el;
}

function installFakeDom() {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createFakeElement(),
    },
  });

  return {
    restore: () => {
      Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument,
      });
    },
  };
}

function buildContext() {
  const state = createRendererState();
  state.controllerConfig = {
    enabled: true,
    preferredGamepadId: 'pad-1',
    preferredGamepadLabel: 'pad-1',
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
    profiles: {},
  };
  state.connectedGamepads = [
    { id: 'pad-1', index: 0, mapping: 'standard', connected: true },
    { id: 'pad-2', index: 1, mapping: 'standard', connected: true },
  ];
  state.activeGamepadId = 'pad-1';

  const dom = {
    overlay: { classList: createClassList(), focus: () => {} },
    controllerSelectModal: { classList: createClassList(['hidden']), setAttribute: () => {} },
    controllerSelectClose: createFakeElement(),
    controllerSelectPicker: createFakeElement(),
    controllerSelectSummary: createFakeElement(),
    controllerConfigList: createFakeElement(),
    controllerSelectStatus: { textContent: '', classList: createClassList() },
    controllerSelectSave: createFakeElement(),
  };

  return { state, dom };
}

test('controller select modal saves preferred controller from dropdown selection', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();
    state.controllerDeviceSelectedIndex = 1;

    await modal.handleControllerSelectKeydown({
      key: 'Enter',
      preventDefault: () => {},
    } as KeyboardEvent);
    await Promise.resolve();

    assert.deepEqual(saved, [
      {
        preferredGamepadId: 'pad-2',
        preferredGamepadLabel: 'pad-2',
      },
    ]);
  } finally {
    domHandle.restore();
  }
});

test('controller select modal learn mode captures fresh button input and persists binding', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    // In the new compact list layout, children are:
    // [0] group header, [1] first binding row, [2] second binding row, ...
    // Click the row to expand the inline edit panel
    const firstRow = dom.controllerConfigList.children[1];
    firstRow.dispatch('click');

    // After expanding, the edit panel is inserted after the row:
    // [0] group header, [1] row, [2] edit panel, [3] next row, ...
    const editPanel = dom.controllerConfigList.children[2];
    // editPanel > inner > actions > learnButton
    const inner = editPanel.children[0];
    const actions = inner.children[1];
    const learnButton = actions.children[0];
    learnButton.dispatch('click');

    state.controllerRawButtons = Array.from({ length: 12 }, () => ({
      value: 0,
      pressed: false,
      touched: false,
    }));
    state.controllerRawButtons[11] = { value: 1, pressed: true, touched: true };
    modal.updateDevices();

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          label: 'pad-1',
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
        },
      },
    });
    assert.deepEqual(state.controllerConfig?.profiles['pad-1']?.bindings.toggleLookup, {
      kind: 'button',
      buttonIndex: 11,
    });
  } finally {
    domHandle.restore();
  }
});

test('controller select modal reset control stores the default binding in the selected profile', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    if (state.controllerConfig) {
      state.controllerConfig.profiles = {
        'pad-1': {
          label: 'pad-1',
          buttonIndices: state.controllerConfig.buttonIndices,
          bindings: {
            ...state.controllerConfig.bindings,
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
        },
      };
    }
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const firstRow = dom.controllerConfigList.children[1];
    const right = firstRow.children[1];
    const resetButton = right.children[1];
    resetButton.dispatch('click');

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          label: 'pad-1',
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 0 },
          },
        },
      },
    });
    assert.deepEqual(state.controllerConfig?.profiles['pad-1']?.bindings.toggleLookup, {
      kind: 'button',
      buttonIndex: 0,
    });
  } finally {
    domHandle.restore();
  }
});

test('controller select modal binding badge starts learn mode and persists binding', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const firstRow = dom.controllerConfigList.children[1];
    const right = firstRow.children[1];
    const badge = right.children[0];
    badge.dispatch('click');

    state.controllerRawButtons = Array.from({ length: 12 }, () => ({
      value: 0,
      pressed: false,
      touched: false,
    }));
    state.controllerRawButtons[11] = { value: 1, pressed: true, touched: true };
    modal.updateDevices();

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          label: 'pad-1',
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
        },
      },
    });
  } finally {
    domHandle.restore();
  }
});

test('controller select modal learn mode falls back to global bindings without a controller', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    state.connectedGamepads = [];
    state.activeGamepadId = null;
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const firstRow = dom.controllerConfigList.children[1];
    firstRow.dispatch('click');
    const editPanel = dom.controllerConfigList.children[2];
    const learnButton = editPanel.children[0].children[1].children[0];
    learnButton.dispatch('click');

    state.controllerRawButtons = Array.from({ length: 12 }, () => ({
      value: 0,
      pressed: false,
      touched: false,
    }));
    state.controllerRawButtons[11] = { value: 1, pressed: true, touched: true };
    modal.updateDevices();

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      bindings: {
        toggleLookup: { kind: 'button', buttonIndex: 11 },
      },
    });
    assert.deepEqual(state.controllerConfig?.bindings.toggleLookup, {
      kind: 'button',
      buttonIndex: 11,
    });
  } finally {
    domHandle.restore();
  }
});

test('controller select modal edit control starts learn mode and persists binding', async () => {
  const domHandle = installFakeDom();
  const saved: unknown[] = [];

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async (update: unknown) => {
          saved.push(update);
        },
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const firstRow = dom.controllerConfigList.children[1];
    const right = firstRow.children[1];
    const editButton = right.children[2];
    editButton.dispatch('click');

    state.controllerRawButtons = Array.from({ length: 12 }, () => ({
      value: 0,
      pressed: false,
      touched: false,
    }));
    state.controllerRawButtons[11] = { value: 1, pressed: true, touched: true };
    modal.updateDevices();

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          label: 'pad-1',
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
        },
      },
    });
  } finally {
    domHandle.restore();
  }
});

test('controller select modal stays closed and notifies when controller support is disabled', async () => {
  const domHandle = installFakeDom();
  let disabledNotices = 0;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async () => {},
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    if (state.controllerConfig) state.controllerConfig.enabled = false;
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {
        disabledNotices += 1;
      },
    });

    assert.equal(modal.openControllerSelectModal(), false);

    assert.equal(state.controllerSelectModalOpen, false);
    assert.equal(dom.controllerSelectModal.classList.contains('hidden'), true);
    assert.equal(disabledNotices, 1);
  } finally {
    domHandle.restore();
  }
});

test('controller select modal uses unique picker values for duplicate controller ids', async () => {
  const domHandle = installFakeDom();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async () => {},
        notifyOverlayModalClosed: () => {},
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    state.connectedGamepads = [
      { id: 'same-pad', index: 0, mapping: 'standard', connected: true },
      { id: 'same-pad', index: 1, mapping: 'standard', connected: true },
    ];
    state.activeGamepadId = 'same-pad';

    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });

    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const [firstOption, secondOption] = dom.controllerSelectPicker.children;
    assert.notEqual(firstOption.value, secondOption.value);

    dom.controllerSelectPicker.value = secondOption.value;
    dom.controllerSelectPicker.dispatch('change');

    assert.equal(state.controllerDeviceSelectedIndex, 1);
  } finally {
    domHandle.restore();
  }
});
