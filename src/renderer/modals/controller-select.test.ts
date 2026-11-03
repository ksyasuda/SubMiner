import assert from 'node:assert/strict';
import test from 'node:test';

import { createRendererState } from '../state.js';
import {
  createControllerSelectModal,
  findMostRecentlyUsedIndex,
  shortControllerName,
} from './controller-select.js';
import { getDefaultControllerReviewBindings } from '../../shared/controller-review';
import { getDefaultControllerFieldGroupingBindings } from '../../shared/controller-field-grouping';

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
    faceButtonLabels: 'xbox',
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
      holdSecondarySub: { kind: 'none' },
      previousEntry: { kind: 'none' },
      nextEntry: { kind: 'none' },
      previousSubtitle: { kind: 'none' },
      nextSubtitle: { kind: 'none' },
      replaySubtitle: { kind: 'none' },
      playNextSubtitle: { kind: 'none' },
      leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
      leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
      rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
      rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
    },
    reviewBindings: getDefaultControllerReviewBindings(),
    fieldGroupingBindings: getDefaultControllerFieldGroupingBindings(),
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
    controllerSelectFaceButtonStyle: createFakeElement(),
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
          faceButtonLabels: state.controllerConfig.faceButtonLabels,
          buttonIndices: state.controllerConfig.buttonIndices,
          bindings: {
            ...state.controllerConfig.bindings,
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
          reviewBindings: state.controllerConfig.reviewBindings,
          fieldGroupingBindings: state.controllerConfig.fieldGroupingBindings,
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

test('controller select modal locks editing until this window sees a controller', async () => {
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
    // A freshly opened window sees no pad until a button is pressed in it.
    state.connectedGamepads = [];
    state.activeGamepadId = null;
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });
    modal.wireDomEvents();
    modal.openControllerSelectModal();

    const press = 'Press any button on your controller to edit its bindings.';
    assert.equal(dom.controllerSelectStatus.textContent, press);
    assert.equal(dom.controllerConfigList.classList.contains('controller-config-locked'), true);
    assert.equal(dom.controllerSelectFaceButtonStyle.disabled, true);

    // Learn and Clear do nothing yet.
    const row = findReviewRow(dom.controllerConfigList, 'Toggle Lookup');
    row.children[1].children[0].dispatch('click');
    state.controllerRawButtons = Array.from({ length: 12 }, (_, index) => ({
      value: index === 11 ? 1 : 0,
      pressed: index === 11,
      touched: index === 11,
    }));
    modal.updateDevices();
    await Promise.resolve();
    assert.deepEqual(saved, []);
    assert.equal(dom.controllerSelectStatus.textContent, press, 'kept across frames');

    // The press makes the pad visible: editing unlocks for it.
    state.connectedGamepads = [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }];
    state.activeGamepadId = 'pad-1';
    modal.updateDevices();
    assert.equal(dom.controllerConfigList.classList.contains('controller-config-locked'), false);
    assert.equal(dom.controllerSelectFaceButtonStyle.disabled, false);
    assert.equal(
      dom.controllerSelectStatus.textContent,
      'Choose a controller or click Learn to remap an action.',
    );
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

function findByClass(root: any, className: string, text?: string): any {
  for (const child of root.children ?? []) {
    const classes = String(child.className ?? '').split(' ');
    if (classes.includes(className) && (text === undefined || child.textContent === text)) {
      return child;
    }
    const nested = findByClass(child, className, text);
    if (nested) return nested;
  }
  return null;
}

function findReviewRow(root: any, label: string): any {
  for (const child of root.children ?? []) {
    if (String(child.className).split(' ').includes('controller-config-row')) {
      if (child.children[0]?.textContent === label) return child;
    }
  }
  return null;
}

test('controller select modal learns a review combo and saves it to the controller profile', async () => {
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

    const header = findByClass(dom.controllerConfigList, 'controller-config-group-toggle');
    assert.ok(header, 'review group header');
    header.dispatch('click');
    const row = findReviewRow(dom.controllerConfigList, 'Start +0.1s');
    assert.ok(row, 'Start +0.1s row');
    row.children[1].children[0].dispatch('click');

    const buttons = (pressed: number[]) =>
      Array.from({ length: 17 }, (_, index) => ({
        value: pressed.includes(index) ? 1 : 0,
        pressed: pressed.includes(index),
        touched: pressed.includes(index),
      }));

    // Hold L2, then press A.
    state.controllerRawButtons = buttons([6]);
    modal.updateDevices();
    assert.equal(
      findReviewRow(dom.controllerConfigList, 'Start +0.1s').children[1].children[0].textContent,
      'L2 + …',
    );
    state.controllerRawButtons = buttons([6, 0]);
    modal.updateDevices();

    await Promise.resolve();

    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          reviewBindings: {
            startForward: {
              kind: 'combo',
              press: { kind: 'button', buttonIndex: 0 },
              hold: { kind: 'button', buttonIndex: 6 },
            },
          },
        },
      },
    });
    assert.deepEqual(state.controllerConfig?.profiles['pad-1']?.reviewBindings.startForward, {
      kind: 'combo',
      press: { kind: 'button', buttonIndex: 0 },
      hold: { kind: 'button', buttonIndex: 6 },
    });
  } finally {
    domHandle.restore();
  }
});

function findAllByClass(root: any, className: string, found: any[] = []): any[] {
  for (const child of root.children ?? []) {
    if (
      String(child.className ?? '')
        .split(' ')
        .includes(className)
    )
      found.push(child);
    findAllByClass(child, className, found);
  }
  return found;
}

test('controller select modal learns a duplicate card button and saves it to the profile', async () => {
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

    const headers = findAllByClass(dom.controllerConfigList, 'controller-config-group-toggle');
    assert.equal(headers.length, 2, 'review and duplicate card groups');
    headers[1].dispatch('click');
    const row = findReviewRow(dom.controllerConfigList, 'Back');
    assert.ok(row, 'Back row');
    row.children[1].children[0].dispatch('click');

    const buttons = (pressed: number[]) =>
      Array.from({ length: 17 }, (_, index) => ({
        value: pressed.includes(index) ? 1 : 0,
        pressed: pressed.includes(index),
        touched: pressed.includes(index),
      }));

    // Press and release Select.
    state.controllerRawButtons = buttons([8]);
    modal.updateDevices();
    state.controllerRawButtons = buttons([]);
    modal.updateDevices();

    await Promise.resolve();

    const back = { kind: 'combo', press: { kind: 'button', buttonIndex: 8 }, hold: null };
    assert.deepEqual(saved.at(-1), {
      profiles: { 'pad-1': { fieldGroupingBindings: { back } } },
    });
    const profile = state.controllerConfig?.profiles['pad-1'];
    assert.deepEqual(profile?.fieldGroupingBindings.back, back);
    // Other dialogs keep their layout.
    assert.deepEqual(profile?.reviewBindings, state.controllerConfig?.reviewBindings);
  } finally {
    domHandle.restore();
  }
});

test('controller select modal previews the saved controller, then edits it once seen', async () => {
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
    // A fresh Alt+C window has not seen the pad yet (no button pressed in it).
    state.connectedGamepads = [];
    state.activeGamepadId = null;
    const config = state.controllerConfig!;
    config.profiles = {
      'pad-1': {
        label: 'pad-1',
        faceButtonLabels: config.faceButtonLabels,
        buttonIndices: config.buttonIndices,
        bindings: config.bindings,
        fieldGroupingBindings: config.fieldGroupingBindings,
        reviewBindings: {
          ...config.reviewBindings,
          startBack: {
            kind: 'combo',
            press: { kind: 'button', buttonIndex: 15 },
            hold: { kind: 'button', buttonIndex: 6 },
          },
        },
      },
    };
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });
    modal.wireDomEvents();
    modal.openControllerSelectModal();

    // The preview is the saved controller's layout, labelled as such.
    assert.equal(dom.controllerSelectSummary.textContent, 'Showing pad-1 (saved).');
    findByClass(dom.controllerConfigList, 'controller-config-group-toggle').dispatch('click');
    const startBack = () => findReviewRow(dom.controllerConfigList, 'Start \u22120.1s');
    assert.equal(startBack().children[1].children[0].textContent, 'L2 + D-pad \u25b6');
    const press = 'Press any button on your controller to edit its bindings.';
    assert.equal(dom.controllerSelectStatus.textContent, press);
    // Controller polling keeps publishing snapshots while no pad is visible.
    modal.updateDevices();
    modal.updateDevices();
    assert.equal(dom.controllerSelectStatus.textContent, press, 'the prompt stays');

    // Reset is locked until the window sees the pad.
    startBack().children[1].children[1].dispatch('click');
    await Promise.resolve();
    assert.deepEqual(saved, []);

    state.connectedGamepads = [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }];
    state.activeGamepadId = 'pad-1';
    modal.updateDevices();
    assert.equal(
      dom.controllerSelectStatus.textContent,
      'Choose a controller or click Learn to remap an action.',
    );
    startBack().children[1].children[1].dispatch('click');
    await Promise.resolve();
    assert.deepEqual(saved.at(-1), {
      profiles: {
        'pad-1': {
          reviewBindings: {
            startBack: {
              kind: 'combo',
              press: { kind: 'dpad', direction: 'left' },
              hold: { kind: 'button', buttonIndex: 4 },
            },
          },
        },
      },
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    assert.equal(dom.controllerSelectStatus.textContent, 'Start \u22120.1s reset.');
    modal.updateDevices();
    assert.equal(
      dom.controllerSelectStatus.textContent,
      'Start \u22120.1s reset.',
      'the reset message stays',
    );

    // The pad disappears, then comes back.
    state.connectedGamepads = [];
    state.activeGamepadId = null;
    modal.updateDevices();
    assert.equal(dom.controllerSelectStatus.textContent, press);
    state.connectedGamepads = [{ id: 'pad-1', index: 0, mapping: 'standard', connected: true }];
    state.activeGamepadId = 'pad-1';
    modal.updateDevices();
    assert.equal(
      dom.controllerSelectStatus.textContent,
      'Choose a controller or click Learn to remap an action.',
    );
  } finally {
    domHandle.restore();
  }
});

test('controller select modal previews the controller the overlay reports in use', async () => {
  const domHandle = installFakeDom();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: {
        saveControllerConfig: async () => {},
        notifyOverlayModalClosed: () => {},
        getActiveControllerId: async () => 'DualSense Wireless Controller (STANDARD GAMEPAD)',
      },
    },
  });

  try {
    const { state, dom } = buildContext();
    state.connectedGamepads = [];
    state.activeGamepadId = null;
    const config = state.controllerConfig!;
    // The saved controller is pad-1, but the overlay is reading the DualSense.
    config.profiles = {
      'DualSense Wireless Controller (STANDARD GAMEPAD)': {
        label: 'DualSense',
        faceButtonLabels: 'playstation',
        buttonIndices: config.buttonIndices,
        bindings: { ...config.bindings, mineCard: { kind: 'button', buttonIndex: 3 } },
        reviewBindings: config.reviewBindings,
        fieldGroupingBindings: config.fieldGroupingBindings,
      },
    };
    const modal = createControllerSelectModal({ state, dom } as never, {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {},
      notifyControllerDisabled: () => {},
    });
    modal.wireDomEvents();
    modal.openControllerSelectModal();
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.equal(
      dom.controllerSelectSummary.textContent,
      'Showing DualSense Wireless Controller (in use).',
    );
    assert.equal(
      findReviewRow(dom.controllerConfigList, 'Mine Card').children[1].children[0].textContent,
      '\u25b3',
    );
    assert.equal(dom.controllerSelectFaceButtonStyle.value, 'playstation');
  } finally {
    domHandle.restore();
  }
});

test('controller select modal selects the controller you pressed when several appear', () => {
  const domHandle = installFakeDom();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      focus: () => {},
      electronAPI: { saveControllerConfig: async () => {}, notifyOverlayModalClosed: () => {} },
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

    // One press reveals every pad; pad-2 has the newest input, pad-1 is the saved one.
    state.connectedGamepads = [
      { id: 'pad-1', index: 0, mapping: 'standard', connected: true, timestamp: 100 },
      { id: 'pad-2', index: 1, mapping: 'standard', connected: true, timestamp: 500 },
    ];
    state.activeGamepadId = 'pad-1';
    modal.updateDevices();

    assert.equal(state.controllerDeviceSelectedIndex, 1);
    assert.equal(state.controllerSelectGamepadId, 'pad-2', 'this window then reads pad-2');

    modal.closeControllerSelectModal();
    assert.equal(state.controllerSelectGamepadId, null);
  } finally {
    domHandle.restore();
  }
});

test('controller select modal shows and saves the button style per controller', async () => {
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

    const stylePicker = dom.controllerSelectFaceButtonStyle;
    assert.deepEqual(
      stylePicker.children.map((option: any) => option.value),
      ['xbox', 'nintendo', 'playstation'],
    );
    assert.equal(stylePicker.value, 'xbox');

    stylePicker.value = 'playstation';
    stylePicker.dispatch('change');
    await new Promise((resolve) => setTimeout(resolve, 0));

    assert.deepEqual(saved.at(-1), {
      profiles: { 'pad-1': { faceButtonLabels: 'playstation' } },
    });
    assert.equal(state.controllerConfig?.profiles['pad-1']?.faceButtonLabels, 'playstation');
    assert.equal(state.controllerConfig?.faceButtonLabels, 'xbox', 'other controllers keep theirs');
    assert.equal(dom.controllerSelectStatus.textContent, 'Button style: PlayStation.');
    // Toggle Lookup is on the bottom face button.
    assert.equal(
      findReviewRow(dom.controllerConfigList, 'Toggle Lookup').children[1].children[0].textContent,
      '╳',
    );

    // The other controller still uses the default style.
    dom.controllerSelectPicker.value = 'pad-2:1';
    dom.controllerSelectPicker.dispatch('change');
    assert.equal(stylePicker.value, 'xbox');
  } finally {
    domHandle.restore();
  }
});

test('shortControllerName drops the driver details in brackets', () => {
  assert.equal(
    shortControllerName('Xbox Wireless Controller (STANDARD GAMEPAD Vendor: 045e Product: 0b13)'),
    'Xbox Wireless Controller',
  );
  assert.equal(shortControllerName('pad-1'), 'pad-1');
  assert.equal(shortControllerName('(only details)'), '(only details)');
});

test('findMostRecentlyUsedIndex picks the newest input and ignores missing timestamps', () => {
  assert.equal(findMostRecentlyUsedIndex([{ timestamp: 100 }, { timestamp: 500 }, {}]), 1);
  assert.equal(findMostRecentlyUsedIndex([{ timestamp: 900 }, { timestamp: Number.NaN }]), 0);
  assert.equal(findMostRecentlyUsedIndex([{}, { timestamp: 0 }]), -1);
  assert.equal(findMostRecentlyUsedIndex([]), -1);
});
