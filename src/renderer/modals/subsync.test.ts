import test from 'node:test';
import assert from 'node:assert/strict';

import { createSubsyncModal } from './subsync.js';

type Listener = () => void;

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => {
      for (const token of tokens) classes.add(token);
    },
    remove: (...tokens: string[]) => {
      for (const token of tokens) classes.delete(token);
    },
    toggle: (token: string, force?: boolean) => {
      if (force === undefined) {
        if (classes.has(token)) classes.delete(token);
        else classes.add(token);
        return classes.has(token);
      }
      if (force) classes.add(token);
      else classes.delete(token);
      return force;
    },
    contains: (token: string) => classes.has(token),
  };
}

function createEventTarget() {
  const listeners = new Map<string, Listener[]>();
  return {
    addEventListener: (event: string, listener: Listener) => {
      const existing = listeners.get(event) ?? [];
      existing.push(listener);
      listeners.set(event, existing);
    },
    dispatch: (event: string) => {
      for (const listener of listeners.get(event) ?? []) {
        listener();
      }
    },
  };
}

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

function createTestHarness(runSubsyncManual: () => Promise<{ ok: boolean; message: string }>) {
  const overlayClassList = createClassList();
  const modalClassList = createClassList();
  const statusClassList = createClassList();
  const sourceLabelClassList = createClassList();
  const runButtonEvents = createEventTarget();
  const closeButtonEvents = createEventTarget();
  const engineAlassEvents = createEventTarget();
  const engineFfsubsyncEvents = createEventTarget();

  const sourceOptions: Array<{ value: string; textContent: string }> = [];

  const runButton = {
    disabled: false,
    addEventListener: runButtonEvents.addEventListener,
    dispatch: runButtonEvents.dispatch,
  };

  const closeButton = {
    addEventListener: closeButtonEvents.addEventListener,
    dispatch: closeButtonEvents.dispatch,
  };

  const subsyncEngineAlass = {
    checked: false,
    addEventListener: engineAlassEvents.addEventListener,
    dispatch: engineAlassEvents.dispatch,
  };

  const subsyncEngineFfsubsync = {
    checked: false,
    disabled: false,
    addEventListener: engineFfsubsyncEvents.addEventListener,
    dispatch: engineFfsubsyncEvents.dispatch,
  };

  const sourceSelect = {
    innerHTML: '',
    value: '',
    disabled: false,
    appendChild: (option: { value: string; textContent: string }) => {
      sourceOptions.push(option);
      if (!sourceSelect.value) {
        sourceSelect.value = option.value;
      }
      return option;
    },
  };

  let notifyClosedCalls = 0;
  let notifyOpenedCalls = 0;

  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousDocument = (globalThis as { document?: unknown }).document;

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        runSubsyncManual,
        notifyOverlayModalOpened: () => {
          notifyOpenedCalls += 1;
        },
        notifyOverlayModalClosed: () => {
          notifyClosedCalls += 1;
        },
      },
    },
  });

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => ({ value: '', textContent: '' }),
    },
  });

  const ctx = {
    dom: {
      overlay: { classList: overlayClassList },
      subsyncModal: {
        classList: modalClassList,
        setAttribute: () => {},
      },
      subsyncCloseButton: closeButton,
      subsyncEngineAlass,
      subsyncEngineFfsubsync,
      subsyncSourceLabel: { classList: sourceLabelClassList },
      subsyncSourceSelect: sourceSelect,
      subsyncRunButton: runButton,
      subsyncStatus: {
        textContent: '',
        classList: statusClassList,
      },
    },
    state: {
      subsyncModalOpen: false,
      subsyncSourceTracks: [],
      subsyncSubmitting: false,
      isOverSubtitle: false,
    },
  };

  const modal = createSubsyncModal(ctx as never, {
    modalStateReader: {
      isAnyModalOpen: () => false,
    },
    syncSettingsModalSubtitleSuppression: () => {},
  });

  return {
    ctx,
    modal,
    runButton,
    statusClassList,
    getNotifyClosedCalls: () => notifyClosedCalls,
    getNotifyOpenedCalls: () => notifyOpenedCalls,
    restoreGlobals: () => {
      Object.defineProperty(globalThis, 'window', {
        configurable: true,
        value: previousWindow,
      });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        value: previousDocument,
      });
    },
  };
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

test('manual subsync failure closes during run, then reopens modal with error', async () => {
  const deferred = createDeferred<{ ok: boolean; message: string }>();
  const harness = createTestHarness(async () => deferred.promise);

  try {
    harness.modal.wireDomEvents();
    harness.modal.openSubsyncModal({
      sourceTracks: [{ id: 2, label: 'External #2 - eng' }],
      ffsubsyncAvailable: true,
    });

    harness.runButton.dispatch('click');
    await Promise.resolve();

    assert.equal(harness.ctx.state.subsyncModalOpen, false);
    assert.equal(harness.getNotifyClosedCalls(), 1);
    assert.equal(harness.getNotifyOpenedCalls(), 0);

    deferred.resolve({
      ok: false,
      message: 'alass synchronization failed: code=1 stderr: invalid subtitle format',
    });
    await flushMicrotasks();

    assert.equal(harness.ctx.state.subsyncModalOpen, true);
    assert.equal(
      harness.ctx.dom.subsyncStatus.textContent,
      'alass synchronization failed: code=1 stderr: invalid subtitle format',
    );
    assert.equal(harness.statusClassList.contains('error'), true);
    assert.equal(harness.ctx.dom.subsyncRunButton.disabled, false);
    assert.equal(harness.ctx.dom.subsyncEngineAlass.checked, true);
    assert.equal(harness.ctx.dom.subsyncSourceSelect.value, '2');
    assert.equal(harness.getNotifyClosedCalls(), 1);
    assert.equal(harness.getNotifyOpenedCalls(), 1);
  } finally {
    harness.restoreGlobals();
  }
});

test('subsync modal disables ffsubsync when payload marks it unavailable', () => {
  const harness = createTestHarness(async () => ({ ok: true, message: 'ok' }));

  try {
    harness.modal.openSubsyncModal({
      sourceTracks: [{ id: 2, label: 'External #2 - eng' }],
      ffsubsyncAvailable: false,
    });

    assert.equal(harness.ctx.dom.subsyncEngineAlass.checked, true);
    assert.equal(harness.ctx.dom.subsyncEngineFfsubsync.checked, false);
    assert.equal(harness.ctx.dom.subsyncEngineFfsubsync.disabled, true);
    assert.equal(harness.ctx.dom.subsyncStatus.textContent, 'Choose alass source, then run.');
  } finally {
    harness.restoreGlobals();
  }
});
