import assert from 'node:assert/strict';
import test from 'node:test';

import type { ElectronAPI, RuntimeOptionState } from '../../types';
import { createRendererState } from '../state.js';
import { createRuntimeOptionsModal } from './runtime-options.js';

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
        if (tokens.has(entry)) {
          tokens.delete(entry);
          return false;
        }
        tokens.add(entry);
        return true;
      }
      if (force) tokens.add(entry);
      else tokens.delete(entry);
      return force;
    },
    contains: (entry: string) => tokens.has(entry),
  };
}

function createElementStub() {
  return {
    className: '',
    textContent: '',
    title: '',
    classList: createClassList(),
    appendChild: () => {},
    addEventListener: () => {},
  };
}

function createRuntimeOptionsListStub() {
  return {
    innerHTML: '',
    appendChild: () => {},
    querySelector: () => null,
  };
}

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, resolve, reject };
}

function flushAsyncWork(): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, 0);
  });
}

function withRuntimeOptionsModal(
  getRuntimeOptions: () => Promise<RuntimeOptionState[]>,
  run: (input: {
    modal: ReturnType<typeof createRuntimeOptionsModal>;
    state: ReturnType<typeof createRendererState>;
    overlayClassList: ReturnType<typeof createClassList>;
    modalClassList: ReturnType<typeof createClassList>;
    statusNode: {
      textContent: string;
      classList: ReturnType<typeof createClassList>;
    };
    syncCalls: string[];
  }) => Promise<void> | void,
): Promise<void> {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const statusNode = {
    textContent: '',
    classList: createClassList(),
  };
  const overlayClassList = createClassList();
  const modalClassList = createClassList(['hidden']);
  const syncCalls: string[] = [];
  const state = createRendererState();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: {
      electronAPI: {
        getRuntimeOptions,
        setRuntimeOptionValue: async () => ({ ok: true }),
        notifyOverlayModalClosed: () => {},
      } satisfies Pick<
        ElectronAPI,
        'getRuntimeOptions' | 'setRuntimeOptionValue' | 'notifyOverlayModalClosed'
      >,
    },
  });

  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: {
      createElement: () => createElementStub(),
    },
  });

  const modal = createRuntimeOptionsModal(
    {
      dom: {
        overlay: { classList: overlayClassList },
        runtimeOptionsModal: {
          classList: modalClassList,
          setAttribute: () => {},
        },
        runtimeOptionsClose: {
          addEventListener: () => {},
        },
        runtimeOptionsList: createRuntimeOptionsListStub(),
        runtimeOptionsStatus: statusNode,
      },
      state,
    } as never,
    {
      modalStateReader: { isAnyModalOpen: () => false },
      syncSettingsModalSubtitleSuppression: () => {
        syncCalls.push('sync');
      },
    },
  );

  return Promise.resolve()
    .then(() =>
      run({
        modal,
        state,
        overlayClassList,
        modalClassList,
        statusNode,
        syncCalls,
      }),
    )
    .finally(() => {
      Object.defineProperty(globalThis, 'window', {
        configurable: true,
        writable: true,
        value: previousWindow,
      });
      Object.defineProperty(globalThis, 'document', {
        configurable: true,
        writable: true,
        value: previousDocument,
      });
    });
}

test('openRuntimeOptionsModal shows loading shell before runtime options resolve', async () => {
  const deferred = createDeferred<RuntimeOptionState[]>();

  await withRuntimeOptionsModal(() => deferred.promise, async (input) => {
    input.modal.openRuntimeOptionsModal();

    assert.equal(input.state.runtimeOptionsModalOpen, true);
    assert.equal(input.overlayClassList.contains('interactive'), true);
    assert.equal(input.modalClassList.contains('hidden'), false);
    assert.equal(input.statusNode.textContent, 'Loading runtime options...');
    assert.deepEqual(input.syncCalls, ['sync']);

    deferred.resolve([
      {
        id: 'anki.autoUpdateNewCards',
        label: 'Auto-update new cards',
        scope: 'ankiConnect',
        valueType: 'boolean',
        value: true,
        allowedValues: [true, false],
        requiresRestart: false,
      },
    ]);
    await flushAsyncWork();

    assert.equal(
      input.statusNode.textContent,
      'Use arrow keys. Click value to cycle. Enter or double-click to apply.',
    );
    assert.equal(input.statusNode.classList.contains('error'), false);
  });
});

test('openRuntimeOptionsModal keeps modal visible when loading fails', async () => {
  const deferred = createDeferred<RuntimeOptionState[]>();

  await withRuntimeOptionsModal(() => deferred.promise, async (input) => {
    input.modal.openRuntimeOptionsModal();
    deferred.reject(new Error('boom'));
    await flushAsyncWork();

    assert.equal(input.state.runtimeOptionsModalOpen, true);
    assert.equal(input.overlayClassList.contains('interactive'), true);
    assert.equal(input.modalClassList.contains('hidden'), false);
    assert.equal(input.statusNode.textContent, 'Failed to load runtime options');
    assert.equal(input.statusNode.classList.contains('error'), true);
  });
});
