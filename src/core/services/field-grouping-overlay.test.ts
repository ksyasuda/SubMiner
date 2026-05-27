import test from 'node:test';
import assert from 'node:assert/strict';
import { KikuFieldGroupingChoice } from '../../types';
import { createFieldGroupingOverlayRuntime } from './field-grouping-overlay';

test('createFieldGroupingOverlayRuntime sends overlay messages and sets restore flag', () => {
  const sent: unknown[][] = [];
  let visible = false;
  const restore = new Set<'runtime-options' | 'subsync'>();

  const runtime = createFieldGroupingOverlayRuntime<'runtime-options' | 'subsync'>({
    getMainWindow: () => ({
      isDestroyed: () => false,
      webContents: {
        isLoading: () => false,
        send: (...args: unknown[]) => {
          sent.push(args);
        },
      },
    }),
    getVisibleOverlayVisible: () => visible,
    setVisibleOverlayVisible: (next) => {
      visible = next;
    },
    getResolver: () => null,
    setResolver: () => {},
    getRestoreVisibleOverlayOnModalClose: () => restore,
  });

  const ok = runtime.sendToVisibleOverlay('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(ok, true);
  assert.equal(visible, true);
  assert.equal(restore.has('runtime-options'), true);
  assert.deepEqual(sent, [['runtime-options:open']]);
});

test('createFieldGroupingOverlayRuntime callback cancels when send fails', async () => {
  let resolver: ((choice: KikuFieldGroupingChoice) => void) | null = null;
  const runtime = createFieldGroupingOverlayRuntime<'runtime-options' | 'subsync'>({
    getMainWindow: () => null,
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: () => resolver,
    setResolver: (next: ((choice: KikuFieldGroupingChoice) => void) | null) => {
      resolver = next;
    },
    getRestoreVisibleOverlayOnModalClose: () => new Set<'runtime-options' | 'subsync'>(),
  });

  const callback = runtime.createFieldGroupingCallback();
  const result = await callback({
    original: {
      noteId: 1,
      expression: 'a',
      sentencePreview: 'a',
      hasAudio: false,
      hasImage: false,
      isOriginal: true,
    },
    duplicate: {
      noteId: 2,
      expression: 'b',
      sentencePreview: 'b',
      hasAudio: false,
      hasImage: false,
      isOriginal: false,
    },
  });

  assert.equal(result.cancelled, true);
  assert.equal(result.keepNoteId, 0);
  assert.equal(result.deleteNoteId, 0);
});

test('createFieldGroupingOverlayRuntime callback restores hidden visible overlay after resolver settles', async () => {
  let resolver: unknown = null;
  let visible = false;
  const visibilityTransitions: boolean[] = [];

  const runtime = createFieldGroupingOverlayRuntime<'runtime-options' | 'subsync'>({
    getMainWindow: () => null,
    getVisibleOverlayVisible: () => visible,
    setVisibleOverlayVisible: (nextVisible) => {
      visible = nextVisible;
      visibilityTransitions.push(nextVisible);
    },
    getResolver: () => resolver as ((choice: KikuFieldGroupingChoice) => void) | null,
    setResolver: (nextResolver: ((choice: KikuFieldGroupingChoice) => void) | null) => {
      resolver = nextResolver;
    },
    getRestoreVisibleOverlayOnModalClose: () => new Set<'runtime-options' | 'subsync'>(),
    sendToVisibleOverlay: () => true,
  });

  const callback = runtime.createFieldGroupingCallback();
  const pendingChoice = callback({
    original: {
      noteId: 1,
      expression: 'a',
      sentencePreview: 'a',
      hasAudio: false,
      hasImage: false,
      isOriginal: true,
    },
    duplicate: {
      noteId: 2,
      expression: 'b',
      sentencePreview: 'b',
      hasAudio: false,
      hasImage: false,
      isOriginal: false,
    },
  });

  assert.equal(visible, true);
  assert.ok(resolver);

  if (typeof resolver !== 'function') {
    throw new Error('expected field grouping resolver to be assigned');
  }

  (resolver as (choice: KikuFieldGroupingChoice) => void)({
    keepNoteId: 1,
    deleteNoteId: 2,
    deleteDuplicate: true,
    cancelled: false,
  });
  await pendingChoice;

  assert.equal(visible, false);
  assert.deepEqual(visibilityTransitions, [true, false]);
});

async function settleWithinMicrotasks<T>(
  promise: Promise<T>,
  attempts = 10,
): Promise<T | 'timeout'> {
  let settled = false;
  let settledValue: T | undefined;
  void promise.then((value) => {
    settled = true;
    settledValue = value;
  });

  for (let i = 0; i < attempts; i += 1) {
    await Promise.resolve();
    if (settled) {
      return settledValue as T;
    }
  }
  return 'timeout';
}

test('createFieldGroupingOverlayRuntime callback cancels and cleans up when kiku modal never acknowledges open', async () => {
  let resolver: ((choice: KikuFieldGroupingChoice) => void) | null = null;
  const sends: Array<{
    channel: string;
    payload: unknown;
    restoreOnModalClose?: string;
    preferModalWindow?: boolean;
  }> = [];
  const waitCalls: Array<{ modal: string; timeoutMs: number }> = [];
  const warnings: string[] = [];
  const closed: string[] = [];
  const originalSetTimeout = globalThis.setTimeout;
  globalThis.setTimeout = (() => 0) as unknown as typeof globalThis.setTimeout;

  try {
    const runtime = createFieldGroupingOverlayRuntime<'kiku'>({
      getMainWindow: () => null,
      getVisibleOverlayVisible: () => true,
      setVisibleOverlayVisible: () => {},
      getResolver: () => resolver,
      setResolver: (nextResolver) => {
        resolver = nextResolver;
      },
      getRestoreVisibleOverlayOnModalClose: () => new Set<'kiku'>(),
      sendToVisibleOverlay: (channel, payload, runtimeOptions) => {
        sends.push({
          channel,
          payload,
          restoreOnModalClose: runtimeOptions?.restoreOnModalClose,
          preferModalWindow: runtimeOptions?.preferModalWindow,
        });
        return true;
      },
      waitForModalOpen: async (modal, timeoutMs) => {
        waitCalls.push({ modal, timeoutMs });
        return false;
      },
      handleOverlayModalClosed: (modal) => {
        closed.push(modal);
      },
      logWarn: (message) => {
        warnings.push(message);
      },
    });

    const request = {
      original: {
        noteId: 1,
        expression: 'a',
        sentencePreview: 'a',
        hasAudio: false,
        hasImage: false,
        isOriginal: true,
      },
      duplicate: {
        noteId: 2,
        expression: 'b',
        sentencePreview: 'b',
        hasAudio: false,
        hasImage: false,
        isOriginal: false,
      },
    };
    const pendingChoice = runtime.createFieldGroupingCallback()(request);
    const result = await settleWithinMicrotasks(pendingChoice);

    assert.notEqual(result, 'timeout');
    assert.deepEqual(result, {
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: true,
      cancelled: true,
    });
    assert.equal(resolver, null);
    assert.deepEqual(
      sends.map(({ channel, restoreOnModalClose, preferModalWindow }) => ({
        channel,
        restoreOnModalClose,
        preferModalWindow,
      })),
      [
        {
          channel: 'kiku:field-grouping-request',
          restoreOnModalClose: 'kiku',
          preferModalWindow: true,
        },
        {
          channel: 'kiku:field-grouping-request',
          restoreOnModalClose: 'kiku',
          preferModalWindow: true,
        },
      ],
    );
    assert.deepEqual(waitCalls, [
      { modal: 'kiku', timeoutMs: 1500 },
      { modal: 'kiku', timeoutMs: 1500 },
    ]);
    assert.deepEqual(warnings, [
      'Kiku field grouping modal did not acknowledge modal open on first attempt; retrying dedicated modal window.',
    ]);
    assert.deepEqual(closed, ['kiku']);
  } finally {
    globalThis.setTimeout = originalSetTimeout;
  }
});
