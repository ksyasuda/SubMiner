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
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (next) => {
      visible = next;
    },
    setInvisibleOverlayVisible: () => {},
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
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    setInvisibleOverlayVisible: () => {},
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
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (nextVisible) => {
      visible = nextVisible;
      visibilityTransitions.push(nextVisible);
    },
    setInvisibleOverlayVisible: () => {},
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
