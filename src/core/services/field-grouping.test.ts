import test from 'node:test';
import assert from 'node:assert/strict';
import { KikuFieldGroupingChoice, KikuFieldGroupingRequestData } from '../../types';
import { createFieldGroupingCallback } from './field-grouping';

function makeRequestData(): KikuFieldGroupingRequestData {
  return {
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
      expression: 'a',
      sentencePreview: 'b',
      hasAudio: false,
      hasImage: false,
      isOriginal: false,
    },
  };
}

/**
 * Mirrors how main stores the resolver: it wraps the callback's resolver in a
 * sequence-guarded closure, so the value read back is never identity-equal to the
 * callback's own `finish`. The old `getResolver() === finish` clear-guard therefore
 * never matched and leaked the resolver, wedging every later grouping attempt.
 */
function createWrappedResolverStore() {
  let stored: ((choice: KikuFieldGroupingChoice) => void) | null = null;
  return {
    getResolver: () => stored,
    setResolver: (resolver: ((choice: KikuFieldGroupingChoice) => void) | null) => {
      stored = resolver ? (choice) => resolver(choice) : null;
    },
    respond: (choice: KikuFieldGroupingChoice) => {
      stored?.(choice);
    },
  };
}

test('field grouping callback clears the wrapped resolver after a renderer response', async () => {
  const store = createWrappedResolverStore();
  let visible = false;
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => visible,
    setVisibleOverlayVisible: (next) => {
      visible = next;
    },
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => true,
  });

  const pending = callback(makeRequestData());
  await Promise.resolve();
  assert.notEqual(store.getResolver(), null);

  const choice: KikuFieldGroupingChoice = {
    keepNoteId: 1,
    deleteNoteId: 2,
    deleteDuplicate: true,
    cancelled: false,
  };
  store.respond(choice);

  assert.deepEqual(await pending, choice);
  assert.equal(store.getResolver(), null);
});

test('field grouping callback does not reject the next request after a response', async () => {
  const store = createWrappedResolverStore();
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => true,
  });

  const first = callback(makeRequestData());
  await Promise.resolve();
  store.respond({ keepNoteId: 1, deleteNoteId: 2, deleteDuplicate: true, cancelled: false });
  const firstChoice = await first;
  assert.equal(firstChoice.cancelled, false);

  // The second attempt must reach the renderer, not short-circuit to an instant cancel.
  const second = callback(makeRequestData());
  await Promise.resolve();
  assert.notEqual(store.getResolver(), null);
  store.respond({ keepNoteId: 2, deleteNoteId: 1, deleteDuplicate: false, cancelled: false });
  const secondChoice = await second;
  assert.equal(secondChoice.cancelled, false);
  assert.equal(secondChoice.keepNoteId, 2);
});

test('field grouping callback dismisses the modal UI when the send fails', async () => {
  const store = createWrappedResolverStore();
  let dismissed = 0;
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => false,
    dismissModalUi: () => {
      dismissed += 1;
    },
  });

  const result = await callback(makeRequestData());
  assert.equal(result.cancelled, true);
  assert.equal(dismissed, 1);
  assert.equal(store.getResolver(), null);
});

test('field grouping callback handles modal dismiss failures on send failure', async () => {
  const store = createWrappedResolverStore();
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const callback = createFieldGroupingCallback({
      getVisibleOverlayVisible: () => false,
      setVisibleOverlayVisible: () => {},
      getResolver: store.getResolver,
      setResolver: store.setResolver,
      sendRequestToVisibleOverlay: () => false,
      dismissModalUi: () => {
        throw new Error('dismiss failed');
      },
    });

    const result = await callback(makeRequestData());

    assert.equal(result.cancelled, true);
    assert.equal(store.getResolver(), null);
  } finally {
    console.error = originalConsoleError;
  }
});

test('field grouping callback handles modal dismiss failures on timeout', async () => {
  const store = createWrappedResolverStore();
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const callback = createFieldGroupingCallback({
      getVisibleOverlayVisible: () => false,
      setVisibleOverlayVisible: () => {},
      getResolver: store.getResolver,
      setResolver: store.setResolver,
      sendRequestToVisibleOverlay: () => true,
      dismissModalUi: () => {
        throw new Error('dismiss failed');
      },
      responseTimeoutMs: 5,
    });

    const result = await callback(makeRequestData());

    assert.equal(result.cancelled, true);
    assert.equal(store.getResolver(), null);
  } finally {
    console.error = originalConsoleError;
  }
});

test('field grouping callback reports modal dismiss failures', async () => {
  const store = createWrappedResolverStore();
  const errors: unknown[] = [];
  const originalConsoleError = console.error;
  console.error = (...args: unknown[]) => {
    errors.push(args);
  };
  try {
    const callback = createFieldGroupingCallback({
      getVisibleOverlayVisible: () => false,
      setVisibleOverlayVisible: () => {},
      getResolver: store.getResolver,
      setResolver: store.setResolver,
      sendRequestToVisibleOverlay: () => false,
      dismissModalUi: () => {
        throw new Error('dismiss failed');
      },
    });

    await callback(makeRequestData());

    assert.equal(errors.length, 1);
  } finally {
    console.error = originalConsoleError;
  }
});

test('field grouping callback dismisses the modal UI when the response times out', async () => {
  const store = createWrappedResolverStore();
  let dismissed = 0;
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => true,
    dismissModalUi: () => {
      dismissed += 1;
    },
    responseTimeoutMs: 5,
  });

  const result = await callback(makeRequestData());
  assert.equal(result.cancelled, true);
  assert.equal(dismissed, 1);
  assert.equal(store.getResolver(), null);
});

test('field grouping callback does not dismiss the modal UI on a normal response', async () => {
  const store = createWrappedResolverStore();
  let dismissed = 0;
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => true,
    dismissModalUi: () => {
      dismissed += 1;
    },
    responseTimeoutMs: 10000,
  });

  const pending = callback(makeRequestData());
  await Promise.resolve();
  store.respond({ keepNoteId: 1, deleteNoteId: 2, deleteDuplicate: true, cancelled: false });
  await pending;
  assert.equal(dismissed, 0);
});

test('field grouping callback rejects a concurrent request while one is pending', async () => {
  const store = createWrappedResolverStore();
  let sends = 0;
  let dismissed = 0;
  const callback = createFieldGroupingCallback({
    getVisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    getResolver: store.getResolver,
    setResolver: store.setResolver,
    sendRequestToVisibleOverlay: () => {
      sends += 1;
      return true;
    },
    dismissModalUi: () => {
      dismissed += 1;
    },
    responseTimeoutMs: 10000,
  });

  const first = callback(makeRequestData());
  await Promise.resolve();
  assert.equal(sends, 1);

  const second = await callback(makeRequestData());
  assert.equal(second.cancelled, true);
  assert.equal(sends, 1);
  assert.equal(dismissed, 0);

  store.respond({ keepNoteId: 1, deleteNoteId: 2, deleteDuplicate: true, cancelled: false });
  await first;
});
