import assert from 'node:assert/strict';
import test from 'node:test';
import { AnkiConnectProxyServer } from './anki-connect-proxy';

async function waitForCondition(
  condition: () => boolean,
  timeoutMs = 2000,
  intervalMs = 10,
): Promise<void> {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (condition()) return;
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
  throw new Error('Timed out waiting for condition');
}

test('proxy enqueues addNote result for enrichment', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    { action: 'addNote' },
    Buffer.from(JSON.stringify({ result: 42, error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(processed, [42]);
});

test('proxy enqueues addNote bare numeric response for enrichment', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest({ action: 'addNote' }, Buffer.from('42', 'utf8'));

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(processed, [42]);
});

test('proxy de-duplicates addNotes IDs within the same response', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
      await new Promise((resolve) => setTimeout(resolve, 5));
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    { action: 'addNotes' },
    Buffer.from(JSON.stringify({ result: [101, 102, 101, null], error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 2);
  assert.deepEqual(processed, [101, 102]);
});

test('proxy enqueues note IDs from multi action addNote/addNotes results', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [
          { action: 'version' },
          { action: 'addNote' },
          { action: 'addNotes' },
        ],
      },
    },
    Buffer.from(JSON.stringify({ result: [6, 777, [888, 777, null]], error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 2);
  assert.deepEqual(processed, [777, 888]);
});

test('proxy enqueues note IDs from bare multi action results', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [{ action: 'version' }, { action: 'addNote' }, { action: 'addNotes' }],
      },
    },
    Buffer.from(JSON.stringify([6, 777, [888, null]]), 'utf8'),
  );

  await waitForCondition(() => processed.length === 2);
  assert.deepEqual(processed, [777, 888]);
});

test('proxy enqueues note IDs from multi action envelope results', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [
          { action: 'version' },
          { action: 'addNote' },
          { action: 'addNotes' },
        ],
      },
    },
    Buffer.from(
      JSON.stringify({
        result: [
          { result: 6, error: null },
          { result: 777, error: null },
          { result: [888, 777, null], error: null },
        ],
        error: null,
      }),
      'utf8',
    ),
  );

  await waitForCondition(() => processed.length === 2);
  assert.deepEqual(processed, [777, 888]);
});

test('proxy skips auto-enrichment when auto-update is disabled', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => false,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    { action: 'addNote' },
    Buffer.from(JSON.stringify({ result: 303, error: null }), 'utf8'),
  );

  await new Promise((resolve) => setTimeout(resolve, 30));
  assert.deepEqual(processed, []);
});

test('proxy ignores addNote when upstream response reports error', async () => {
  const processed: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    { action: 'addNote' },
    Buffer.from(JSON.stringify({ result: 123, error: 'duplicate' }), 'utf8'),
  );

  await new Promise((resolve) => setTimeout(resolve, 30));
  assert.deepEqual(processed, []);
});

test('proxy does not fallback-enqueue latest note for multi requests without add actions', async () => {
  const processed: number[] = [];
  const findNotesQueries: string[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    getDeck: () => 'Mining',
    findNotes: async (query) => {
      findNotesQueries.push(query);
      return [999];
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (proxy as unknown as {
    maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
  }).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [{ action: 'version' }, { action: 'deckNames' }],
      },
    },
    Buffer.from(JSON.stringify({ result: [6, ['Default']], error: null }), 'utf8'),
  );

  await new Promise((resolve) => setTimeout(resolve, 30));
  assert.deepEqual(findNotesQueries, []);
  assert.deepEqual(processed, []);
});

test('proxy detects self-referential loop configuration', () => {
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async () => undefined,
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  const result = (proxy as unknown as {
    isSelfReferentialProxy: (options: {
      host: string;
      port: number;
      upstreamUrl: string;
    }) => boolean;
  }).isSelfReferentialProxy({
    host: '127.0.0.1',
    port: 8766,
    upstreamUrl: 'http://localhost:8766',
  });

  assert.equal(result, true);
});
