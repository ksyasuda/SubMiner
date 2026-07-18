import assert from 'node:assert/strict';
import http from 'node:http';
import { once } from 'node:events';
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
  const recordedCards: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    recordCardsAdded: (count) => {
      recordedCards.push(count);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    { action: 'addNote' },
    Buffer.from(JSON.stringify({ result: 42, error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(processed, [42]);
  assert.deepEqual(recordedCards, [1]);
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest({ action: 'addNote' }, Buffer.from('42', 'utf8'));

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(processed, [42]);
});

test('proxy de-duplicates addNotes IDs within the same response', async () => {
  const processed: number[] = [];
  const recordedCards: number[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
      await new Promise((resolve) => setTimeout(resolve, 5));
    },
    recordCardsAdded: (count) => {
      recordedCards.push(count);
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    { action: 'addNotes' },
    Buffer.from(JSON.stringify({ result: [101, 102, 101, null], error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 2);
  assert.deepEqual(processed, [101, 102]);
  assert.deepEqual(recordedCards, [2]);
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [{ action: 'version' }, { action: 'addNote' }, { action: 'addNotes' }],
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    {
      action: 'multi',
      params: {
        actions: [{ action: 'version' }, { action: 'addNote' }, { action: 'addNotes' }],
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
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

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
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

test('proxy fallback-enqueues latest note for addNote responses without note IDs and escapes deck quotes', async () => {
  const processed: number[] = [];
  const recordedCards: number[] = [];
  const findNotesQueries: string[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    recordCardsAdded: (count) => {
      recordedCards.push(count);
    },
    getDeck: () => 'My "Japanese" Deck',
    findNotes: async (query) => {
      findNotesQueries.push(query);
      return [500, 501];
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    { action: 'addNote' },
    Buffer.from(JSON.stringify({ result: 0, error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(findNotesQueries, ['"deck:My \\"Japanese\\" Deck" added:1']);
  assert.deepEqual(processed, [501]);
  assert.deepEqual(recordedCards, [1]);
});

test('proxy tracks duplicate note ids from addNote request metadata before enrichment', async () => {
  const processed: number[] = [];
  const tracked: Array<{ noteId: number; duplicateNoteIds: number[] }> = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
    },
    trackAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
      tracked.push({ noteId, duplicateNoteIds });
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  (
    proxy as unknown as {
      maybeEnqueueFromRequest: (request: Record<string, unknown>, responseBody: Buffer) => void;
    }
  ).maybeEnqueueFromRequest(
    {
      action: 'addNote',
      params: {
        note: {},
        subminerDuplicateNoteIds: [11, -1, 40, 11, 25],
      },
    },
    Buffer.from(JSON.stringify({ result: 42, error: null }), 'utf8'),
  );

  await waitForCondition(() => processed.length === 1);
  assert.deepEqual(tracked, [{ noteId: 42, duplicateNoteIds: [11, 25, 40] }]);
  assert.deepEqual(processed, [42]);
});

test('proxy strips SubMiner duplicate metadata before forwarding upstream addNote request', async () => {
  let upstreamBody = '';
  const upstream = http.createServer(async (req, res) => {
    upstreamBody = await new Promise<string>((resolve) => {
      const chunks: Buffer[] = [];
      req.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    });
    res.statusCode = 200;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ result: 42, error: null }));
  });
  upstream.listen(0, '127.0.0.1');
  await once(upstream, 'listening');
  const upstreamAddress = upstream.address();
  assert.ok(upstreamAddress && typeof upstreamAddress === 'object');
  const upstreamPort = upstreamAddress.port;

  const tracked: Array<{ noteId: number; duplicateNoteIds: number[] }> = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async () => undefined,
    trackAddedDuplicateNoteIds: (noteId, duplicateNoteIds) => {
      tracked.push({ noteId, duplicateNoteIds });
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  try {
    proxy.start({
      host: '127.0.0.1',
      port: 0,
      upstreamUrl: `http://127.0.0.1:${upstreamPort}`,
    });

    const proxyServer = (
      proxy as unknown as {
        server: http.Server | null;
      }
    ).server;
    assert.ok(proxyServer);
    if (!proxyServer.listening) {
      await once(proxyServer, 'listening');
    }
    const proxyAddress = proxyServer.address();
    assert.ok(proxyAddress && typeof proxyAddress === 'object');
    const proxyPort = proxyAddress.port;

    const response = await fetch(`http://127.0.0.1:${proxyPort}`, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        action: 'addNote',
        version: 6,
        params: {
          note: {
            deckName: 'Mining',
            modelName: 'Sentence',
            fields: { Expression: '食べる' },
          },
          subminerDuplicateNoteIds: [18, 7],
        },
      }),
    });

    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), { result: 42, error: null });
    await waitForCondition(() => tracked.length === 1);
    assert.equal(upstreamBody.includes('subminerDuplicateNoteIds'), false);
    assert.deepEqual(tracked, [{ noteId: 42, duplicateNoteIds: [7, 18] }]);
  } finally {
    proxy.stop();
    upstream.close();
    await once(upstream, 'close');
  }
});

test('proxy returns addNote response without waiting for background enrichment', async () => {
  const processed: number[] = [];
  let releaseProcessing: (() => void) | undefined;
  const processingGate = new Promise<void>((resolve) => {
    releaseProcessing = resolve;
  });

  const upstream = http.createServer((req, res) => {
    assert.equal(req.method, 'POST');
    res.statusCode = 200;
    res.setHeader('content-type', 'application/json');
    res.end(JSON.stringify({ result: 42, error: null }));
  });
  upstream.listen(0, '127.0.0.1');
  await once(upstream, 'listening');
  const upstreamAddress = upstream.address();
  assert.ok(upstreamAddress && typeof upstreamAddress === 'object');
  const upstreamPort = upstreamAddress.port;

  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async (noteId) => {
      processed.push(noteId);
      await processingGate;
    },
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  try {
    proxy.start({
      host: '127.0.0.1',
      port: 0,
      upstreamUrl: `http://127.0.0.1:${upstreamPort}`,
    });

    const proxyServer = (
      proxy as unknown as {
        server: http.Server | null;
      }
    ).server;
    assert.ok(proxyServer);
    if (!proxyServer.listening) {
      await once(proxyServer, 'listening');
    }
    const proxyAddress = proxyServer.address();
    assert.ok(proxyAddress && typeof proxyAddress === 'object');
    const proxyPort = proxyAddress.port;

    const response = await Promise.race([
      fetch(`http://127.0.0.1:${proxyPort}`, {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
        },
        body: JSON.stringify({ action: 'addNote', version: 6, params: {} }),
      }),
      new Promise<never>((_, reject) => {
        setTimeout(() => reject(new Error('Timed out waiting for proxy response')), 500);
      }),
    ]);

    assert.equal(response.status, 200);
    assert.deepEqual(await response.json(), { result: 42, error: null });
    await waitForCondition(() => processed.length === 1);
    assert.deepEqual(processed, [42]);
  } finally {
    if (releaseProcessing) {
      releaseProcessing();
    }
    proxy.stop();
    upstream.close();
    await once(upstream, 'close');
  }
});

test('proxy detects self-referential loop configuration', () => {
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async () => undefined,
    logInfo: () => undefined,
    logWarn: () => undefined,
    logError: () => undefined,
  });

  const result = (
    proxy as unknown as {
      isSelfReferentialProxy: (options: {
        host: string;
        port: number;
        upstreamUrl: string;
      }) => boolean;
    }
  ).isSelfReferentialProxy({
    host: '127.0.0.1',
    port: 8766,
    upstreamUrl: 'http://localhost:8766',
  });

  assert.equal(result, true);
});

test('proxy continues without a local listener when its address is already bound', async () => {
  const occupiedServer = http.createServer();
  occupiedServer.listen(0, '127.0.0.1');
  await once(occupiedServer, 'listening');
  const occupiedAddress = occupiedServer.address();
  assert.ok(occupiedAddress && typeof occupiedAddress === 'object');
  const info: string[] = [];
  const warnings: string[] = [];
  const proxy = new AnkiConnectProxyServer({
    shouldAutoUpdateNewCards: () => true,
    processNewCard: async () => undefined,
    logInfo: (message) => info.push(message),
    logWarn: (message, ...args) => warnings.push([message, ...args].join(' ')),
    logError: () => undefined,
  });

  try {
    proxy.start({
      host: '127.0.0.1',
      port: occupiedAddress.port,
      upstreamUrl: 'http://127.0.0.1:8765',
    });

    await proxy.waitUntilReady();

    assert.equal(proxy.isRunning, false);
    assert.deepEqual(warnings, [
      `[anki-proxy] Local proxy unavailable because http://127.0.0.1:${occupiedAddress.port} is already in use; continuing without it. Change ankiConnect.proxy.port or stop the process using that address.`,
    ]);
    proxy.stop();
    assert.deepEqual(info, []);
  } finally {
    proxy.stop();
    occupiedServer.close();
    await once(occupiedServer, 'close');
  }
});
