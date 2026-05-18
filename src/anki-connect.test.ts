import test from 'node:test';
import assert from 'node:assert/strict';
import { AnkiConnectClient } from './anki-connect';

test('AnkiConnectClient disables keep-alive agents to avoid stale socket retries', () => {
  const client = new AnkiConnectClient('http://127.0.0.1:8765') as unknown as {
    client: {
      defaults: {
        httpAgent?: { options?: { keepAlive?: boolean } };
        httpsAgent?: { options?: { keepAlive?: boolean } };
      };
    };
  };

  assert.equal(client.client.defaults.httpAgent?.options?.keepAlive, false);
  assert.equal(client.client.defaults.httpsAgent?.options?.keepAlive, false);
});

test('AnkiConnectClient includes action name in retry logs', async () => {
  const client = new AnkiConnectClient('http://127.0.0.1:8765') as unknown as {
    client: { post: (url: string, body: unknown, options: unknown) => Promise<unknown> };
    sleep: (ms: number) => Promise<void>;
  };
  let shouldFail = true;
  client.client = {
    post: async () => {
      if (shouldFail) {
        shouldFail = false;
        const error = Object.assign(new Error('socket hang up'), { code: 'ECONNRESET' });
        throw error;
      }
      return { data: { result: [], error: null } };
    },
  };
  client.sleep = async () => undefined;

  const originalInfo = console.info;
  const messages: string[] = [];
  try {
    console.info = (...args: unknown[]) => {
      messages.push(args.map((value) => String(value)).join(' '));
    };

    await (client as unknown as AnkiConnectClient).invoke('notesInfo', { notes: [1] });

    assert.match(messages.join('\n'), /AnkiConnect notesInfo retry 1\/3 after 200ms delay/);
  } finally {
    console.info = originalInfo;
  }
});

test('AnkiConnectClient lists decks and note type fields', async () => {
  const client = new AnkiConnectClient('http://127.0.0.1:8765') as unknown as {
    client: { post: (url: string, body: { action: string; params: unknown }) => Promise<unknown> };
  };
  const calls: Array<{ action: string; params: unknown }> = [];
  client.client = {
    post: async (_url, body) => {
      calls.push({ action: body.action, params: body.params });
      if (body.action === 'deckNames') {
        return { data: { result: ['Core', 'Mining'], error: null } };
      }
      if (body.action === 'modelNames') {
        return { data: { result: ['Japanese sentences'], error: null } };
      }
      if (body.action === 'modelFieldNames') {
        return { data: { result: ['Expression', 'Sentence'], error: null } };
      }
      return { data: { result: [], error: null } };
    },
  };

  const typedClient = client as unknown as AnkiConnectClient;

  assert.deepEqual(await typedClient.deckNames(), ['Core', 'Mining']);
  assert.deepEqual(await typedClient.modelNames(), ['Japanese sentences']);
  assert.deepEqual(await typedClient.modelFieldNames('Japanese sentences'), [
    'Expression',
    'Sentence',
  ]);
  assert.deepEqual(
    calls.map((call) => call.action),
    ['deckNames', 'modelNames', 'modelFieldNames'],
  );
});

test('AnkiConnectClient derives field names from sampled notes in a deck', async () => {
  const client = new AnkiConnectClient('http://127.0.0.1:8765') as unknown as {
    client: { post: (url: string, body: { action: string; params: unknown }) => Promise<unknown> };
  };
  const calls: Array<{ action: string; params: unknown }> = [];
  client.client = {
    post: async (_url, body) => {
      calls.push({ action: body.action, params: body.params });
      if (body.action === 'findNotes') {
        return { data: { result: [3, 1, 2], error: null } };
      }
      if (body.action === 'notesInfo') {
        return {
          data: {
            result: [
              { fields: { Sentence: { value: 'x' }, Expression: { value: 'y' } } },
              { fields: { Reading: { value: 'z' } } },
            ],
            error: null,
          },
        };
      }
      return { data: { result: [], error: null } };
    },
  };

  assert.deepEqual(
    await (client as unknown as AnkiConnectClient).fieldNamesForDeck('Mining "Current"'),
    ['Expression', 'Reading', 'Sentence'],
  );
  assert.deepEqual(calls[0], {
    action: 'findNotes',
    params: { query: 'deck:"Mining \\"Current\\""' },
  });
  assert.deepEqual(calls[1], {
    action: 'notesInfo',
    params: { notes: [3, 1, 2] },
  });
});

test('AnkiConnectClient derives model names from sampled notes in a deck', async () => {
  const client = new AnkiConnectClient('http://127.0.0.1:8765') as unknown as {
    client: { post: (url: string, body: { action: string; params: unknown }) => Promise<unknown> };
  };
  const calls: Array<{ action: string; params: unknown }> = [];
  client.client = {
    post: async (_url, body) => {
      calls.push({ action: body.action, params: body.params });
      if (body.action === 'findNotes') {
        return { data: { result: [5, 4], error: null } };
      }
      if (body.action === 'notesInfo') {
        return {
          data: {
            result: [{ modelName: 'Lapis Morph' }, { modelName: 'Kiku' }],
            error: null,
          },
        };
      }
      return { data: { result: [], error: null } };
    },
  };

  assert.deepEqual(await (client as unknown as AnkiConnectClient).modelNamesForDeck('Mining'), [
    'Kiku',
    'Lapis Morph',
  ]);
  assert.deepEqual(
    calls.map((call) => call.action),
    ['findNotes', 'notesInfo'],
  );
});
