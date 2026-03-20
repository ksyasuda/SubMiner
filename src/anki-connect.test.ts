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
