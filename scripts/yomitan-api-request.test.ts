import assert from 'node:assert/strict';
import test from 'node:test';

import { fetchYomitanApi } from './yomitan-api-request.js';

test('fetchYomitanApi reports a clear timeout when the bridge stalls', async () => {
  const server = Bun.serve({
    hostname: '127.0.0.1',
    port: 0,
    fetch: () => new Promise<Response>(() => {}),
  });

  try {
    await assert.rejects(
      () => fetchYomitanApi(`${server.url}tokenize`, { method: 'POST' }, 20),
      /Yomitan API request timed out after 20ms/,
    );
  } finally {
    server.stop(true);
  }
});
