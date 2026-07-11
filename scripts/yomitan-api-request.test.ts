import assert from 'node:assert/strict';
import test from 'node:test';

import { fetchYomitanApi } from './yomitan-api-request.js';

function createStalledServer(): ReturnType<typeof Bun.serve> {
  return Bun.serve({
    hostname: '127.0.0.1',
    port: 0,
    fetch: () => new Promise<Response>(() => {}),
  });
}

test('fetchYomitanApi reports a clear timeout when the bridge stalls', async () => {
  const server = createStalledServer();

  try {
    await assert.rejects(
      () => fetchYomitanApi(`${server.url}tokenize`, { method: 'POST' }, 20),
      /Yomitan API request timed out after 20ms/,
    );
  } finally {
    server.stop(true);
  }
});

test('fetchYomitanApi preserves an init cancellation signal', async () => {
  const server = createStalledServer();
  const controller = new AbortController();
  const reason = new Error('cancelled by caller');

  try {
    const request = fetchYomitanApi(`${server.url}tokenize`, { signal: controller.signal }, 50);
    controller.abort(reason);
    await assert.rejects(request, (error) => error === reason);
  } finally {
    server.stop(true);
  }
});

test('fetchYomitanApi preserves a Request input cancellation signal', async () => {
  const server = createStalledServer();
  const controller = new AbortController();
  const reason = new Error('cancelled by request');
  const input = new Request(`${server.url}tokenize`, { signal: controller.signal });

  try {
    const request = fetchYomitanApi(input, undefined, 50);
    controller.abort(reason);
    await assert.rejects(request, (error) => error === reason);
  } finally {
    server.stop(true);
  }
});

test('fetchYomitanApi accepts a null init cancellation signal', async () => {
  const server = Bun.serve({
    hostname: '127.0.0.1',
    port: 0,
    fetch: () => new Response('ok'),
  });

  try {
    const response = await fetchYomitanApi(`${server.url}tokenize`, { signal: null }, 50);
    assert.equal(response.status, 200);
  } finally {
    server.stop(true);
  }
});
