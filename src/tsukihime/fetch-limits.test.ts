import assert from 'node:assert/strict';
import test from 'node:test';
import * as http from 'node:http';
import type { AddressInfo } from 'node:net';

import { tsukihimeFetchJson } from './utils.js';

interface TestServer {
  baseUrl: string;
  close: () => Promise<void>;
}

function startServer(handler: http.RequestListener): Promise<TestServer> {
  return new Promise((resolve) => {
    const server = http.createServer(handler);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as AddressInfo;
      resolve({
        baseUrl: `http://127.0.0.1:${port}`,
        close: () =>
          new Promise((done) => {
            server.closeAllConnections?.();
            server.close(() => done());
          }),
      });
    });
  });
}

test('tsukihimeFetchJson gives up on a hanging response', async () => {
  const server = await startServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    // Never end the response.
    res.write('[');
  });

  try {
    const result = await tsukihimeFetchJson(
      '/json',
      {},
      { baseUrl: server.baseUrl, timeoutMs: 150 },
    );
    assert.equal(result.ok, false);
    if (!result.ok) {
      assert.match(result.error.error, /timed out/i);
    }
  } finally {
    await server.close();
  }
});

test('tsukihimeFetchJson rejects an oversized response', async () => {
  const server = await startServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(Array.from({ length: 2000 }, (_, index) => ({ id: index }))));
  });

  try {
    const result = await tsukihimeFetchJson(
      '/json',
      {},
      { baseUrl: server.baseUrl, maxResponseBytes: 512 },
    );
    assert.equal(result.ok, false);
    if (!result.ok) {
      assert.match(result.error.error, /too large/i);
    }
  } finally {
    await server.close();
  }
});

test('tsukihimeFetchJson decodes multi-byte characters split across chunks', async () => {
  const title = '葬送のフリーレン';
  const body = Buffer.from(JSON.stringify([{ id: 1, title }]), 'utf8');
  // Split inside the first Japanese character's 3-byte UTF-8 sequence.
  const splitAt = body.indexOf(Buffer.from(title, 'utf8')) + 1;

  const server = await startServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.write(body.subarray(0, splitAt));
    setTimeout(() => res.end(body.subarray(splitAt)), 10);
  });

  try {
    const result = await tsukihimeFetchJson<Array<{ id: number; title: string }>>(
      '/json',
      {},
      { baseUrl: server.baseUrl },
    );
    assert.equal(result.ok, true);
    if (result.ok) {
      assert.equal(result.data[0]!.title, title);
    }
  } finally {
    await server.close();
  }
});

test('tsukihimeFetchJson still parses a normal response', async () => {
  const server = await startServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify([{ id: 1, title: 'ok' }]));
  });

  try {
    const result = await tsukihimeFetchJson<Array<{ id: number }>>(
      '/json',
      { q: 'x' },
      { baseUrl: server.baseUrl },
    );
    assert.equal(result.ok, true);
    if (result.ok) {
      assert.deepEqual(result.data, [{ id: 1, title: 'ok' }]);
    }
  } finally {
    await server.close();
  }
});

test('tsukihimeFetchJson reports a structured error for a malformed base URL', async () => {
  const result = await tsukihimeFetchJson('/search/torrents', {}, { baseUrl: 'not a url' });
  assert.equal(result.ok, false);
  if (!result.ok) {
    assert.match(result.error.error, /base url/i);
  }
});

test('tsukihimeFetchJson rejects non-HTTP(S) base URLs', async () => {
  const result = await tsukihimeFetchJson('/search/torrents', {}, { baseUrl: 'file:///etc' });
  assert.equal(result.ok, false);
  if (!result.ok) {
    assert.match(result.error.error, /http/i);
  }
});
