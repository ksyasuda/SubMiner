import assert from 'node:assert/strict';
import test from 'node:test';
import * as fs from 'node:fs';
import * as http from 'node:http';
import * as os from 'node:os';
import * as path from 'node:path';
import type { AddressInfo } from 'node:net';

import { downloadToFile } from './utils.js';

interface TestServer {
  port: number;
  close: () => Promise<void>;
}

function startServer(handler: http.RequestListener): Promise<TestServer> {
  return new Promise((resolve) => {
    const server = http.createServer(handler);
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address() as AddressInfo;
      resolve({
        port,
        close: () =>
          new Promise((done) => {
            server.close(() => done());
          }),
      });
    });
  });
}

test('downloadToFile follows redirects that pass the allow-list', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-redirect-test-'));
  const server = await startServer((req, res) => {
    if (req.url === '/start') {
      res.writeHead(302, { Location: '/final' });
      res.end();
      return;
    }
    res.writeHead(200);
    res.end('subtitle body');
  });

  try {
    const destPath = path.join(dir, 'sub.ass');
    const result = await downloadToFile(
      `http://127.0.0.1:${server.port}/start`,
      destPath,
      {},
      { isAllowedRedirect: (url) => url.hostname === '127.0.0.1' },
    );

    assert.equal(result.ok, true);
    assert.equal(fs.readFileSync(destPath, 'utf8'), 'subtitle body');
  } finally {
    await server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('downloadToFile refuses redirects to a host outside the allow-list', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-redirect-test-'));
  let finalHits = 0;
  const server = await startServer((req, res) => {
    if (req.url === '/start') {
      res.writeHead(302, { Location: 'http://localhost.localdomain/evil' });
      res.end();
      return;
    }
    finalHits += 1;
    res.writeHead(200);
    res.end('should never be fetched');
  });

  try {
    const destPath = path.join(dir, 'sub.ass');
    const result = await downloadToFile(
      `http://127.0.0.1:${server.port}/start`,
      destPath,
      {},
      { isAllowedRedirect: (url) => url.hostname === '127.0.0.1' },
    );

    assert.equal(result.ok, false);
    if (!result.ok) {
      assert.match(result.error.error, /redirect/i);
    }
    assert.equal(finalHits, 0);
    assert.equal(fs.existsSync(destPath), false);
  } finally {
    await server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
