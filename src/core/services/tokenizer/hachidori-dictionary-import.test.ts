import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { once } from 'node:events';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  resolveHachidoriManagementUrl,
  uploadHachidoriDictionary,
} from './hachidori-dictionary-import';
import { importYomitanDictionaryFromZip } from './yomitan-parser-runtime';
import { createDeps } from './yomitan-scan-test-harness';

test('linked Hachidori uploads replacement bytes, retries a busy host, and never invokes local import', async () => {
  const zipPath = path.join(await mkdtemp(path.join(os.tmpdir(), 'hachi-import-')), 'merged.zip');
  const archive = Buffer.from('PK-test-archive');
  await writeFile(zipPath, archive);
  let attempts = 0;
  const server = createServer(async (request, response) => {
    attempts += 1;
    assert.equal(request.url, '/import?name=merged.zip&replace=true');
    assert.equal(request.method, 'POST');
    const chunks: Buffer[] = [];
    for await (const chunk of request) chunks.push(Buffer.from(chunk));
    assert.deepEqual(Buffer.concat(chunks), archive);
    response.writeHead(attempts === 1 ? 409 : 200, { 'Content-Type': 'application/json' });
    response.end(
      JSON.stringify(attempts === 1 ? { error: 'busy' } : { ok: true, report: { success: true } }),
    );
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  assert.ok(address && typeof address === 'object');
  const managementUrl = `http://127.0.0.1:${address.port}`;
  let connected = true;
  const deps = {
    ...createDeps(async (script) => {
      assert.ok(script.includes('hd_sharing_status'), 'must not invoke local ZIP automation');
      return {
        ok: true,
        sharing: {
          client: {
            linked: true,
            connected,
            address: 'ws://127.0.0.1:8771/link',
            host: { name: 'Hachidori Docker host', dictionaryCount: 8 },
          },
        },
      };
    }),
    getYomitanExt: () => ({
      id: 'hachi',
      name: 'Hachidori',
      version: '1',
      path: '',
      url: '',
      manifest: {},
    }),
  };
  try {
    assert.equal(
      await importYomitanDictionaryFromZip(zipPath, deps, { error: assert.fail }, managementUrl),
      true,
    );
    assert.equal(attempts, 2);
    const logger = { error: () => {} };
    connected = false;
    assert.equal(await importYomitanDictionaryFromZip(zipPath, deps, logger, managementUrl), false);
    assert.equal(attempts, 2);
  } finally {
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
});

test('a linked host that accepts imports gets the ZIP over the link, not its management API', async () => {
  const zipPath = path.join(await mkdtemp(path.join(os.tmpdir(), 'hachi-import-')), 'merged.zip');
  await writeFile(zipPath, 'PK-linked');
  const linkedUploads: string[] = [];
  // Serves both the parser window (sharing status) and the settings window (import).
  const runScript = async (script: string): Promise<unknown> => {
    if (script.includes('hd_sharing_status')) {
      return {
        ok: true,
        sharing: {
          client: {
            linked: true,
            connected: true,
            address: 'ws://desktop:8771/link',
            host: { name: 'Helium', dictionaryCount: 8, capabilities: ['linked-import-v1'] },
          },
        },
      };
    }
    const archiveUrl = /importDictionaryArchiveUrl\(\s*("[^"]+")/.exec(script)?.[1];
    if (archiveUrl) linkedUploads.push(await (await fetch(JSON.parse(archiveUrl))).text());
    return true;
  };
  const settingsWindow = {
    isDestroyed: () => false,
    destroy: () => undefined,
    webContents: { executeJavaScript: runScript },
  };
  const deps = {
    ...createDeps(runScript, {
      createYomitanExtensionWindow: async () => settingsWindow,
    }),
    getYomitanExt: () => ({
      id: 'hachi',
      name: 'Hachidori',
      version: '1',
      path: '',
      url: '',
      manifest: {},
    }),
  };
  // Nothing listens here: reaching the management API would fail the import.
  assert.equal(
    await importYomitanDictionaryFromZip(
      zipPath,
      deps,
      { error: assert.fail },
      'http://127.0.0.1:9',
    ),
    true,
  );
  assert.deepEqual(linkedUploads, ['PK-linked']);
});

test('Hachidori upload requires a successful import report, not just HTTP success', async () => {
  const zipPath = path.join(await mkdtemp(path.join(os.tmpdir(), 'hachi-import-')), 'merged.zip');
  await writeFile(zipPath, 'bad archive');
  const server = createServer((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(
      JSON.stringify({ ok: true, report: { success: false, error: 'Invalid archive' } }),
    );
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  assert.ok(address && typeof address === 'object');
  try {
    await assert.rejects(
      uploadHachidoriDictionary(zipPath, `http://127.0.0.1:${address.port}`),
      /Invalid archive/,
    );
  } finally {
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
});

test('management URL comes from the linked Docker host unless overridden', () => {
  const docker = { address: 'ws://pve-main:8771/link', name: 'Hachidori Docker host' };
  const warnings: string[] = [];
  const warn = (message: string) => warnings.push(message);
  assert.equal(
    resolveHachidoriManagementUrl(docker, '', 'merged.zip', warn),
    'http://pve-main:8780',
  );
  assert.equal(
    resolveHachidoriManagementUrl(docker, 'http://pve-main:9000', 'merged.zip', warn),
    'http://pve-main:9000',
  );
  assert.deepEqual(warnings, []);
  assert.equal(
    resolveHachidoriManagementUrl(
      { address: 'ws://127.0.0.1:8771/link', name: 'Hachidori Docker host' },
      'http://localhost:8780',
      'merged.zip',
      warn,
    ),
    'http://localhost:8780',
  );
  assert.deepEqual(warnings, []);
  assert.equal(
    resolveHachidoriManagementUrl(docker, 'http://127.0.0.1:8780', 'merged.zip', warn),
    'http://127.0.0.1:8780',
  );
  assert.match(warnings[0] ?? '', /linked to pve-main/);
});

test('hosts without linked imports or an override explain the update or manual import', () => {
  assert.throws(
    () =>
      resolveHachidoriManagementUrl(
        { address: 'ws://desktop:8771/link', name: 'Chrome' },
        '',
        '/dicts/merged.zip',
      ),
    /Chrome at desktop cannot receive dictionary uploads\. Update its Hachidori to 0\.2\.3 or later, or import \/dicts\/merged\.zip/,
  );
});

test('unreachable management API names the origin instead of a bare fetch failure', async () => {
  const zipPath = path.join(await mkdtemp(path.join(os.tmpdir(), 'hachi-import-')), 'merged.zip');
  await writeFile(zipPath, 'archive');
  const server = createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  assert.ok(address && typeof address === 'object');
  await new Promise<void>((resolve) => server.close(() => resolve()));
  const origin = `http://127.0.0.1:${address.port}`;
  await assert.rejects(
    uploadHachidoriDictionary(zipPath, origin),
    new RegExp(`Could not reach the Hachidori management API at ${origin}: `),
  );
});
