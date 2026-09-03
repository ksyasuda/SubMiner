import assert from 'node:assert/strict';
import test from 'node:test';
import { existsSync } from 'node:fs';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { AnimeBridgeClient } from '../../anime-bridge/bridge-client';
import { createAnimeBrowserRuntime } from './anime-browser-runtime';

const REPO_URL = 'https://repo.example/anime/index.json';
const PKG = 'eu.kanade.tachiyomi.animeextension.all.example';

function createTestRuntime(
  directory: string,
  client: AnimeBridgeClient,
  repos: readonly string[] = [],
) {
  return createAnimeBrowserRuntime({
    extensionsDir: () => directory,
    repos: () => [...repos],
    setRepos: () => undefined,
    preferencesFile: path.join(directory, 'preferences.json'),
    ensureBinaries: async () => ({}) as never,
    checkBridgeUpdate: async () => null,
    stageBridgeUpdate: async () => {
      throw new Error('not under test');
    },
    sendMpvCommand: () => undefined,
    ensureMpvConnected: async () => true,
    onBridgeState: () => undefined,
    log: () => undefined,
    startSidecar: async () => ({
      client,
      baseUrl: 'http://127.0.0.1:12345',
      port: 12345,
      stop: async () => undefined,
      onExit: () => undefined,
    }),
    startStreamStripProxy: async () => ({
      origin: 'http://127.0.0.1:12346',
      port: 12346,
      close: async () => undefined,
    }),
  });
}

test('extension mutations run in request order while a download is pending', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subminer-extension-mutations-'));
  const apkPath = path.join(directory, `${PKG}.apk`);
  await writeFile(apkPath, 'old apk');

  let notifyDownloadStarted: () => void = () => undefined;
  const downloadStarted = new Promise<void>((resolve) => {
    notifyDownloadStarted = resolve;
  });
  let releaseDownload: () => void = () => undefined;
  const downloadGate = new Promise<void>((resolve) => {
    releaseDownload = resolve;
  });
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    if (String(input) === REPO_URL) {
      return new Response(
        JSON.stringify([
          {
            name: 'Aniyomi: Example',
            pkg: PKG,
            apk: 'example.apk',
            lang: 'all',
            code: 2,
            version: '2.0',
          },
        ]),
      );
    }
    notifyDownloadStarted();
    await downloadGate;
    const body = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x01]) as unknown as BodyInit;
    return new Response(body);
  }) as typeof fetch;

  const client = {
    listAnimeSources: async () => [{ id: 'one', name: 'Example', lang: 'all' }],
  } as unknown as AnimeBridgeClient;
  const runtime = createTestRuntime(directory, client, [REPO_URL]);

  try {
    await runtime.ensureBridge();
    const install = runtime.installExtension(PKG);
    await downloadStarted;
    const remove = runtime.removeExtension(PKG);
    releaseDownload();
    await Promise.all([install, remove]);

    assert.equal(existsSync(apkPath), false);
  } finally {
    globalThis.fetch = originalFetch;
    await runtime.dispose();
  }
});

test('startup scanning cannot restore an extension removed concurrently', async () => {
  const directory = await mkdtemp(path.join(tmpdir(), 'subminer-extension-startup-'));
  const apkPath = path.join(directory, `${PKG}.apk`);
  await writeFile(apkPath, 'old apk');

  let notifyScanStarted: () => void = () => undefined;
  const scanStarted = new Promise<void>((resolve) => {
    notifyScanStarted = resolve;
  });
  let releaseScan: () => void = () => undefined;
  const scanGate = new Promise<void>((resolve) => {
    releaseScan = resolve;
  });
  const client = {
    listAnimeSources: async () => {
      notifyScanStarted();
      await scanGate;
      return [{ id: 'one', name: 'Example', lang: 'all' }];
    },
  } as unknown as AnimeBridgeClient;
  const runtime = createTestRuntime(directory, client);

  try {
    const start = runtime.ensureBridge();
    await scanStarted;
    const remove = runtime.removeExtension(PKG);
    releaseScan();
    await Promise.all([start, remove]);

    assert.equal(existsSync(apkPath), false);
    assert.deepEqual(runtime.getSnapshot().installed, []);
    assert.deepEqual(runtime.getSnapshot().sources, []);
  } finally {
    await runtime.dispose();
  }
});
