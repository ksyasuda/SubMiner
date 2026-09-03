import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { AnimeBridgeClient } from '../../anime-bridge/bridge-client';
import { ALL_SOURCES_ID } from '../../types/anime-browser';
import { createAnimeBrowserRuntime } from './anime-browser-runtime';

const client = {
  listAnimeSources: async () => [{ id: 'shared', name: 'Source', lang: 'en' }],
} as unknown as AnimeBridgeClient;

async function setupRuntime(defaultSource: string) {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-anime-default-source-'));
  await writeFile(path.join(dir, 'pkg.one.apk'), 'one');
  await writeFile(path.join(dir, 'pkg.two.apk'), 'two');
  let saved = defaultSource;
  const runtime = createAnimeBrowserRuntime({
    extensionsDir: () => dir,
    repos: () => [],
    setRepos: () => undefined,
    defaultSourceId: () => saved,
    setDefaultSourceId: (sourceId) => {
      saved = sourceId;
    },
    preferencesFile: path.join(dir, 'preferences.json'),
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
  await runtime.ensureBridge();
  return { runtime, saved: () => saved };
}

test('a new browser session opens on the configured default source', async () => {
  const { runtime } = await setupRuntime('pkg.two:shared');
  const snapshot = runtime.getSnapshot('fresh');
  assert.equal(snapshot.selectedSourceId, 'pkg.two:shared');
  assert.equal(snapshot.defaultSourceId, 'pkg.two:shared');
  await runtime.dispose();
});

test('"all" as the default selects every source, and an unknown default falls back to the first', async () => {
  const all = await setupRuntime(ALL_SOURCES_ID);
  assert.equal(all.runtime.getSnapshot('fresh').selectedSourceId, ALL_SOURCES_ID);
  await all.runtime.dispose();

  const missing = await setupRuntime('pkg.gone:shared');
  assert.equal(missing.runtime.getSnapshot('fresh').selectedSourceId, 'pkg.one:shared');
  await missing.runtime.dispose();
});

test('setDefaultSource persists an installed source and rejects an unknown one', async () => {
  const { runtime, saved } = await setupRuntime('');
  runtime.setDefaultSource('pkg.two:shared');
  assert.equal(saved(), 'pkg.two:shared');
  assert.throws(() => runtime.setDefaultSource('pkg.gone:shared'), /no longer installed/);
  assert.equal(saved(), 'pkg.two:shared');
  // Sessions that already exist keep their own selection; only new ones follow the default.
  assert.equal(runtime.getSnapshot('existing').selectedSourceId, 'pkg.two:shared');
  await runtime.dispose();
});
