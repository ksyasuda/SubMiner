import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { BridgePreference } from '../../anime-bridge/types';
import type { BridgeAnimePage } from '../../anime-bridge/types';
import type { AnimeBrowserSearchUpdate } from '../../types/anime-browser';
import { createAnimeBrowserRuntime } from './anime-browser-runtime';
import type { AnimeBridgeClient } from '../../anime-bridge/bridge-client';

function textPreference(key: string, value = ''): BridgePreference {
  return { key, editTextPreference: { title: key, value, text: value } };
}

async function setupRuntime(
  client: Record<string, unknown>,
  packages: Record<string, string>,
  storedPreferences?: Record<string, BridgePreference[]>,
  onSearchUpdate?: (update: AnimeBrowserSearchUpdate, sessionId: string) => void,
) {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-anime-runtime-'));
  for (const [pkg, contents] of Object.entries(packages)) {
    await writeFile(path.join(dir, `${pkg}.apk`), contents);
  }
  const preferencesFile = path.join(dir, 'preferences.json');
  if (storedPreferences) await writeFile(preferencesFile, JSON.stringify(storedPreferences));
  const runtime = createAnimeBrowserRuntime({
    extensionsDir: () => dir,
    repos: () => [],
    setRepos: () => undefined,
    preferencesFile,
    ensureBinaries: async () => ({}) as never,
    checkBridgeUpdate: async () => null,
    stageBridgeUpdate: async () => {
      throw new Error('not under test');
    },
    sendMpvCommand: () => undefined,
    ensureMpvConnected: async () => true,
    onBridgeState: () => undefined,
    onSearchUpdate,
    log: () => undefined,
    startSidecar: async () => ({
      client: client as unknown as AnimeBridgeClient,
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
  return { runtime, preferencesFile };
}

test('setPreference overlays saved values onto the extension current schema', async () => {
  const saved = textPreference('address', 'saved-address');
  const current = [textPreference('address'), textPreference('new-setting')];
  let received: BridgePreference[] = [];
  let schemaCalls = 0;
  const pkgBytes = 'one';
  const client = {
    listAnimeSources: async () => [{ id: 'shared', name: 'One', lang: 'en' }],
    getSourcePreferences: async () => {
      schemaCalls += 1;
      return current;
    },
    setSourcePreference: async (source: { preferences?: BridgePreference[] }) => {
      received = source.preferences ?? [];
      return received;
    },
  };
  const { runtime } = await setupRuntime(
    client,
    { 'pkg.one': pkgBytes },
    { 'pkg.one:shared': [saved] },
  );
  assert.equal(runtime.getSnapshot().sources[0]?.id, 'pkg.one:shared');

  await runtime.setPreference('pkg.one:shared', 'new-setting', 'chosen');

  assert.equal(schemaCalls, 1);
  const address = received.find((entry) => entry.key === 'address')?.editTextPreference as
    | Record<string, unknown>
    | undefined;
  const added = received.find((entry) => entry.key === 'new-setting')?.editTextPreference as
    | Record<string, unknown>
    | undefined;
  assert.equal(address?.value, 'saved-address');
  assert.equal(added?.value, 'chosen');
  await runtime.dispose();
});

test('colliding bridge ids keep package preferences isolated and uninstall clears only its package', async () => {
  const seen: Array<{ sourceId?: string; preferences?: BridgePreference[]; fingerprint: string }> =
    [];
  const client = {
    listAnimeSources: async () => [{ id: 'shared', name: 'Source', lang: 'en' }],
    getSourcePreferences: async (source: {
      sourceId?: string;
      preferences?: BridgePreference[];
      fingerprint: string;
    }) => {
      seen.push(source);
      return [textPreference('password')];
    },
  };
  const one = textPreference('password', 'one-secret');
  const two = textPreference('password', 'two-secret');
  const { runtime, preferencesFile } = await setupRuntime(
    client,
    { 'pkg.one': 'one', 'pkg.two': 'two' },
    { 'pkg.one:shared': [one], 'pkg.two:shared': [two] },
  );

  const sourceIds = runtime.getSnapshot().sources.map((source) => source.id);
  assert.deepEqual(sourceIds, ['pkg.one:shared', 'pkg.two:shared']);
  const oneView = await runtime.getPreferences('pkg.one:shared');
  const twoView = await runtime.getPreferences('pkg.two:shared');

  assert.equal(oneView[0]?.value, 'one-secret');
  assert.equal(twoView[0]?.value, 'two-secret');
  assert.deepEqual(
    seen.map((source) => source.sourceId),
    ['shared', 'shared'],
  );
  assert.notEqual(seen[0]?.fingerprint, seen[1]?.fingerprint);
  assert.equal(
    (seen[0]?.preferences?.[0]?.editTextPreference as Record<string, unknown>)?.value,
    'one-secret',
  );
  assert.equal(
    (seen[1]?.preferences?.[0]?.editTextPreference as Record<string, unknown>)?.value,
    'two-secret',
  );

  await runtime.removeExtension('pkg.one');
  const persisted = JSON.parse(await readFile(preferencesFile, 'utf8')) as Record<string, unknown>;
  assert.equal(persisted['pkg.one:shared'], undefined);
  assert.deepEqual(persisted['pkg.two:shared'], [two]);
  await runtime.dispose();
});

test('standalone and modal browser sessions keep source selection and searches isolated', async () => {
  const client = {
    listAnimeSources: async () => [{ id: 'shared', name: 'Source', lang: 'en' }],
    searchAnime: async (source: { fingerprint: string }, query: string) => ({
      animes: [{ url: `${source.fingerprint}/${query}`, title: source.fingerprint }],
      hasNextPage: false,
    }),
  };
  const { runtime } = await setupRuntime(client, { 'pkg.one': 'one', 'pkg.two': 'two' });
  const sourceIds = runtime.getSnapshot('standalone').sources.map((source) => source.id);
  runtime.selectSource(sourceIds[0]!, 'standalone');
  runtime.selectSource(sourceIds[1]!, 'modal');

  const [standalone, modal] = await Promise.all([
    runtime.search('one', 1, 'standalone'),
    runtime.search('two', 1, 'modal'),
  ]);

  assert.equal(runtime.getSnapshot('standalone').selectedSourceId, sourceIds[0]);
  assert.equal(runtime.getSnapshot('modal').selectedSourceId, sourceIds[1]);
  assert.equal(standalone.entries[0]?.sourceId, sourceIds[0]);
  assert.equal(modal.entries[0]?.sourceId, sourceIds[1]);
  await runtime.dispose();
});

test('releasing a browser session suppresses its pending search updates', async () => {
  let resolveOldSearch: (page: BridgeAnimePage) => void = () => undefined;
  let notifyOldSearchStarted: () => void = () => undefined;
  const oldSearchStarted = new Promise<void>((resolve) => {
    notifyOldSearchStarted = resolve;
  });
  const updates: Array<{ sessionId: string; update: AnimeBrowserSearchUpdate }> = [];
  const client = {
    listAnimeSources: async () => [{ id: 'shared', name: 'Source', lang: 'en' }],
    searchAnime: async (_source: unknown, query: string): Promise<BridgeAnimePage> => {
      if (query !== 'old') {
        return { animes: [{ url: `/${query}`, title: query }], hasNextPage: false };
      }
      notifyOldSearchStarted();
      return await new Promise<BridgeAnimePage>((resolve) => {
        resolveOldSearch = resolve;
      });
    },
  };
  const { runtime } = await setupRuntime(
    client,
    { 'pkg.one': 'one' },
    undefined,
    (update, sessionId) => {
      updates.push({ sessionId, update });
    },
  );

  const oldSearch = runtime.search('old', 1, 'reused');
  await oldSearchStarted;
  runtime.releaseSession('reused');
  await runtime.search('new', 1, 'reused');
  resolveOldSearch({ animes: [{ url: '/old', title: 'old' }], hasNextPage: false });
  await oldSearch;

  const resultUrls = updates.flatMap(({ sessionId, update }) =>
    sessionId === 'reused' && update.kind === 'result'
      ? update.entries.map((entry) => entry.url)
      : [],
  );
  assert.deepEqual(resultUrls, ['/new']);
  await runtime.dispose();
});
