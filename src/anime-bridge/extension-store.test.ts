import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { tmpdir } from 'node:os';
import path from 'node:path';
import {
  listExtensionSources,
  readInstalledExtensions,
  toBridgeSource,
  toInstalledExtensionViews,
  type ExtensionSource,
  type InstalledExtension,
} from './extension-store';
import type { AnimeBridgeClient } from './bridge-client';

async function makeExtensionDir(files: Record<string, string>): Promise<string> {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-ext-'));
  for (const [name, contents] of Object.entries(files)) {
    await writeFile(path.join(dir, name), contents);
  }
  return dir;
}

function fakeClient(
  impl: (source: { fingerprint: string }) => Promise<unknown[]>,
): AnimeBridgeClient {
  return { listAnimeSources: impl } as unknown as AnimeBridgeClient;
}

test('readInstalledExtensions fingerprints apks without holding their bytes', async () => {
  const dir = await makeExtensionDir({ 'my-source.apk': 'APK-BYTES' });
  const extensions = await readInstalledExtensions(dir);

  assert.equal(extensions.length, 1);
  assert.equal(extensions[0]?.fallbackName, 'my-source');
  assert.equal(extensions[0]?.sha256, createHash('sha256').update('APK-BYTES').digest('hex'));
});

test('the fingerprint changes when an apk is replaced in place', async () => {
  const dir = await makeExtensionDir({ 'my-source.apk': 'V1' });
  const before = (await readInstalledExtensions(dir))[0]?.sha256;
  await writeFile(path.join(dir, 'my-source.apk'), 'V2');
  const after = (await readInstalledExtensions(dir))[0]?.sha256;

  assert.notEqual(before, after);
});

test('toBridgeSource reads the apk only when the bridge asks for it', async () => {
  const dir = await makeExtensionDir({ 'lazy.apk': 'APK-BYTES' });
  const extension = (await readInstalledExtensions(dir))[0]!;

  const bridgeSource = toBridgeSource(extension);
  assert.equal(bridgeSource.fingerprint, extension.sha256);
  assert.equal(Buffer.from(await bridgeSource.loadApkBase64(), 'base64').toString(), 'APK-BYTES');
});

test('readInstalledExtensions ignores non-apk files and subdirectories', async () => {
  const dir = await makeExtensionDir({ 'a.apk': 'A', 'notes.txt': 'x', 'b.APK': 'B' });
  await mkdir(path.join(dir, 'nested.apk'), { recursive: true });

  const names = (await readInstalledExtensions(dir)).map((e) => e.fallbackName);
  // Sorted, case-insensitive extension match, directories excluded.
  assert.deepEqual(names, ['a', 'b']);
});

test('readInstalledExtensions returns empty for a missing directory', async () => {
  assert.deepEqual(await readInstalledExtensions('/nonexistent/subminer/extensions'), []);
});

test('toBridgeSource includes sourceId only when selecting inside a factory apk', () => {
  const extension: InstalledExtension = { file: '/x/a.apk', fallbackName: 'a', sha256: 'hash-a' };
  assert.equal(toBridgeSource(extension).sourceId, undefined);
  assert.equal(toBridgeSource(extension, 'src-1').sourceId, 'src-1');
  assert.equal(toBridgeSource(extension, 'src-1').fingerprint, 'hash-a');
});

test('listExtensionSources flattens every source a factory apk provides', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/multi.apk', fallbackName: 'multi', sha256: 'hash-a' },
  ];
  const client = fakeClient(async () => [
    { id: 101, name: 'Source One', lang: 'en' },
    { id: '102', name: 'Source Two', lang: 'ja' },
  ]);

  const sources = await listExtensionSources(client, extensions);

  assert.equal(sources.length, 2);
  // Numeric bridge ids are normalized and package-qualified for UI state.
  assert.equal(sources[0]?.id, 'multi:101');
  assert.equal(sources[0]?.bridgeId, '101');
  assert.equal(sources[0]?.name, 'Source One');
  assert.equal(sources[1]?.lang, 'ja');
});

test('sources with the same bridge id in different packages have distinct runtime ids', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/one.apk', fallbackName: 'pkg.one', sha256: 'hash-one' },
    { file: '/x/two.apk', fallbackName: 'pkg.two', sha256: 'hash-two' },
  ];
  const client = fakeClient(async () => [{ id: 'shared', name: 'Source', lang: 'en' }]);

  const sources = await listExtensionSources(client, extensions);

  assert.deepEqual(
    sources.map(({ id, bridgeId, pkg }) => ({ id, bridgeId, pkg })),
    [
      { id: 'pkg.one:shared', bridgeId: 'shared', pkg: 'pkg.one' },
      { id: 'pkg.two:shared', bridgeId: 'shared', pkg: 'pkg.two' },
    ],
  );
});

test('listExtensionSources falls back to the file name and a default language', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/my-ext.apk', fallbackName: 'my-ext', sha256: 'hash-a' },
  ];
  const client = fakeClient(async () => [{ id: '1', name: '   ' }]);

  const sources = await listExtensionSources(client, extensions);
  assert.equal(sources[0]?.name, 'my-ext');
  assert.equal(sources[0]?.lang, 'all');
});

test('listExtensionSources drops descriptors with no usable id', async () => {
  const client = fakeClient(async () => [{ name: 'No Id' }, { id: '', name: 'Empty' }]);
  const sources = await listExtensionSources(client, [
    { file: '/x/a.apk', fallbackName: 'a', sha256: 'hash-a' },
  ]);
  assert.deepEqual(sources, []);
});

test('toInstalledExtensionViews names an extension after the sources it provides', () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/multi.apk', fallbackName: 'multi', sha256: 'hash-a' },
  ];
  const sources: ExtensionSource[] = [
    { id: 'multi:1', bridgeId: '1', name: 'One', lang: 'en', pkg: 'multi', file: '/x/multi.apk' },
    { id: 'multi:2', bridgeId: '2', name: 'Two', lang: 'ja', pkg: 'multi', file: '/x/multi.apk' },
  ];

  assert.deepEqual(toInstalledExtensionViews(extensions, sources, []), [
    { pkg: 'multi', name: 'One, Two', langs: ['en', 'ja'], sourceCount: 2, error: null },
  ]);
});

test('toInstalledExtensionViews lists an extension that loaded nothing, with its reason', () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/broken.apk', fallbackName: 'broken', sha256: 'hash-a' },
  ];

  // A broken APK is still installed, so it must stay listed and removable.
  assert.deepEqual(
    toInstalledExtensionViews(extensions, [], [{ pkg: 'broken', error: 'dex2jar failed' }]),
    [{ pkg: 'broken', name: 'broken', langs: [], sourceCount: 0, error: 'dex2jar failed' }],
  );
});

test('one broken extension does not hide the working ones', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/broken.apk', fallbackName: 'broken', sha256: 'hash-a' },
    { file: '/x/good.apk', fallbackName: 'good', sha256: 'hash-b' },
  ];
  const failures: string[] = [];
  const client = fakeClient(async (source) => {
    if (source.fingerprint === 'hash-a') throw new Error('dex2jar failed');
    return [{ id: '7', name: 'Good Source', lang: 'en' }];
  });

  const sources = await listExtensionSources(client, extensions, (extension) => {
    failures.push(extension.fallbackName);
  });

  assert.deepEqual(failures, ['broken']);
  assert.equal(sources.length, 1);
  assert.equal(sources[0]?.name, 'Good Source');
});
