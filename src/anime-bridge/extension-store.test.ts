import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile } from 'node:fs/promises';
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
  impl: (source: { apkBase64: string }) => Promise<unknown[]>,
): AnimeBridgeClient {
  return { listAnimeSources: impl } as unknown as AnimeBridgeClient;
}

test('readInstalledExtensions reads apks and base64-encodes them', async () => {
  const dir = await makeExtensionDir({ 'my-source.apk': 'APK-BYTES' });
  const extensions = await readInstalledExtensions(dir);

  assert.equal(extensions.length, 1);
  assert.equal(extensions[0]?.fallbackName, 'my-source');
  assert.equal(Buffer.from(extensions[0]!.apkBase64, 'base64').toString(), 'APK-BYTES');
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
  const extension: InstalledExtension = {
    file: '/x/a.apk',
    fallbackName: 'a',
    apkBase64: 'QQ==',
  };
  assert.deepEqual(toBridgeSource(extension), { apkBase64: 'QQ==' });
  assert.deepEqual(toBridgeSource(extension, 'src-1'), {
    apkBase64: 'QQ==',
    sourceId: 'src-1',
  });
});

test('listExtensionSources flattens every source a factory apk provides', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/multi.apk', fallbackName: 'multi', apkBase64: 'QQ==' },
  ];
  const client = fakeClient(async () => [
    { id: 101, name: 'Source One', lang: 'en' },
    { id: '102', name: 'Source Two', lang: 'ja' },
  ]);

  const sources = await listExtensionSources(client, extensions);

  assert.equal(sources.length, 2);
  // Numeric ids are normalized to strings so they can key UI state.
  assert.equal(sources[0]?.id, '101');
  assert.equal(sources[0]?.name, 'Source One');
  assert.equal(sources[1]?.lang, 'ja');
});

test('listExtensionSources falls back to the file name and a default language', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/my-ext.apk', fallbackName: 'my-ext', apkBase64: 'QQ==' },
  ];
  const client = fakeClient(async () => [{ id: '1', name: '   ' }]);

  const sources = await listExtensionSources(client, extensions);
  assert.equal(sources[0]?.name, 'my-ext');
  assert.equal(sources[0]?.lang, 'all');
});

test('listExtensionSources drops descriptors with no usable id', async () => {
  const client = fakeClient(async () => [{ name: 'No Id' }, { id: '', name: 'Empty' }]);
  const sources = await listExtensionSources(client, [
    { file: '/x/a.apk', fallbackName: 'a', apkBase64: 'QQ==' },
  ]);
  assert.deepEqual(sources, []);
});

test('toInstalledExtensionViews names an extension after the sources it provides', () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/multi.apk', fallbackName: 'multi', apkBase64: 'QQ==' },
  ];
  const sources: ExtensionSource[] = [
    { id: '1', name: 'One', lang: 'en', pkg: 'multi', file: '/x/multi.apk' },
    { id: '2', name: 'Two', lang: 'ja', pkg: 'multi', file: '/x/multi.apk' },
  ];

  assert.deepEqual(toInstalledExtensionViews(extensions, sources, []), [
    { pkg: 'multi', name: 'One, Two', langs: ['en', 'ja'], sourceCount: 2, error: null },
  ]);
});

test('toInstalledExtensionViews lists an extension that loaded nothing, with its reason', () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/broken.apk', fallbackName: 'broken', apkBase64: 'QQ==' },
  ];

  // A broken APK is still installed, so it must stay listed and removable.
  assert.deepEqual(
    toInstalledExtensionViews(extensions, [], [{ pkg: 'broken', error: 'dex2jar failed' }]),
    [{ pkg: 'broken', name: 'broken', langs: [], sourceCount: 0, error: 'dex2jar failed' }],
  );
});

test('one broken extension does not hide the working ones', async () => {
  const extensions: InstalledExtension[] = [
    { file: '/x/broken.apk', fallbackName: 'broken', apkBase64: 'QQ==' },
    { file: '/x/good.apk', fallbackName: 'good', apkBase64: 'Qg==' },
  ];
  const failures: string[] = [];
  const client = fakeClient(async (source) => {
    if (source.apkBase64 === 'QQ==') throw new Error('dex2jar failed');
    return [{ id: '7', name: 'Good Source', lang: 'en' }];
  });

  const sources = await listExtensionSources(client, extensions, (extension) => {
    failures.push(extension.fallbackName);
  });

  assert.deepEqual(failures, ['broken']);
  assert.equal(sources.length, 1);
  assert.equal(sources[0]?.name, 'Good Source');
});
