import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { createCharacterDictionaryAutoSyncRuntimeService } from './character-dictionary-auto-sync';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-char-dict-auto-sync-'));
}

test('auto sync imports merged dictionary and persists MRU state', async () => {
  const userDataPath = makeTempDir();
  const imported: string[] = [];
  const deleted: string[] = [];
  const upserts: Array<{ title: string; scope: 'all' | 'active' }> = [];
  const mergedBuilds: number[][] = [];
  const logs: string[] = [];

  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 130298,
      mediaTitle: 'The Eminence in Shadow',
      entryCount: 2544,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async (mediaIds) => {
      mergedBuilds.push([...mediaIds]);
      return {
        zipPath: '/tmp/subminer-character-dictionary.zip',
        revision: 'rev-1',
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 2544,
      };
    },
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async (zipPath) => {
      imported.push(zipPath);
      importedRevision = 'rev-1';
      return true;
    },
    deleteYomitanDictionary: async (dictionaryTitle) => {
      deleted.push(dictionaryTitle);
      importedRevision = null;
      return true;
    },
    upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
      upserts.push({ title: dictionaryTitle, scope: profileScope });
      return true;
    },
    now: () => 1000,
    logInfo: (message) => {
      logs.push(message);
    },
  });

  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[130298]]);
  assert.deepEqual(imported, ['/tmp/subminer-character-dictionary.zip']);
  assert.deepEqual(deleted, []);
  assert.deepEqual(upserts, [{ title: 'SubMiner Character Dictionary', scope: 'all' }]);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: number[];
    mergedRevision: string | null;
    mergedDictionaryTitle: string | null;
  };
  assert.deepEqual(state.activeMediaIds, [130298]);
  assert.equal(state.mergedRevision, 'rev-1');
  assert.equal(state.mergedDictionaryTitle, 'SubMiner Character Dictionary');
  assert.deepEqual(logs, [
    '[dictionary:auto-sync] syncing current anime snapshot',
    '[dictionary:auto-sync] active AniList media set: 130298',
    '[dictionary:auto-sync] rebuilding merged dictionary for active anime set',
    '[dictionary:auto-sync] importing merged dictionary: /tmp/subminer-character-dictionary.zip',
    '[dictionary:auto-sync] applying Yomitan settings for SubMiner Character Dictionary',
    '[dictionary:auto-sync] synced AniList 130298: SubMiner Character Dictionary (2544 entries)',
  ]);
});

test('auto sync skips rebuild/import on unchanged revisit when merged dictionary is current', async () => {
  const userDataPath = makeTempDir();
  const mergedBuilds: number[][] = [];
  const imports: string[] = [];
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 7,
      mediaTitle: 'Frieren',
      entryCount: 100,
      fromCache: true,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async (mediaIds) => {
      mergedBuilds.push([...mediaIds]);
      return {
        zipPath: '/tmp/merged.zip',
        revision: 'rev-7',
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 100,
      };
    },
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async (zipPath) => {
      imports.push(zipPath);
      importedRevision = 'rev-7';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
  });

  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[7]]);
  assert.deepEqual(imports, ['/tmp/merged.zip']);
});

test('auto sync rebuilds merged dictionary when MRU order changes', async () => {
  const userDataPath = makeTempDir();
  const sequence = [1, 2, 1];
  const mergedBuilds: number[][] = [];
  const deleted: string[] = [];
  let importedRevision: string | null = null;
  let runIndex = 0;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => {
      const mediaId = sequence[Math.min(runIndex, sequence.length - 1)]!;
      runIndex += 1;
      return {
        mediaId,
        mediaTitle: `Title ${mediaId}`,
        entryCount: 10,
        fromCache: true,
        updatedAt: mediaId,
      };
    },
    buildMergedDictionary: async (mediaIds) => {
      mergedBuilds.push([...mediaIds]);
      const revision = `rev-${mediaIds.join('-')}`;
      return {
        zipPath: `/tmp/${revision}.zip`,
        revision,
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: mediaIds.length * 10,
      };
    },
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async (zipPath) => {
      importedRevision = path.basename(zipPath, '.zip');
      return true;
    },
    deleteYomitanDictionary: async (dictionaryTitle) => {
      deleted.push(dictionaryTitle);
      importedRevision = null;
      return true;
    },
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
  });

  await runtime.runSyncNow();
  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[1], [2, 1], [1, 2]]);
  assert.ok(deleted.length >= 2);
});

test('auto sync evicts least recently used media from merged set', async () => {
  const userDataPath = makeTempDir();
  const sequence = [1, 2, 3, 4];
  const mergedBuilds: number[][] = [];
  let runIndex = 0;
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => {
      const mediaId = sequence[Math.min(runIndex, sequence.length - 1)]!;
      runIndex += 1;
      return {
        mediaId,
        mediaTitle: `Title ${mediaId}`,
        entryCount: 10,
        fromCache: true,
        updatedAt: mediaId,
      };
    },
    buildMergedDictionary: async (mediaIds) => {
      mergedBuilds.push([...mediaIds]);
      const revision = `rev-${mediaIds.join('-')}`;
      return {
        zipPath: `/tmp/${revision}.zip`,
        revision,
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: mediaIds.length * 10,
      };
    },
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async (zipPath) => {
      importedRevision = path.basename(zipPath, '.zip');
      return true;
    },
    deleteYomitanDictionary: async () => {
      importedRevision = null;
      return true;
    },
    upsertYomitanDictionarySettings: async () => true,
    now: () => Date.now(),
  });

  await runtime.runSyncNow();
  await runtime.runSyncNow();
  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[1], [2, 1], [3, 2, 1], [4, 3, 2]]);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: number[];
  };
  assert.deepEqual(state.activeMediaIds, [4, 3, 2]);
});
