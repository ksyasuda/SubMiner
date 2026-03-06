import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { createCharacterDictionaryAutoSyncRuntimeService } from './character-dictionary-auto-sync';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-char-dict-auto-sync-'));
}

test('auto sync imports current dictionary and updates persisted state', async () => {
  const userDataPath = makeTempDir();
  const imported: string[] = [];
  const upserts: Array<{ title: string; scope: 'all' | 'active' }> = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      refreshTtlHours: 168,
      maxLoaded: 3,
      evictionPolicy: 'delete',
      profileScope: 'all',
    }),
    generateCharacterDictionary: async () => ({
      zipPath: '/tmp/anilist-130298.zip',
      fromCache: false,
      mediaId: 130298,
      mediaTitle: 'The Eminence in Shadow',
      entryCount: 2544,
      dictionaryTitle: 'SubMiner Character Dictionary (AniList 130298)',
      revision: '100',
    }),
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async (zipPath) => {
      imported.push(zipPath);
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async (dictionaryTitle, profileScope) => {
      upserts.push({ title: dictionaryTitle, scope: profileScope });
      return true;
    },
    removeYomitanDictionarySettings: async () => true,
    now: () => 1000,
  });

  await runtime.runSyncNow();

  assert.deepEqual(imported, ['/tmp/anilist-130298.zip']);
  assert.deepEqual(upserts, [
    { title: 'SubMiner Character Dictionary (AniList 130298)', scope: 'all' },
  ]);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: number[];
    dictionariesByMediaId: Record<string, { lastImportedRevision: string }>;
  };
  assert.deepEqual(state.activeMediaIds, [130298]);
  assert.equal(state.dictionariesByMediaId['130298']?.lastImportedRevision, '100');
});

test('auto sync rotates dictionaries by LRU and deletes overflow when policy=delete', async () => {
  const userDataPath = makeTempDir();
  const generated = [
    { mediaId: 1, zipPath: '/tmp/anilist-1.zip', title: 'SubMiner Character Dictionary (AniList 1)' },
    { mediaId: 2, zipPath: '/tmp/anilist-2.zip', title: 'SubMiner Character Dictionary (AniList 2)' },
    { mediaId: 3, zipPath: '/tmp/anilist-3.zip', title: 'SubMiner Character Dictionary (AniList 3)' },
    { mediaId: 4, zipPath: '/tmp/anilist-4.zip', title: 'SubMiner Character Dictionary (AniList 4)' },
  ];
  let runIndex = 0;
  const deletes: string[] = [];
  const removals: Array<{ title: string; mode: 'delete' | 'disable' }> = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      refreshTtlHours: 168,
      maxLoaded: 3,
      evictionPolicy: 'delete',
      profileScope: 'all',
    }),
    generateCharacterDictionary: async () => {
      const current = generated[Math.min(runIndex, generated.length - 1)]!;
      runIndex += 1;
      return {
        zipPath: current.zipPath,
        fromCache: false,
        mediaId: current.mediaId,
        mediaTitle: `Title ${current.mediaId}`,
        entryCount: 10,
        dictionaryTitle: current.title,
        revision: String(current.mediaId),
      };
    },
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => true,
    deleteYomitanDictionary: async (dictionaryTitle) => {
      deletes.push(dictionaryTitle);
      return true;
    },
    upsertYomitanDictionarySettings: async () => true,
    removeYomitanDictionarySettings: async (dictionaryTitle, _scope, mode) => {
      removals.push({ title: dictionaryTitle, mode });
      return true;
    },
    now: () => Date.now(),
  });

  await runtime.runSyncNow();
  await runtime.runSyncNow();
  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.ok(removals.some((entry) => entry.title.includes('(AniList 1)') && entry.mode === 'delete'));
  assert.ok(deletes.some((title) => title.includes('(AniList 1)')));

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: number[];
    dictionariesByMediaId: Record<string, unknown>;
  };
  assert.deepEqual(state.activeMediaIds, [4, 3, 2]);
  assert.equal(state.dictionariesByMediaId['1'], undefined);
});

test('auto sync disable eviction keeps dictionary in DB and only disables settings', async () => {
  const userDataPath = makeTempDir();
  let runIndex = 0;
  const deletes: string[] = [];
  const removals: Array<{ title: string; mode: 'delete' | 'disable' }> = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      refreshTtlHours: 168,
      maxLoaded: 1,
      evictionPolicy: 'disable',
      profileScope: 'all',
    }),
    generateCharacterDictionary: async () => {
      runIndex += 1;
      return {
        zipPath: `/tmp/anilist-${runIndex}.zip`,
        fromCache: false,
        mediaId: runIndex,
        mediaTitle: `Title ${runIndex}`,
        entryCount: 10,
        dictionaryTitle: `SubMiner Character Dictionary (AniList ${runIndex})`,
        revision: String(runIndex),
      };
    },
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => true,
    deleteYomitanDictionary: async (dictionaryTitle) => {
      deletes.push(dictionaryTitle);
      return true;
    },
    upsertYomitanDictionarySettings: async () => true,
    removeYomitanDictionarySettings: async (dictionaryTitle, _scope, mode) => {
      removals.push({ title: dictionaryTitle, mode });
      return true;
    },
    now: () => Date.now(),
  });

  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.ok(removals.some((entry) => entry.mode === 'disable' && entry.title.includes('(AniList 1)')));
  assert.equal(deletes.some((title) => title.includes('(AniList 1)')), false);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: number[];
    dictionariesByMediaId: Record<string, unknown>;
  };
  assert.deepEqual(state.activeMediaIds, [2]);
  assert.ok(state.dictionariesByMediaId['1']);
  assert.ok(state.dictionariesByMediaId['2']);
});

test('auto sync fails fast when yomitan import hangs', async () => {
  const userDataPath = makeTempDir();

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    operationTimeoutMs: 5,
    getConfig: () => ({
      enabled: true,
      refreshTtlHours: 168,
      maxLoaded: 3,
      evictionPolicy: 'delete',
      profileScope: 'all',
    }),
    generateCharacterDictionary: async () => ({
      zipPath: '/tmp/anilist-130298.zip',
      fromCache: true,
      mediaId: 130298,
      mediaTitle: 'The Eminence in Shadow',
      entryCount: 2544,
      dictionaryTitle: 'SubMiner Character Dictionary (AniList 130298)',
      revision: '100',
    }),
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () =>
      new Promise<boolean>(() => {
        // never resolve
      }),
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    removeYomitanDictionarySettings: async () => true,
    now: () => Date.now(),
  });

  await assert.rejects(async () => runtime.runSyncNow(), /importYomitanDictionary\(anilist-130298\.zip\) timed out after 5ms/);
});
