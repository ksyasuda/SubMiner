import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { buildDictionaryZip } from '../character-dictionary-runtime/zip';
import {
  createCharacterDictionaryAutoSyncRuntimeService,
  getCharacterDictionaryManagerSnapshot,
  moveCharacterDictionaryManagedEntry,
  removeCharacterDictionaryManagedEntry,
} from './character-dictionary-auto-sync';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-char-dict-auto-sync-'));
}

async function waitUntil(predicate: () => boolean, label: string): Promise<void> {
  for (let attempt = 0; attempt < 500; attempt += 1) {
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error(`Timed out waiting for ${label}`);
}

function createDeferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

test('character dictionary manager snapshots, reorders, and removes MRU entries', () => {
  const userDataPath = makeTempDir();
  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  fs.mkdirSync(path.dirname(statePath), { recursive: true });
  fs.writeFileSync(
    statePath,
    JSON.stringify(
      {
        activeMediaIds: ['21202 - KonoSuba', '115230 - Tower of God', '130298 - Eminence'],
        mergedRevision: 'rev-1',
        mergedDictionaryTitle: 'SubMiner Character Dictionary',
      },
      null,
      2,
    ),
    'utf8',
  );

  assert.deepEqual(getCharacterDictionaryManagerSnapshot(userDataPath).entries, [
    { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: true },
    { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: false },
    { mediaId: 130298, label: '130298 - Eminence', title: 'Eminence', current: false },
  ]);

  assert.deepEqual(moveCharacterDictionaryManagedEntry(userDataPath, 130298, -1), {
    ok: true,
    entries: [
      { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: true },
      { mediaId: 130298, label: '130298 - Eminence', title: 'Eminence', current: false },
      { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: false },
    ],
    rebuildRequired: true,
  });
  const reorderedState = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    mergedRevision: string | null;
  };
  assert.equal(reorderedState.mergedRevision, null);

  assert.deepEqual(removeCharacterDictionaryManagedEntry(userDataPath, 115230), {
    ok: true,
    entries: [
      { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: true },
      { mediaId: 130298, label: '130298 - Eminence', title: 'Eminence', current: false },
    ],
    rebuildRequired: true,
  });
});

test('character dictionary manager protects the actual current media after LRU reorder', () => {
  const userDataPath = makeTempDir();
  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  fs.mkdirSync(path.dirname(statePath), { recursive: true });
  fs.writeFileSync(
    statePath,
    JSON.stringify(
      {
        activeMediaIds: ['21202 - KonoSuba', '115230 - Tower of God'],
        mergedRevision: 'rev-1',
        mergedDictionaryTitle: 'SubMiner Character Dictionary',
      },
      null,
      2,
    ),
    'utf8',
  );

  assert.deepEqual(getCharacterDictionaryManagerSnapshot(userDataPath, 115230).entries, [
    { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: false },
    { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: true },
  ]);
  assert.deepEqual(moveCharacterDictionaryManagedEntry(userDataPath, 115230, -1, 115230), {
    ok: false,
    message: 'The current anime stays anchored while you are watching it.',
    entries: [
      { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: false },
      { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: true },
    ],
  });
  assert.deepEqual(removeCharacterDictionaryManagedEntry(userDataPath, 115230, 115230), {
    ok: false,
    message: 'The current anime stays loaded while you are watching it.',
    entries: [
      { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba', current: false },
      { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God', current: true },
    ],
  });
});

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
    activeMediaIds: string[];
    mergedRevision: string | null;
    mergedDictionaryTitle: string | null;
  };
  assert.deepEqual(state.activeMediaIds, ['130298 - The Eminence in Shadow']);
  assert.equal(state.mergedRevision, 'rev-1');
  assert.equal(state.mergedDictionaryTitle, 'SubMiner Character Dictionary');
  assert.deepEqual(logs, [
    '[dictionary:auto-sync] syncing current anime snapshot',
    '[dictionary:auto-sync] active AniList media set: 130298 - The Eminence in Shadow',
    '[dictionary:auto-sync] rebuilding merged dictionary for active anime set',
    '[dictionary:auto-sync] importing merged dictionary: /tmp/subminer-character-dictionary.zip (timeout 120000ms)',
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

test('auto sync does not emit updating progress for unchanged revisit when merged dictionary is current', async () => {
  const userDataPath = makeTempDir();
  let importedRevision: string | null = null;
  let currentRun: string[] = [];
  const phaseHistory: string[][] = [];

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
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-7',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 100,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-7';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => false,
    now: () => 1000,
    onSyncStatus: (event) => {
      currentRun.push(event.phase);
    },
  });

  currentRun = [];
  await runtime.runSyncNow();
  phaseHistory.push([...currentRun]);
  currentRun = [];
  await runtime.runSyncNow();
  phaseHistory.push([...currentRun]);

  assert.deepEqual(phaseHistory[0], ['building', 'importing', 'ready']);
  assert.deepEqual(phaseHistory[1], ['ready']);
});

test('auto sync updates MRU order without rebuilding merged dictionary when membership is unchanged', async () => {
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

  assert.deepEqual(mergedBuilds, [[1], [2, 1]]);
  assert.equal(deleted.length, 1);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: string[];
  };
  assert.deepEqual(state.activeMediaIds, ['1 - Title 1', '2 - Title 2']);
});

test('auto sync reimports existing merged zip without rebuilding on unchanged revisit', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });
  buildDictionaryZip(
    path.join(dictionariesDir, 'merged.zip'),
    'SubMiner Character Dictionary',
    'Character names',
    'rev-7',
    [{ term: 'フリーレン', reading: 'フリーレン', role: 'main', glossary: [] } as never],
    [],
  );
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
  importedRevision = null;
  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[7]]);
  assert.deepEqual(imports, [
    '/tmp/merged.zip',
    path.join(userDataPath, 'character-dictionaries', 'merged.zip'),
  ]);
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
    activeMediaIds: string[];
  };
  assert.deepEqual(state.activeMediaIds, ['4 - Title 4', '3 - Title 3', '2 - Title 2']);
});

test('auto sync keeps revisited media retained when a new title is added afterward', async () => {
  const userDataPath = makeTempDir();
  const sequence = [1, 2, 3, 1, 4, 1];
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
  await runtime.runSyncNow();
  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[1], [2, 1], [3, 2, 1], [4, 1, 3]]);

  const statePath = path.join(userDataPath, 'character-dictionaries', 'auto-sync-state.json');
  const state = JSON.parse(fs.readFileSync(statePath, 'utf8')) as {
    activeMediaIds: string[];
  };
  assert.deepEqual(state.activeMediaIds, ['1 - Title 1', '4 - Title 4', '3 - Title 3']);
});

test('auto sync removes stale manual-selection media ids when applying corrected snapshot', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });
  fs.writeFileSync(
    path.join(dictionariesDir, 'auto-sync-state.json'),
    JSON.stringify(
      {
        activeMediaIds: ['10607 - Rerere no Tensai Bakabon', '130298 - The Eminence in Shadow'],
        mergedRevision: 'old',
        mergedDictionaryTitle: 'SubMiner Character Dictionary',
      },
      null,
      2,
    ),
  );
  const builtMediaIds: number[][] = [];
  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 5,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 21355,
      mediaTitle: 'Re:ZERO -Starting Life in Another World-',
      entryCount: 120,
      fromCache: false,
      updatedAt: 99,
      staleMediaIds: [10607],
    }),
    buildMergedDictionary: async (mediaIds) => {
      builtMediaIds.push([...mediaIds]);
      return {
        zipPath: path.join(dictionariesDir, 'merged.zip'),
        revision: `rev-${mediaIds.join('-')}`,
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 200,
      };
    },
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => true,
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => false,
    now: () => 1,
  });

  await runtime.runSyncNow();

  assert.deepEqual(builtMediaIds, [[21355, 130298]]);
  const state = JSON.parse(
    fs.readFileSync(path.join(dictionariesDir, 'auto-sync-state.json'), 'utf8'),
  ) as { activeMediaIds: string[] };
  assert.deepEqual(state.activeMediaIds, [
    '21355 - Re:ZERO -Starting Life in Another World-',
    '130298 - The Eminence in Shadow',
  ]);
});

test('auto sync persists rebuilt MRU state even if Yomitan import fails afterward', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });
  fs.writeFileSync(
    path.join(dictionariesDir, 'auto-sync-state.json'),
    JSON.stringify(
      {
        activeMediaIds: [2, 3, 4],
        mergedRevision: 'rev-2-3-4',
        mergedDictionaryTitle: 'SubMiner Character Dictionary',
      },
      null,
      2,
    ),
  );

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 1,
      mediaTitle: 'Title 1',
      entryCount: 10,
      fromCache: true,
      updatedAt: 1,
    }),
    buildMergedDictionary: async (mediaIds) => {
      assert.deepEqual(mediaIds, [1, 2, 3]);
      return {
        zipPath: '/tmp/rev-1-2-3.zip',
        revision: 'rev-1-2-3',
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 30,
      };
    },
    waitForYomitanMutationReady: async () => undefined,
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => {
      throw new Error('import failed');
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
  });

  await assert.rejects(runtime.runSyncNow(), /import failed/);

  const state = JSON.parse(
    fs.readFileSync(path.join(dictionariesDir, 'auto-sync-state.json'), 'utf8'),
  ) as {
    activeMediaIds: string[];
    mergedRevision: string | null;
    mergedDictionaryTitle: string | null;
  };
  assert.deepEqual(state.activeMediaIds, ['1 - Title 1', '2', '3']);
  assert.equal(state.mergedRevision, 'rev-1-2-3');
  assert.equal(state.mergedDictionaryTitle, 'SubMiner Character Dictionary');
});

test('auto sync invokes completion callback after successful sync', async () => {
  const userDataPath = makeTempDir();
  const completions: Array<{ mediaId: number; mediaTitle: string; changed: boolean }> = [];
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      entryCount: 2560,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-101291',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 2560,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-101291';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    onSyncComplete: (completion) => {
      completions.push(completion);
    },
  });

  await runtime.runSyncNow();

  assert.deepEqual(completions, [
    {
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      changed: true,
    },
  ]);
});

test('auto sync emits progress events for start import and completion', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{
    phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
    mediaId?: number;
    mediaTitle?: string;
    message: string;
    changed?: boolean;
  }> = [];
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onChecking?.({
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      });
      progress?.onGenerating?.({
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      });
      return {
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
        entryCount: 2560,
        fromCache: false,
        updatedAt: 1000,
      };
    },
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-101291',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 2560,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-101291';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    onSyncStatus: (event) => {
      events.push(event);
    },
  });

  await runtime.runSyncNow();

  assert.deepEqual(events, [
    {
      phase: 'checking',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Checking character dictionary for Rascal Does Not Dream of Bunny Girl Senpai...',
    },
    {
      phase: 'generating',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Generating character dictionary for Rascal Does Not Dream of Bunny Girl Senpai...',
    },
    {
      phase: 'building',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Building character dictionary for Rascal Does Not Dream of Bunny Girl Senpai...',
    },
    {
      phase: 'importing',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Importing character dictionary for Rascal Does Not Dream of Bunny Girl Senpai...',
    },
    {
      phase: 'ready',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Character dictionary ready for Rascal Does Not Dream of Bunny Girl Senpai',
      changed: true,
    },
  ]);
});

test('auto sync emits checking before snapshot resolves and skips generating on cache hit', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{
    phase: 'checking' | 'generating' | 'syncing' | 'building' | 'importing' | 'ready' | 'failed';
    mediaId?: number;
    mediaTitle?: string;
    message: string;
    changed?: boolean;
  }> = [];
  const snapshotDeferred = createDeferred<{
    mediaId: number;
    mediaTitle: string;
    entryCount: number;
    fromCache: boolean;
    updatedAt: number;
  }>();
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onChecking?.({
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      });
      return await snapshotDeferred.promise;
    },
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-101291',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 2560,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-101291';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    onSyncStatus: (event) => {
      events.push(event);
    },
  });

  const syncPromise = runtime.runSyncNow();
  await Promise.resolve();

  assert.deepEqual(events, [
    {
      phase: 'checking',
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      message: 'Checking character dictionary for Rascal Does Not Dream of Bunny Girl Senpai...',
    },
  ]);

  snapshotDeferred.resolve({
    mediaId: 101291,
    mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
    entryCount: 2560,
    fromCache: true,
    updatedAt: 1000,
  });
  await syncPromise;

  assert.equal(
    events.some((event) => event.phase === 'generating'),
    false,
  );
});

test('auto sync emits building while merged dictionary generation is in flight', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{
    phase: 'checking' | 'generating' | 'building' | 'syncing' | 'importing' | 'ready' | 'failed';
    mediaId?: number;
    mediaTitle?: string;
    message: string;
    changed?: boolean;
  }> = [];
  const buildDeferred = createDeferred<{
    zipPath: string;
    revision: string;
    dictionaryTitle: string;
    entryCount: number;
  }>();
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onChecking?.({
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      });
      return {
        mediaId: 101291,
        mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
        entryCount: 2560,
        fromCache: true,
        updatedAt: 1000,
      };
    },
    buildMergedDictionary: async () => await buildDeferred.promise,
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-101291';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    onSyncStatus: (event) => {
      events.push(event);
    },
  });

  const syncPromise = runtime.runSyncNow();
  await waitUntil(
    () => events.some((event) => event.phase === 'building'),
    'the building status event',
  );

  buildDeferred.resolve({
    zipPath: '/tmp/merged.zip',
    revision: 'rev-101291',
    dictionaryTitle: 'SubMiner Character Dictionary',
    entryCount: 2560,
  });
  await syncPromise;
});

test('auto sync waits for tokenization-ready gate before Yomitan mutations', async () => {
  const userDataPath = makeTempDir();
  const gate = (() => {
    let resolve!: () => void;
    const promise = new Promise<void>((nextResolve) => {
      resolve = nextResolve;
    });
    return { promise, resolve };
  })();
  const calls: string[] = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({
      enabled: true,
      maxLoaded: 3,
      profileScope: 'all',
    }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 101291,
      mediaTitle: 'Rascal Does Not Dream of Bunny Girl Senpai',
      entryCount: 2560,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async () => {
      calls.push('build');
      return {
        zipPath: '/tmp/merged.zip',
        revision: 'rev-101291',
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 2560,
      };
    },
    waitForYomitanMutationReady: async () => {
      calls.push('wait');
      await gate.promise;
    },
    getYomitanDictionaryInfo: async () => {
      calls.push('info');
      return [];
    },
    importYomitanDictionary: async () => {
      calls.push('import');
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => {
      calls.push('settings');
      return true;
    },
    now: () => 1000,
  });

  const syncPromise = runtime.runSyncNow();
  await waitUntil(() => calls.includes('wait'), 'the tokenization-ready gate');

  assert.deepEqual(calls, ['build', 'wait']);

  gate.resolve();
  await syncPromise;

  assert.deepEqual(calls, ['build', 'wait', 'info', 'import', 'settings']);
});

test('auto sync scales the import timeout with the merged dictionary size', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });
  const zipPath = path.join(dictionariesDir, 'merged.zip');
  // 2 MB of merged dictionary buys ~12s of import budget on top of the base.
  fs.writeFileSync(zipPath, Buffer.alloc(2 * 1024 * 1024));
  const events: Array<{ phase: string; message: string }> = [];
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 21,
      mediaTitle: 'ONE PIECE',
      entryCount: 4000,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async () => ({
      zipPath,
      revision: 'rev-21',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 4000,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      // Far longer than the quick-operation budget, well inside the size-scaled one.
      await new Promise((resolve) => setTimeout(resolve, 400));
      importedRevision = 'rev-21';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    // Comfortable for the stubs that resolve immediately, still far under the import's 400ms.
    operationTimeoutMs: 100,
    dictionaryImportTimeoutBaseMs: 20,
    onSyncStatus: (event) => {
      events.push({ phase: event.phase, message: event.message });
    },
  });

  await runtime.runSyncNow();

  assert.equal(
    events.some((event) => event.phase === 'failed'),
    false,
  );
  assert.deepEqual(events.at(-1), {
    phase: 'ready',
    message: 'Character dictionary ready for ONE PIECE',
  });
});

test('auto sync reports the scaled budget when an import really does hang', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{ phase: string; message: string }> = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 21,
      mediaTitle: 'ONE PIECE',
      entryCount: 4000,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async () => ({
      zipPath: path.join(userDataPath, 'character-dictionaries', 'missing.zip'),
      revision: 'rev-21',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 4000,
    }),
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: () => new Promise<boolean>(() => {}),
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
    dictionaryImportTimeoutBaseMs: 20,
    onSyncStatus: (event) => {
      events.push({ phase: event.phase, message: event.message });
    },
  });

  await assert.rejects(
    runtime.runSyncNow(),
    /importYomitanDictionary\(missing\.zip\) timed out after 20ms/,
  );
  assert.equal(events.at(-1)?.phase, 'failed');
});

test('auto sync ticks the importing notification while the import runs', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{ phase: string; message: string }> = [];
  const scheduled: Array<() => void> = [];
  const importDeferred = createDeferred<boolean>();
  let clock = 1000;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async () => ({
      mediaId: 21,
      mediaTitle: 'ONE PIECE',
      entryCount: 4000,
      fromCache: false,
      updatedAt: 1000,
    }),
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-21',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 4000,
    }),
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: () => importDeferred.promise,
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => clock,
    schedule: (fn) => {
      scheduled.push(fn);
      return 0 as unknown as ReturnType<typeof setTimeout>;
    },
    clearSchedule: () => undefined,
    onSyncStatus: (event) => {
      events.push({ phase: event.phase, message: event.message });
    },
  });

  const syncPromise = runtime.runSyncNow();
  await waitUntil(
    () => events.some((event) => event.phase === 'importing'),
    'the importing status event',
  );

  clock = 1000 + 65_000;
  // The importing heartbeat is the most recently scheduled tick.
  scheduled.at(-1)!();

  assert.deepEqual(events.at(-1), {
    phase: 'importing',
    message: 'Importing character dictionary for ONE PIECE (1m 05s)...',
  });

  importDeferred.resolve(true);
  await syncPromise;
  assert.equal(events.at(-1)?.phase, 'ready');
});

test('auto sync reports character and image counts while generating a snapshot', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{ phase: string; message: string }> = [];
  let clock = 1000;
  let importedRevision: string | null = null;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onGenerating?.({ mediaId: 21, mediaTitle: 'ONE PIECE' });
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'characters',
        completed: 50,
        total: null,
        page: 12,
      });
      // Same stage, same clock tick: throttled away so a 33-page fetch cannot spam the overlay.
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'characters',
        completed: 100,
        total: null,
        page: 13,
      });
      // A stage change always reports, throttle window or not.
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'images',
        completed: 1,
        total: 1220,
      });
      clock += 2000;
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'images',
        completed: 240,
        total: 1220,
      });
      clock += 6000;
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'names',
        completed: 800,
        total: 1220,
      });
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'saving',
        completed: 0,
        total: null,
      });
      return {
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        entryCount: 4000,
        fromCache: false,
        updatedAt: 1000,
      };
    },
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-21',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 4000,
    }),
    getYomitanDictionaryInfo: async () =>
      importedRevision
        ? [{ title: 'SubMiner Character Dictionary', revision: importedRevision }]
        : [],
    importYomitanDictionary: async () => {
      importedRevision = 'rev-21';
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => clock,
    onSyncStatus: (event) => {
      events.push({ phase: event.phase, message: event.message });
    },
  });

  await runtime.runSyncNow();

  assert.deepEqual(
    events.filter((event) => event.phase === 'generating').map((event) => event.message),
    [
      'Generating character dictionary for ONE PIECE...',
      'Generating character dictionary for ONE PIECE (page 12, 50 characters)...',
      'Generating character dictionary for ONE PIECE (image 1/1220)...',
      'Generating character dictionary for ONE PIECE (image 240/1220, ~10s left)...',
      'Generating character dictionary for ONE PIECE (name 800/1220 · 8s)...',
      'Generating character dictionary for ONE PIECE (saving snapshot · 8s)...',
    ],
  );
});

test('auto sync keeps the generating clock ticking when a stage stalls', async () => {
  const userDataPath = makeTempDir();
  const events: Array<{ phase: string; message: string }> = [];
  const scheduled: Array<() => void> = [];
  const snapshotDeferred = createDeferred<{
    mediaId: number;
    mediaTitle: string;
    entryCount: number;
    fromCache: boolean;
    updatedAt: number;
  }>();
  let clock = 1000;

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onGenerating?.({ mediaId: 21, mediaTitle: 'ONE PIECE' });
      progress?.onGenerateProgress?.({
        mediaId: 21,
        mediaTitle: 'ONE PIECE',
        stage: 'images',
        completed: 240,
        total: 1220,
      });
      return await snapshotDeferred.promise;
    },
    buildMergedDictionary: async () => ({
      zipPath: '/tmp/merged.zip',
      revision: 'rev-21',
      dictionaryTitle: 'SubMiner Character Dictionary',
      entryCount: 4000,
    }),
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async () => true,
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => clock,
    schedule: (fn) => {
      scheduled.push(fn);
      return 0 as unknown as ReturnType<typeof setTimeout>;
    },
    clearSchedule: () => undefined,
    onSyncStatus: (event) => {
      events.push({ phase: event.phase, message: event.message });
    },
  });

  const syncPromise = runtime.runSyncNow();
  await waitUntil(() => scheduled.length > 0, 'the generating heartbeat');

  // No further progress arrives: only the clock moves.
  clock += 95_000;
  scheduled.at(-1)!();

  assert.deepEqual(events.at(-1), {
    phase: 'generating',
    message: 'Generating character dictionary for ONE PIECE (image 240/1220 · 1m 35s)...',
  });

  snapshotDeferred.resolve({
    mediaId: 21,
    mediaTitle: 'ONE PIECE',
    entryCount: 4000,
    fromCache: false,
    updatedAt: 1000,
  });
  await syncPromise;
  assert.equal(events.at(-1)?.phase, 'ready');
});

test('auto sync rebuilds instead of importing a cached merged ZIP with a mismatched revision', async () => {
  const userDataPath = makeTempDir();
  const dictionariesDir = path.join(userDataPath, 'character-dictionaries');
  fs.mkdirSync(dictionariesDir, { recursive: true });
  const statePath = path.join(dictionariesDir, 'auto-sync-state.json');
  fs.writeFileSync(
    statePath,
    JSON.stringify({
      activeMediaIds: ['7 - Frieren'],
      mergedRevision: 'rev-7',
      mergedDictionaryTitle: 'SubMiner Character Dictionary',
    }),
    'utf8',
  );
  // Left over from an interrupted run: the archive on disk is not the revision state recorded.
  buildDictionaryZip(
    path.join(dictionariesDir, 'merged.zip'),
    'SubMiner Character Dictionary',
    'Character names',
    'rev-stale',
    [{ term: 'フリーレン', reading: 'フリーレン', role: 'main', glossary: [] } as never],
    [],
  );
  const mergedBuilds: number[][] = [];
  const imports: string[] = [];

  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
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
        zipPath: '/tmp/rebuilt-merged.zip',
        revision: 'rev-7',
        dictionaryTitle: 'SubMiner Character Dictionary',
        entryCount: 100,
      };
    },
    // Yomitan does not have the dictionary, so the sync has to import despite the cached state.
    getYomitanDictionaryInfo: async () => [],
    importYomitanDictionary: async (zipPath) => {
      imports.push(zipPath);
      return true;
    },
    deleteYomitanDictionary: async () => true,
    upsertYomitanDictionarySettings: async () => true,
    now: () => 1000,
  });

  await runtime.runSyncNow();

  assert.deepEqual(mergedBuilds, [[7]]);
  assert.deepEqual(imports, ['/tmp/rebuilt-merged.zip']);
});
