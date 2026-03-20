import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import { createCharacterDictionaryAutoSyncRuntimeService } from './character-dictionary-auto-sync';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-char-dict-auto-sync-'));
}

function createDeferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
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
  fs.writeFileSync(path.join(dictionariesDir, 'merged.zip'), 'cached-zip', 'utf8');
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
  await Promise.resolve();

  assert.equal(
    events.some((event) => event.phase === 'building'),
    true,
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
  await Promise.resolve();
  await Promise.resolve();

  assert.deepEqual(calls, ['build', 'wait']);

  gate.resolve();
  await syncPromise;

  assert.deepEqual(calls, ['build', 'wait', 'info', 'import', 'settings']);
});
