import assert from 'node:assert/strict';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import test from 'node:test';
import type { AnilistCharacterDictionaryProfileScope } from '../../types';
import type {
  CharacterDictionarySnapshotResult,
  MergedCharacterDictionaryBuildResult,
} from '../character-dictionary-runtime';
import { buildDictionaryZip } from '../character-dictionary-runtime/zip';
import {
  createCharacterDictionaryAutoSyncRuntimeService,
  getCharacterDictionaryManagerSnapshot,
  moveCharacterDictionaryManagedEntry,
  removeCharacterDictionaryManagedEntry,
  type CharacterDictionaryAutoSyncRuntimeDeps,
  type CharacterDictionaryAutoSyncStatusEvent,
} from './character-dictionary-auto-sync';

const DICTIONARY_TITLE = 'SubMiner Character Dictionary';
const FRIEREN = { mediaId: 7, mediaTitle: 'Frieren' };
const ONE_PIECE = { mediaId: 21, mediaTitle: 'ONE PIECE' };
const DEFAULT_SNAPSHOT: CharacterDictionarySnapshotResult = {
  ...FRIEREN,
  entryCount: 100,
  fromCache: true,
  updatedAt: 1000,
};

type PersistedState = {
  activeMediaIds: string[];
  mergedRevision: string | null;
  mergedDictionaryTitle: string | null;
};

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-char-dict-auto-sync-'));
}

function dictionariesDir(userDataPath: string): string {
  return path.join(userDataPath, 'character-dictionaries');
}

function statePath(userDataPath: string): string {
  return path.join(dictionariesDir(userDataPath), 'auto-sync-state.json');
}

function writeState(
  userDataPath: string,
  state: Omit<PersistedState, 'activeMediaIds'> & { activeMediaIds: Array<string | number> },
): void {
  fs.mkdirSync(dictionariesDir(userDataPath), { recursive: true });
  fs.writeFileSync(statePath(userDataPath), JSON.stringify(state, null, 2), 'utf8');
}

function readState(userDataPath: string): PersistedState {
  return JSON.parse(fs.readFileSync(statePath(userDataPath), 'utf8')) as PersistedState;
}

/** Writes a real `merged.zip` whose index.json carries `revision`. */
async function writeMergedZip(userDataPath: string, revision: string): Promise<string> {
  const zipPath = path.join(dictionariesDir(userDataPath), 'merged.zip');
  await buildDictionaryZip(zipPath, DICTIONARY_TITLE, 'Character names', revision, [], []);
  return zipPath;
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

/** Captures heartbeat ticks so a test can fire them by hand instead of waiting on real timers. */
function manualScheduler(): {
  scheduled: Array<() => void>;
  deps: Pick<CharacterDictionaryAutoSyncRuntimeDeps, 'schedule' | 'clearSchedule'>;
} {
  const scheduled: Array<() => void> = [];
  return {
    scheduled,
    deps: {
      schedule: (fn) => {
        scheduled.push(fn);
        return 0 as unknown as ReturnType<typeof setTimeout>;
      },
      clearSchedule: () => undefined,
    },
  };
}

type RuntimeOptions = Partial<CharacterDictionaryAutoSyncRuntimeDeps> & {
  snapshot?: Partial<CharacterDictionarySnapshotResult>;
  /** What the settings upsert reports; `false` means Yomitan settings were already in place. */
  settingsChanged?: boolean;
};

/**
 * Builds the runtime against an in-memory Yomitan that remembers the last imported revision.
 * Default merged builds are named `/tmp/rev-<ids>.zip` with revision `rev-<ids>`, and an import
 * records the ZIP basename as the installed revision.
 */
function makeRuntime({ snapshot, settingsChanged = true, ...overrides }: RuntimeOptions = {}) {
  const userDataPath = overrides.userDataPath ?? makeTempDir();
  const recorded = {
    userDataPath,
    builds: [] as number[][],
    imports: [] as string[],
    deletes: [] as string[],
    upserts: [] as Array<{ title: string; scope: AnilistCharacterDictionaryProfileScope }>,
    events: [] as CharacterDictionaryAutoSyncStatusEvent[],
    completions: [] as Array<{ mediaId: number; mediaTitle: string; changed: boolean }>,
    logs: [] as string[],
    yomitan: { revision: null as string | null, infoQueries: 0 },
  };
  const runtime = createCharacterDictionaryAutoSyncRuntimeService({
    userDataPath,
    getConfig: () => ({ enabled: true, maxLoaded: 3, profileScope: 'all' }),
    getOrCreateCurrentSnapshot: async () => ({ ...DEFAULT_SNAPSHOT, ...snapshot }),
    buildMergedDictionary: async (mediaIds): Promise<MergedCharacterDictionaryBuildResult> => {
      recorded.builds.push([...mediaIds]);
      const revision = `rev-${mediaIds.join('-')}`;
      return {
        zipPath: `/tmp/${revision}.zip`,
        revision,
        dictionaryTitle: DICTIONARY_TITLE,
        entryCount: mediaIds.length * 10,
      };
    },
    getYomitanDictionaryInfo: async () => {
      recorded.yomitan.infoQueries += 1;
      const { revision } = recorded.yomitan;
      return revision ? [{ title: DICTIONARY_TITLE, revision }] : [];
    },
    importYomitanDictionary: async (zipPath) => {
      recorded.imports.push(zipPath);
      recorded.yomitan.revision = path.basename(zipPath, '.zip');
      return true;
    },
    deleteYomitanDictionary: async (title) => {
      recorded.deletes.push(title);
      recorded.yomitan.revision = null;
      return true;
    },
    upsertYomitanDictionarySettings: async (title, scope) => {
      recorded.upserts.push({ title, scope });
      return settingsChanged;
    },
    now: () => 1000,
    logInfo: (message) => {
      recorded.logs.push(message);
    },
    onSyncStatus: (event) => {
      recorded.events.push(event);
    },
    onSyncComplete: (completion) => {
      recorded.completions.push(completion);
    },
    ...overrides,
  });
  return {
    ...recorded,
    runtime,
    phases: () => recorded.events.map((event) => event.phase),
  };
}

const konosuba = { mediaId: 21202, label: '21202 - KonoSuba', title: 'KonoSuba' };
const towerOfGod = { mediaId: 115230, label: '115230 - Tower of God', title: 'Tower of God' };
const eminence = { mediaId: 130298, label: '130298 - Eminence', title: 'Eminence' };

test('character dictionary manager snapshots, reorders, and removes MRU entries', () => {
  const userDataPath = makeTempDir();
  writeState(userDataPath, {
    activeMediaIds: [konosuba.label, towerOfGod.label, eminence.label],
    mergedRevision: 'rev-1',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });

  assert.deepEqual(getCharacterDictionaryManagerSnapshot(userDataPath).entries, [
    { ...konosuba, current: true },
    { ...towerOfGod, current: false },
    { ...eminence, current: false },
  ]);

  assert.deepEqual(moveCharacterDictionaryManagedEntry(userDataPath, eminence.mediaId, -1), {
    ok: true,
    entries: [
      { ...konosuba, current: true },
      { ...eminence, current: false },
      { ...towerOfGod, current: false },
    ],
    rebuildRequired: true,
  });
  assert.equal(readState(userDataPath).mergedRevision, null);

  assert.deepEqual(removeCharacterDictionaryManagedEntry(userDataPath, towerOfGod.mediaId), {
    ok: true,
    entries: [
      { ...konosuba, current: true },
      { ...eminence, current: false },
    ],
    rebuildRequired: true,
  });
});

test('character dictionary manager protects the actual current media after LRU reorder', () => {
  const userDataPath = makeTempDir();
  writeState(userDataPath, {
    activeMediaIds: [konosuba.label, towerOfGod.label],
    mergedRevision: 'rev-1',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });
  const current = towerOfGod.mediaId;
  const entries = [
    { ...konosuba, current: false },
    { ...towerOfGod, current: true },
  ];

  assert.deepEqual(getCharacterDictionaryManagerSnapshot(userDataPath, current).entries, entries);
  assert.deepEqual(moveCharacterDictionaryManagedEntry(userDataPath, current, -1, current), {
    ok: false,
    message: 'The current anime stays anchored while you are watching it.',
    entries,
  });
  assert.deepEqual(removeCharacterDictionaryManagedEntry(userDataPath, current, current), {
    ok: false,
    message: 'The current anime stays loaded while you are watching it.',
    entries,
  });
});

test('auto sync imports merged dictionary, persists MRU state, and reports completion', async () => {
  const r = makeRuntime({
    snapshot: {
      mediaId: 130298,
      mediaTitle: 'The Eminence in Shadow',
      entryCount: 2544,
      fromCache: false,
    },
  });

  await r.runtime.runSyncNow();

  assert.deepEqual(r.builds, [[130298]]);
  assert.deepEqual(r.imports, ['/tmp/rev-130298.zip']);
  assert.deepEqual(r.deletes, []);
  assert.deepEqual(r.upserts, [{ title: DICTIONARY_TITLE, scope: 'all' }]);
  assert.deepEqual(readState(r.userDataPath), {
    activeMediaIds: ['130298 - The Eminence in Shadow'],
    mergedRevision: 'rev-130298',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });
  assert.deepEqual(r.logs, [
    '[dictionary:auto-sync] syncing current anime snapshot',
    '[dictionary:auto-sync] active AniList media set: 130298 - The Eminence in Shadow',
    '[dictionary:auto-sync] rebuilding merged dictionary for active anime set',
    '[dictionary:auto-sync] importing merged dictionary: /tmp/rev-130298.zip (timeout 120000ms)',
    `[dictionary:auto-sync] applying Yomitan settings for ${DICTIONARY_TITLE}`,
    `[dictionary:auto-sync] synced AniList 130298: ${DICTIONARY_TITLE} (2544 entries)`,
  ]);
  assert.deepEqual(r.completions, [
    { mediaId: 130298, mediaTitle: 'The Eminence in Shadow', changed: true },
  ]);
});

test('auto sync skips rebuild, import, and progress phases on an unchanged revisit', async () => {
  const r = makeRuntime({ settingsChanged: false });

  await r.runtime.runSyncNow();
  assert.deepEqual(r.phases(), ['building', 'importing', 'ready']);

  r.events.length = 0;
  await r.runtime.runSyncNow();
  assert.deepEqual(r.phases(), ['ready']);
  assert.deepEqual(r.builds, [[7]]);
  assert.deepEqual(r.imports, ['/tmp/rev-7.zip']);
});

const mruCases: Array<{
  name: string;
  sequence: number[];
  builds: number[][];
  activeMediaIds: string[];
}> = [
  {
    name: 'auto sync updates MRU order without rebuilding merged dictionary when membership is unchanged',
    sequence: [1, 2, 1],
    builds: [[1], [2, 1]],
    activeMediaIds: ['1 - Title 1', '2 - Title 2'],
  },
  {
    name: 'auto sync evicts least recently used media from merged set',
    sequence: [1, 2, 3, 4],
    builds: [[1], [2, 1], [3, 2, 1], [4, 3, 2]],
    activeMediaIds: ['4 - Title 4', '3 - Title 3', '2 - Title 2'],
  },
  {
    name: 'auto sync keeps revisited media retained when a new title is added afterward',
    sequence: [1, 2, 3, 1, 4, 1],
    builds: [[1], [2, 1], [3, 2, 1], [4, 1, 3]],
    activeMediaIds: ['1 - Title 1', '4 - Title 4', '3 - Title 3'],
  },
];

for (const c of mruCases) {
  test(c.name, async () => {
    let runIndex = 0;
    const r = makeRuntime({
      getOrCreateCurrentSnapshot: async () => {
        const mediaId = c.sequence[runIndex++]!;
        return {
          mediaId,
          mediaTitle: `Title ${mediaId}`,
          entryCount: 10,
          fromCache: true,
          updatedAt: mediaId,
        };
      },
    });

    for (let run = 0; run < c.sequence.length; run += 1) {
      await r.runtime.runSyncNow();
    }

    assert.deepEqual(r.builds, c.builds);
    // A revisit that only reorders the set neither rebuilds nor reimports.
    assert.deepEqual(
      r.imports,
      c.builds.map((mediaIds) => `/tmp/rev-${mediaIds.join('-')}.zip`),
    );
    assert.deepEqual(readState(r.userDataPath).activeMediaIds, c.activeMediaIds);
  });
}

test('auto sync reimports existing merged zip without rebuilding on unchanged revisit', async () => {
  const r = makeRuntime();
  const cachedZipPath = await writeMergedZip(r.userDataPath, 'rev-7');

  await r.runtime.runSyncNow();
  // Yomitan lost the dictionary between syncs.
  r.yomitan.revision = null;
  await r.runtime.runSyncNow();

  assert.deepEqual(r.builds, [[7]]);
  assert.deepEqual(r.imports, ['/tmp/rev-7.zip', cachedZipPath]);
});

test('auto sync rebuilds instead of importing a cached merged ZIP with a mismatched revision', async () => {
  const r = makeRuntime();
  writeState(r.userDataPath, {
    activeMediaIds: ['7 - Frieren'],
    mergedRevision: 'rev-7',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });
  // Left over from an interrupted run: the archive on disk is not the revision state recorded.
  await writeMergedZip(r.userDataPath, 'rev-stale');

  // Yomitan does not have the dictionary, so the sync has to import despite the cached state.
  await r.runtime.runSyncNow();

  assert.deepEqual(r.builds, [[7]]);
  assert.deepEqual(r.imports, ['/tmp/rev-7.zip']);
});

test('auto sync removes stale manual-selection media ids when applying corrected snapshot', async () => {
  const r = makeRuntime({
    snapshot: {
      mediaId: 21355,
      mediaTitle: 'Re:ZERO -Starting Life in Another World-',
      fromCache: false,
      staleMediaIds: [10607],
    },
  });
  writeState(r.userDataPath, {
    activeMediaIds: ['10607 - Rerere no Tensai Bakabon', '130298 - The Eminence in Shadow'],
    mergedRevision: 'old',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });

  await r.runtime.runSyncNow();

  assert.deepEqual(r.builds, [[21355, 130298]]);
  assert.deepEqual(readState(r.userDataPath).activeMediaIds, [
    '21355 - Re:ZERO -Starting Life in Another World-',
    '130298 - The Eminence in Shadow',
  ]);
});

test('auto sync persists rebuilt MRU state even if Yomitan import fails afterward', async () => {
  const r = makeRuntime({
    snapshot: { mediaId: 1, mediaTitle: 'Title 1' },
    importYomitanDictionary: async () => {
      throw new Error('import failed');
    },
  });
  writeState(r.userDataPath, {
    activeMediaIds: [2, 3, 4],
    mergedRevision: 'rev-2-3-4',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });

  await assert.rejects(r.runtime.runSyncNow(), /import failed/);

  assert.deepEqual(r.builds, [[1, 2, 3]]);
  assert.deepEqual(readState(r.userDataPath), {
    activeMediaIds: ['1 - Title 1', '2', '3'],
    mergedRevision: 'rev-1-2-3',
    mergedDictionaryTitle: DICTIONARY_TITLE,
  });
});

test('auto sync emits progress events for start import and completion', async () => {
  const r = makeRuntime({
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onChecking?.(FRIEREN);
      progress?.onGenerating?.(FRIEREN);
      return { ...DEFAULT_SNAPSHOT, fromCache: false };
    },
  });

  await r.runtime.runSyncNow();

  assert.deepEqual(r.events, [
    { phase: 'checking', ...FRIEREN, message: 'Checking character dictionary for Frieren...' },
    { phase: 'generating', ...FRIEREN, message: 'Generating character dictionary for Frieren...' },
    { phase: 'building', ...FRIEREN, message: 'Building character dictionary for Frieren...' },
    { phase: 'importing', ...FRIEREN, message: 'Importing character dictionary for Frieren...' },
    {
      phase: 'ready',
      ...FRIEREN,
      message: 'Character dictionary ready for Frieren',
      changed: true,
    },
  ]);
});

test('auto sync emits checking before snapshot resolves and skips generating on cache hit', async () => {
  const snapshot = createDeferred<CharacterDictionarySnapshotResult>();
  const r = makeRuntime({
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onChecking?.(FRIEREN);
      return await snapshot.promise;
    },
  });

  const syncPromise = r.runtime.runSyncNow();
  await Promise.resolve();

  assert.deepEqual(r.events, [
    { phase: 'checking', ...FRIEREN, message: 'Checking character dictionary for Frieren...' },
  ]);

  snapshot.resolve(DEFAULT_SNAPSHOT);
  await syncPromise;

  assert.equal(r.phases().includes('generating'), false);
});

test('auto sync emits building while merged dictionary generation is in flight', async () => {
  const build = createDeferred<MergedCharacterDictionaryBuildResult>();
  const r = makeRuntime({ buildMergedDictionary: () => build.promise });

  const syncPromise = r.runtime.runSyncNow();
  await waitUntil(() => r.phases().includes('building'), 'the building status event');

  build.resolve({
    zipPath: '/tmp/rev-7.zip',
    revision: 'rev-7',
    dictionaryTitle: DICTIONARY_TITLE,
    entryCount: 100,
  });
  await syncPromise;
});

test('auto sync waits for tokenization-ready gate before Yomitan mutations', async () => {
  const gate = createDeferred<void>();
  let gateReached = false;
  const r = makeRuntime({
    waitForYomitanMutationReady: async () => {
      gateReached = true;
      await gate.promise;
    },
  });

  const syncPromise = r.runtime.runSyncNow();
  await waitUntil(() => gateReached, 'the tokenization-ready gate');

  assert.deepEqual(r.builds, [[7]]);
  assert.equal(r.yomitan.infoQueries, 0);
  assert.deepEqual(r.imports, []);
  assert.deepEqual(r.upserts, []);

  gate.resolve();
  await syncPromise;

  assert.equal(r.yomitan.infoQueries, 1);
  assert.deepEqual(r.imports, ['/tmp/rev-7.zip']);
  assert.equal(r.upserts.length, 1);
});

test('auto sync scales the import timeout with the merged dictionary size', async () => {
  const userDataPath = makeTempDir();
  const zipPath = path.join(dictionariesDir(userDataPath), 'merged.zip');
  fs.mkdirSync(path.dirname(zipPath), { recursive: true });
  // 2 MB of merged dictionary buys ~12s of import budget on top of the base.
  fs.writeFileSync(zipPath, Buffer.alloc(2 * 1024 * 1024));

  const r = makeRuntime({
    userDataPath,
    buildMergedDictionary: async () => ({
      zipPath,
      revision: 'rev-7',
      dictionaryTitle: DICTIONARY_TITLE,
      entryCount: 100,
    }),
    importYomitanDictionary: async () => {
      // Far longer than the quick-operation budget, well inside the size-scaled one.
      await new Promise((resolve) => setTimeout(resolve, 400));
      return true;
    },
    // Comfortable for the stubs that resolve immediately, still far under the import's 400ms.
    operationTimeoutMs: 100,
    dictionaryImportTimeoutBaseMs: 20,
  });

  await r.runtime.runSyncNow();

  assert.equal(r.phases().includes('failed'), false);
  assert.equal(r.events.at(-1)?.message, 'Character dictionary ready for Frieren');
});

test('auto sync reports the scaled budget when an import really does hang', async () => {
  const userDataPath = makeTempDir();
  const r = makeRuntime({
    userDataPath,
    buildMergedDictionary: async () => ({
      zipPath: path.join(dictionariesDir(userDataPath), 'missing.zip'),
      revision: 'rev-7',
      dictionaryTitle: DICTIONARY_TITLE,
      entryCount: 100,
    }),
    importYomitanDictionary: () => new Promise<boolean>(() => {}),
    dictionaryImportTimeoutBaseMs: 20,
  });

  await assert.rejects(
    r.runtime.runSyncNow(),
    /importYomitanDictionary\(missing\.zip\) timed out after 20ms/,
  );
  assert.equal(r.events.at(-1)?.phase, 'failed');
});

test('auto sync ticks the importing notification while the import runs', async () => {
  const scheduler = manualScheduler();
  const importResult = createDeferred<boolean>();
  let clock = 1000;
  const r = makeRuntime({
    ...scheduler.deps,
    importYomitanDictionary: () => importResult.promise,
    now: () => clock,
  });

  const syncPromise = r.runtime.runSyncNow();
  await waitUntil(() => r.phases().includes('importing'), 'the importing status event');

  clock += 65_000;
  // The importing heartbeat is the most recently scheduled tick.
  scheduler.scheduled.at(-1)!();

  assert.deepEqual(r.events.at(-1), {
    phase: 'importing',
    ...FRIEREN,
    message: 'Importing character dictionary for Frieren (1m 05s)...',
  });

  importResult.resolve(true);
  await syncPromise;
  assert.equal(r.events.at(-1)?.phase, 'ready');
});

test('auto sync reports character and image counts while generating a snapshot', async () => {
  let clock = 1000;
  const r = makeRuntime({
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onGenerating?.(ONE_PIECE);
      progress?.onGenerateProgress?.({
        ...ONE_PIECE,
        stage: 'characters',
        completed: 50,
        total: null,
        page: 12,
      });
      // Same stage, same clock tick: throttled away so a 33-page fetch cannot spam the overlay.
      progress?.onGenerateProgress?.({
        ...ONE_PIECE,
        stage: 'characters',
        completed: 100,
        total: null,
        page: 13,
      });
      // A stage change always reports, throttle window or not.
      progress?.onGenerateProgress?.({ ...ONE_PIECE, stage: 'images', completed: 1, total: 1220 });
      clock += 2000;
      progress?.onGenerateProgress?.({
        ...ONE_PIECE,
        stage: 'images',
        completed: 240,
        total: 1220,
      });
      clock += 6000;
      progress?.onGenerateProgress?.({ ...ONE_PIECE, stage: 'names', completed: 800, total: 1220 });
      progress?.onGenerateProgress?.({ ...ONE_PIECE, stage: 'saving', completed: 0, total: null });
      return { ...ONE_PIECE, entryCount: 4000, fromCache: false, updatedAt: 1000 };
    },
    now: () => clock,
  });

  await r.runtime.runSyncNow();

  assert.deepEqual(
    r.events.filter((event) => event.phase === 'generating').map((event) => event.message),
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
  const scheduler = manualScheduler();
  const snapshot = createDeferred<CharacterDictionarySnapshotResult>();
  let clock = 1000;
  const r = makeRuntime({
    ...scheduler.deps,
    getOrCreateCurrentSnapshot: async (_targetPath, progress) => {
      progress?.onGenerating?.(ONE_PIECE);
      progress?.onGenerateProgress?.({
        ...ONE_PIECE,
        stage: 'images',
        completed: 240,
        total: 1220,
      });
      return await snapshot.promise;
    },
    now: () => clock,
  });

  const syncPromise = r.runtime.runSyncNow();
  await waitUntil(() => scheduler.scheduled.length > 0, 'the generating heartbeat');

  // No further progress arrives: only the clock moves.
  clock += 95_000;
  scheduler.scheduled.at(-1)!();

  assert.deepEqual(r.events.at(-1), {
    phase: 'generating',
    ...ONE_PIECE,
    message: 'Generating character dictionary for ONE PIECE (image 240/1220 · 1m 35s)...',
  });

  snapshot.resolve({ ...ONE_PIECE, entryCount: 4000, fromCache: false, updatedAt: 1000 });
  await syncPromise;
  assert.equal(r.events.at(-1)?.phase, 'ready');
});
