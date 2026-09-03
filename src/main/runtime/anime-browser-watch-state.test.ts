import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { buildAnimeStreamStatsPath } from '../../anime-bridge/episode-metadata';
import { createAnimeBrowserRuntime } from './anime-browser-runtime';
import type { AnimeBrowserRuntimeDeps } from './anime-browser-runtime-deps';
import type { AnimeBridgeClient } from '../../anime-bridge/bridge-client';

async function setupRuntime(overrides: Partial<AnimeBrowserRuntimeDeps> = {}) {
  const dir = await mkdtemp(path.join(tmpdir(), 'subminer-anime-watch-'));
  await writeFile(path.join(dir, 'pkg.one.apk'), 'one');
  const client = {
    listAnimeSources: async () => [{ id: 'shared', name: 'One', lang: 'en' }],
  };
  const runtime = createAnimeBrowserRuntime({
    extensionsDir: () => dir,
    repos: () => [],
    setRepos: () => undefined,
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
    ...overrides,
  });
  await runtime.ensureBridge();
  return runtime;
}

test('getWatchState asks the stats store for the derived per-episode paths', async () => {
  let asked: string[] = [];
  const watched = buildAnimeStreamStatsPath('pkg.one:shared', '/anime/1', '/ep/1');
  const runtime = await setupRuntime({
    getWatchState: async (statsPaths) => {
      asked = statsPaths;
      return new Map([[watched, { watched: true, lastWatchedMs: 42, sessionCount: 2 }]]);
    },
  });

  const state = await runtime.getWatchState({
    sourceId: 'pkg.one:shared',
    animeUrl: '/anime/1',
    // The empty url is what a source with a malformed entry hands over.
    episodeUrls: ['/ep/1', '/ep/2', ''],
  });

  assert.deepEqual(asked, [
    watched,
    buildAnimeStreamStatsPath('pkg.one:shared', '/anime/1', '/ep/2'),
  ]);
  // Episodes with no history are absent rather than reported unwatched.
  assert.deepEqual(state, [
    { episodeUrl: '/ep/1', watched: true, lastWatchedMs: 42, sessionCount: 2 },
  ]);
  await runtime.dispose();
});

test('setWatched records the series metadata a never-played episode needs', async () => {
  const store = new Map<string, { watched: boolean; lastWatchedMs: number | null }>();
  let marked: Array<Record<string, unknown>> = [];
  const runtime = await setupRuntime({
    getWatchState: async (statsPaths) =>
      new Map(
        statsPaths
          .filter((statsPath) => store.has(statsPath))
          .map((statsPath) => [statsPath, { ...store.get(statsPath)!, sessionCount: 0 }]),
      ),
    setWatchState: async (episodes, watched) => {
      marked = episodes as unknown as Array<Record<string, unknown>>;
      for (const episode of episodes) {
        store.set(episode.statsPath, { watched, lastWatchedMs: null });
      }
      return episodes.length;
    },
  });

  const state = await runtime.setWatched({
    sourceId: 'pkg.one:shared',
    animeUrl: '/anime/1',
    animeTitle: 'Test Series Season 3',
    watched: true,
    episodes: [
      { episodeUrl: '/ep/4', episodeName: 'Episode 4 - Homecoming', episodeNumber: 4 },
      { episodeUrl: '/ep/3', episodeName: 'Episode 3', episodeNumber: null },
      { episodeUrl: '', episodeName: 'Broken', episodeNumber: null },
    ],
  });

  assert.equal(marked.length, 2, 'the entry with no url is dropped before the write');
  assert.deepEqual(marked[0], {
    mediaPath: '',
    statsPath: buildAnimeStreamStatsPath('pkg.one:shared', '/anime/1', '/ep/4'),
    displayTitle: 'Test Series S03E04 - Homecoming',
    seriesTitle: 'Test Series',
    seasonNumber: 3,
    episodeNumber: 4,
  });
  // The number is read off the label when the source reported none.
  assert.equal(marked[1]?.episodeNumber, 3);
  assert.deepEqual(
    state.map((entry) => entry.episodeUrl),
    ['/ep/4', '/ep/3'],
  );
  assert.ok(state.every((entry) => entry.watched));
  await runtime.dispose();
});

test('setWatched clears marks and reports the state the write left behind', async () => {
  const store = new Map([
    [
      buildAnimeStreamStatsPath('pkg.one:shared', '/anime/1', '/ep/4'),
      { watched: true, lastWatchedMs: 10, sessionCount: 1 },
    ],
  ]);
  const runtime = await setupRuntime({
    getWatchState: async (statsPaths) =>
      new Map(
        statsPaths
          .filter((statsPath) => store.has(statsPath))
          .map((statsPath) => [statsPath, store.get(statsPath)!]),
      ),
    setWatchState: async (episodes, watched) => {
      for (const episode of episodes) {
        const current = store.get(episode.statsPath);
        if (current) store.set(episode.statsPath, { ...current, watched });
      }
      return episodes.length;
    },
  });

  const state = await runtime.setWatched({
    sourceId: 'pkg.one:shared',
    animeUrl: '/anime/1',
    animeTitle: 'Test Series',
    watched: false,
    episodes: [{ episodeUrl: '/ep/4', episodeName: 'Episode 4', episodeNumber: 4 }],
  });

  assert.deepEqual(state, [
    { episodeUrl: '/ep/4', watched: false, lastWatchedMs: 10, sessionCount: 1 },
  ]);
  await runtime.dispose();
});

test('setWatched reports nothing when stats tracking supplies no writer', async () => {
  const runtime = await setupRuntime();
  assert.deepEqual(
    await runtime.setWatched({
      sourceId: 'pkg.one:shared',
      animeUrl: '/anime/1',
      animeTitle: 'Test Series',
      watched: true,
      episodes: [{ episodeUrl: '/ep/1', episodeName: 'Episode 1', episodeNumber: 1 }],
    }),
    [],
  );
  await runtime.dispose();
});

test('getWatchState returns nothing when stats tracking supplies no lookup', async () => {
  const runtime = await setupRuntime();
  assert.deepEqual(
    await runtime.getWatchState({
      sourceId: 'pkg.one:shared',
      animeUrl: '/anime/1',
      episodeUrls: ['/ep/1'],
    }),
    [],
  );
  await runtime.dispose();
});

test('a failing stats lookup leaves the episode list usable', async () => {
  const logged: string[] = [];
  const runtime = await setupRuntime({
    log: (message) => logged.push(message),
    getWatchState: async () => {
      throw new Error('database is locked');
    },
  });

  assert.deepEqual(
    await runtime.getWatchState({
      sourceId: 'pkg.one:shared',
      animeUrl: '/anime/1',
      episodeUrls: ['/ep/1'],
    }),
    [],
  );
  assert.ok(logged.some((message) => message.includes('database is locked')));
  await runtime.dispose();
});
