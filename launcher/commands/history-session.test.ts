import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import {
  buildHistoryEntryActions,
  buildHistorySessionActions,
  runHistoryPlaybackLoop,
} from './history-command.js';
import type { HistorySeriesEntry } from '../history.js';

type HistoryLoop = (
  initial: { entry: HistorySeriesEntry; videoPath: string },
  deps: {
    play: (videoPath: string) => Promise<void>;
    pickPostPlaybackAction: (input: {
      entry: HistorySeriesEntry;
      justPlayedPath: string;
      previousEpisodePath: string | null;
      nextEpisodePath: string | null;
    }) => Promise<'previous' | 'replay' | 'next' | 'browse' | 'quit' | null>;
    findPreviousEpisode: (videoPath: string) => string | null;
    findNextEpisode: (videoPath: string) => string | null;
    browseEpisodes: (entry: HistorySeriesEntry) => Promise<string | null>;
  },
) => Promise<void>;

const typedRunHistoryPlaybackLoop: HistoryLoop = runHistoryPlaybackLoop;

function makeEntry(lastWatchedPath: string): HistorySeriesEntry {
  return {
    seriesRoot: path.dirname(lastWatchedPath),
    displayName: 'Test Show',
    coverBlobHash: null,
    lastWatched: {
      videoId: 1,
      sourcePath: lastWatchedPath,
      parsedTitle: 'Test Show',
      parsedSeason: 1,
      parsedEpisode: 1,
      animeTitle: 'Test Show',
      lastWatchedMs: 1,
      coverBlobHash: null,
    },
  };
}

test('history loop plays an initial selection, browsed selection, then quits', async () => {
  assert.equal(
    typeof runHistoryPlaybackLoop,
    'function',
    'history playback loop is not implemented',
  );
  const entry = makeEntry('/shows/test-show/episode-01.mkv');
  const played: string[] = [];
  const menuEntries: HistorySeriesEntry[] = [];
  let menuCount = 0;

  await typedRunHistoryPlaybackLoop(
    { entry, videoPath: '/shows/test-show/episode-02.mkv' },
    {
      play: async (videoPath) => {
        played.push(videoPath);
      },
      pickPostPlaybackAction: async ({ entry: menuEntry }) => {
        menuEntries.push(menuEntry);
        return menuCount++ === 0 ? 'browse' : 'quit';
      },
      findPreviousEpisode: () => null,
      findNextEpisode: () => null,
      browseEpisodes: async (browseEntry) => {
        assert.equal(browseEntry, entry);
        return '/shows/test-show/episode-04.mkv';
      },
    },
  );

  assert.deepEqual(played, ['/shows/test-show/episode-02.mkv', '/shows/test-show/episode-04.mkv']);
  assert.deepEqual(menuEntries, [entry, entry]);
});

test('history replay uses the actual just-played path instead of the database row', async () => {
  assert.equal(
    typeof runHistoryPlaybackLoop,
    'function',
    'history playback loop is not implemented',
  );
  const entry = makeEntry('/shows/test-show/stale-episode-01.mkv');
  const played: string[] = [];
  const menuPaths: string[] = [];
  let menuCount = 0;

  await typedRunHistoryPlaybackLoop(
    { entry, videoPath: '/shows/test-show/episode-07.mkv' },
    {
      play: async (videoPath) => {
        played.push(videoPath);
      },
      pickPostPlaybackAction: async ({ justPlayedPath }) => {
        menuPaths.push(justPlayedPath);
        return menuCount++ === 0 ? 'replay' : 'quit';
      },
      findPreviousEpisode: () => null,
      findNextEpisode: () => '/shows/test-show/episode-08.mkv',
      browseEpisodes: async () => null,
    },
  );

  assert.deepEqual(played, ['/shows/test-show/episode-07.mkv', '/shows/test-show/episode-07.mkv']);
  assert.deepEqual(menuPaths, [
    '/shows/test-show/episode-07.mkv',
    '/shows/test-show/episode-07.mkv',
  ]);
});

test('history next is computed from the actual just-played path', async () => {
  assert.equal(
    typeof runHistoryPlaybackLoop,
    'function',
    'history playback loop is not implemented',
  );
  const entry = makeEntry('/shows/test-show/stale-episode-01.mkv');
  const played: string[] = [];
  const nextInputs: string[] = [];
  let menuCount = 0;

  await typedRunHistoryPlaybackLoop(
    { entry, videoPath: '/shows/test-show/episode-07.mkv' },
    {
      play: async (videoPath) => {
        played.push(videoPath);
      },
      pickPostPlaybackAction: async ({ nextEpisodePath }) => {
        if (menuCount++ === 0) {
          assert.equal(nextEpisodePath, '/shows/test-show/episode-08.mkv');
          return 'next';
        }
        assert.equal(nextEpisodePath, null);
        return 'quit';
      },
      findPreviousEpisode: () => null,
      findNextEpisode: (videoPath) => {
        nextInputs.push(videoPath);
        return videoPath.endsWith('episode-07.mkv') ? '/shows/test-show/episode-08.mkv' : null;
      },
      browseEpisodes: async () => null,
    },
  );

  assert.deepEqual(played, ['/shows/test-show/episode-07.mkv', '/shows/test-show/episode-08.mkv']);
  assert.deepEqual(nextInputs, [
    '/shows/test-show/episode-07.mkv',
    '/shows/test-show/episode-08.mkv',
  ]);
});

test('history show menu offers previous, rewatch, next, browse, and quit in order', () => {
  assert.equal(
    typeof buildHistorySessionActions,
    'function',
    'history session actions are not implemented',
  );

  assert.deepEqual(
    buildHistorySessionActions(
      '/shows/test-show/episode-07.mkv',
      '/shows/test-show/episode-06.mkv',
      '/shows/test-show/episode-08.mkv',
    ),
    [
      { kind: 'previous', label: 'Previous episode: episode-06.mkv' },
      { kind: 'replay', label: 'Rewatch episode: episode-07.mkv' },
      { kind: 'next', label: 'Play next episode: episode-08.mkv' },
      { kind: 'browse', label: 'Select / browse episode' },
      { kind: 'quit', label: 'Quit SubMiner' },
    ],
  );
});

test('history show menu omits previous and next when the just-played episode has neither', () => {
  assert.equal(
    typeof buildHistorySessionActions,
    'function',
    'history session actions are not implemented',
  );

  assert.deepEqual(buildHistorySessionActions('/shows/test-show/finale.mkv', null, null), [
    { kind: 'replay', label: 'Rewatch episode: finale.mkv' },
    { kind: 'browse', label: 'Select / browse episode' },
    { kind: 'quit', label: 'Quit SubMiner' },
  ]);
});

test('history entry menu offers previous, replay, and next before playback starts', () => {
  assert.equal(
    typeof buildHistoryEntryActions,
    'function',
    'history entry actions not implemented',
  );

  assert.deepEqual(
    buildHistoryEntryActions(
      '/shows/test-show/episode-03.mkv',
      '/shows/test-show/episode-02.mkv',
      '/shows/test-show/episode-04.mkv',
    ),
    [
      { kind: 'previous', label: 'Previous episode: episode-02.mkv' },
      { kind: 'replay', label: 'Replay last watched: episode-03.mkv' },
      { kind: 'next', label: 'Next episode: episode-04.mkv' },
      { kind: 'browse', label: 'Browse episodes' },
      { kind: 'quit', label: 'Quit SubMiner' },
    ],
  );
});

test('history entry menu omits replay when the last watched file is gone', () => {
  assert.deepEqual(buildHistoryEntryActions(null, null, '/shows/test-show/episode-04.mkv'), [
    { kind: 'next', label: 'Next episode: episode-04.mkv' },
    { kind: 'browse', label: 'Browse episodes' },
    { kind: 'quit', label: 'Quit SubMiner' },
  ]);
});

test('history playback loop selects previous based on the just-played path, then re-derives previous from the new current episode', async () => {
  assert.equal(
    typeof runHistoryPlaybackLoop,
    'function',
    'history playback loop is not implemented',
  );
  const entry = makeEntry('/shows/test-show/stale-episode-09.mkv');
  const played: string[] = [];
  const previousInputs: string[] = [];
  const previousSeenByMenu: Array<string | null> = [];
  let menuCount = 0;

  await typedRunHistoryPlaybackLoop(
    { entry, videoPath: '/shows/test-show/episode-07.mkv' },
    {
      play: async (videoPath) => {
        played.push(videoPath);
      },
      pickPostPlaybackAction: async ({ previousEpisodePath }) => {
        previousSeenByMenu.push(previousEpisodePath);
        return menuCount++ === 0 ? 'previous' : 'quit';
      },
      findPreviousEpisode: (videoPath) => {
        previousInputs.push(videoPath);
        if (videoPath.endsWith('episode-07.mkv')) return '/shows/test-show/episode-06.mkv';
        if (videoPath.endsWith('episode-06.mkv')) return '/shows/test-show/episode-05.mkv';
        return null;
      },
      findNextEpisode: () => null,
      browseEpisodes: async () => null,
    },
  );

  assert.deepEqual(played, ['/shows/test-show/episode-07.mkv', '/shows/test-show/episode-06.mkv']);
  assert.deepEqual(previousInputs, [
    '/shows/test-show/episode-07.mkv',
    '/shows/test-show/episode-06.mkv',
  ]);
  assert.deepEqual(previousSeenByMenu, [
    '/shows/test-show/episode-06.mkv',
    '/shows/test-show/episode-05.mkv',
  ]);
});

test('history playback loop stops advancing when previous is chosen with no prior episode', async () => {
  assert.equal(
    typeof runHistoryPlaybackLoop,
    'function',
    'history playback loop is not implemented',
  );
  const entry = makeEntry('/shows/test-show/episode-01.mkv');
  const played: string[] = [];

  await typedRunHistoryPlaybackLoop(
    { entry, videoPath: '/shows/test-show/episode-01.mkv' },
    {
      play: async (videoPath) => {
        played.push(videoPath);
      },
      pickPostPlaybackAction: async ({ previousEpisodePath }) => {
        assert.equal(previousEpisodePath, null);
        return 'previous';
      },
      findPreviousEpisode: () => null,
      findNextEpisode: () => null,
      browseEpisodes: async () => null,
    },
  );

  assert.deepEqual(played, ['/shows/test-show/episode-01.mkv']);
});
