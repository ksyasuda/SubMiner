import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createJellyfinEpisodePlaylist,
  selectEpisodeWindow,
  type JellyfinEpisodePlaylistSettings,
  type JellyfinPlaylistEntry,
} from './jellyfin-episode-playlist';

const EPISODES = ['e1', 'e2', 'e3', 'e4', 'e5', 'e6'];

test('selectEpisodeWindow centers on the current episode and stays full at series edges', () => {
  assert.deepEqual(selectEpisodeWindow(EPISODES, 'e3', 3), {
    previousIds: ['e2'],
    nextIds: ['e4'],
  });
  assert.deepEqual(selectEpisodeWindow(EPISODES, 'e3', 4), {
    previousIds: ['e2'],
    nextIds: ['e4', 'e5'],
  });
  assert.deepEqual(selectEpisodeWindow(EPISODES, 'e1', 3), {
    previousIds: [],
    nextIds: ['e2', 'e3'],
  });
  assert.deepEqual(selectEpisodeWindow(EPISODES, 'e6', 3), {
    previousIds: ['e4', 'e5'],
    nextIds: [],
  });
  assert.deepEqual(selectEpisodeWindow(['e1', 'e2'], 'e1', 5), {
    previousIds: [],
    nextIds: ['e2'],
  });
  assert.deepEqual(selectEpisodeWindow(EPISODES, 'missing', 3), { previousIds: [], nextIds: [] });
});

type Entry = JellyfinPlaylistEntry;

const url = (itemId: string) => `https://jf.example/Videos/${itemId}/stream`;
const entry = (itemId: string): Entry => ({
  itemId,
  seriesId: 'series',
  playbackUrl: url(itemId),
  loadfileOptions: `opts-${itemId}`,
  title: `Show ${itemId}`,
});

/** mpv's playlist, kept in sync with the commands the queue sends. */
function makeHarness(
  settings: JellyfinEpisodePlaylistSettings = { enabled: true, maxEpisodes: 3 },
) {
  const playlist: Array<{ filename: string; current: boolean; title?: string }> = [];
  const commands: Array<Array<string | number>> = [];
  let episodeListCalls = 0;
  const queue = createJellyfinEpisodePlaylist<null, Entry>({
    getSettings: () => settings,
    listSeriesEpisodeIds: async () => {
      episodeListCalls += 1;
      return EPISODES;
    },
    prepareEntry: async (_context, itemId) => entry(itemId),
    readMpvPlaylist: async () => playlist.map((item) => ({ ...item })),
    sendMpvCommand: (command) => {
      commands.push(command);
      if (command[0] === 'playlist-remove') playlist.splice(Number(command[1]), 1);
      if (command[0] === 'loadfile') {
        playlist.splice(Number(command[3]), 0, { filename: String(command[1]), current: false });
      }
      const titleIndex = /^playlist\/(\d+)\/title$/.exec(String(command[1]))?.[1];
      if (command[0] === 'set_property' && titleIndex !== undefined) {
        playlist[Number(titleIndex)]!.title = String(command[2]);
      }
    },
    logDebug: () => {},
  });
  const playing = (itemId: string) => {
    for (const item of playlist) item.current = item.filename === url(itemId);
    if (!playlist.some((item) => item.current)) {
      playlist.splice(0, playlist.length, { filename: url(itemId), current: true });
    }
  };
  return {
    queue,
    commands,
    playlist,
    playing,
    episodeListCalls: () => episodeListCalls,
    filenames: () => playlist.map((item) => item.filename),
    titles: () => playlist.map((item) => item.title),
  };
}

test('queueNeighbors inserts the episode window around the current entry', async () => {
  const harness = makeHarness();
  harness.playing('e3');

  await harness.queue.queueNeighbors(null, entry('e3'));

  assert.deepEqual(harness.filenames(), [url('e2'), url('e3'), url('e4')]);
  assert.deepEqual(
    harness.commands.filter((command) => command[0] === 'loadfile'),
    [
      ['loadfile', url('e4'), 'insert-at', 1, 'opts-e4'],
      ['loadfile', url('e2'), 'insert-at', 0, 'opts-e2'],
    ],
  );
  assert.deepEqual(harness.commands.at(-1), [
    'script-message',
    'subminer-managed-playlist-entries',
    url('e3'),
    url('e4'),
    url('e2'),
  ]);
  assert.deepEqual(harness.titles(), ['Show e2', 'Show e3', 'Show e4']);
  assert.equal(harness.queue.find(url('e4'))?.entry.itemId, 'e4');
  assert.equal(harness.queue.find(url('e5')), null);
});

test('queueNeighbors slides the window when mpv moves to a queued episode', async () => {
  const harness = makeHarness();
  harness.playing('e3');
  await harness.queue.queueNeighbors(null, entry('e3'));

  harness.playing('e4');
  await harness.queue.queueNeighbors(null, entry('e4'));

  assert.deepEqual(harness.filenames(), [url('e3'), url('e4'), url('e5')]);
  assert.deepEqual(harness.titles(), ['Show e3', 'Show e4', 'Show e5']);
  assert.equal(harness.queue.find(url('e2')), null);
  assert.equal(harness.episodeListCalls(), 1, 'series order is reused while navigating');
});

test('queueNeighbors leaves the playlist alone once mpv has moved past the item', async () => {
  const harness = makeHarness();
  harness.playing('e3');
  await harness.queue.queueNeighbors(null, entry('e3'));
  harness.playing('e4');
  const before = harness.filenames();

  await harness.queue.queueNeighbors(null, entry('e3'));

  assert.deepEqual(harness.filenames(), before);
});

test('queueNeighbors drops queued entries and adds none when the playlist is disabled', async () => {
  const settings = { enabled: true, maxEpisodes: 3 };
  const harness = makeHarness(settings);
  harness.playing('e3');
  await harness.queue.queueNeighbors(null, entry('e3'));

  settings.enabled = false;
  await harness.queue.queueNeighbors(null, entry('e3'));

  assert.deepEqual(harness.filenames(), [url('e3')]);
});

test('reset cancels a queue run that is still resolving episodes', async () => {
  const harness = makeHarness();
  harness.playing('e3');
  const pending = harness.queue.queueNeighbors(null, entry('e3'));
  harness.queue.reset();
  await pending;

  assert.deepEqual(harness.filenames(), [url('e3')]);
  assert.equal(harness.queue.find(url('e3')), null);
});
